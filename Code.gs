/**********************************************************************
 * QA Wall Board — Google Apps Script web app
 *
 * All four views (week / month / quarter / year) are rendered into one
 * response and cycled with a CSS animation — no page navigation, so the
 * Apps Script sandbox iframe is never asked to reload itself.
 * Fresh numbers arrive every REFRESH_MINUTES via a google.script.run
 * data pull that repaints the board in place (no navigation at all).
 *
 * No canvas, no SVG, no charting library — plain boxes and background
 * colours, so an old Smart-TV browser renders it like a table.
 *
 * PUBLIC DEPLOYMENT
 *   Deploy > New deployment > Web app
 *     Execute as        : Me (your account)
 *     Who has access    : Anyone
 *   The resulting .../exec URL then opens on any device, any network,
 *   with no Google sign-in. See DEPLOY.md for the full runbook.
 *
 * OPTIONAL QUERY PARAMETERS
 *   ?only=week|month|quarter|year   one board, no rotation
 *   ?rot=15                         seconds per view (default ROTATE_SECONDS)
 *   ?ref=60                         minutes between data pulls
 **********************************************************************/

/* Leave SHEET_ID empty when this script is bound to the spreadsheet
   (Extensions > Apps Script from inside the sheet). Fill it in with the
   long id from the sheet URL if the script is standalone — a public
   deployment that executes as you can still read a private sheet. */
const SHEET_ID = '';

const SHEET_NAME       = 'Form Responses 1';
const ORG_LINE         = 'Opus · Customer Support';
const BOARD_TITLE      = 'Quality Scoreboard';
const ROTATE_SECONDS   = 10;   // how long each view stays on screen
const REFRESH_MINUTES  = 360;  // how often fresh data is pulled (6 h)
const CACHE_SECONDS    = 300;  // how often the sheet is actually re-read

const TARGET_BUDGET  = 0.85;
const TARGET_STRETCH = 0.90;
const SCALE_FLOOR    = 0;      // bars start at zero, so 93.3% draws 93.3% of the track

const ORDER = ['week', 'month', 'quarter', 'year'];
const NAMES = { week:'WEEKLY', month:'MONTHLY', quarter:'QUARTERLY', year:'YEARLY' };
/* The board always shows the most recent period that has data — which is
   usually the current one — so "Latest" is the only wording that stays true. */
const SCORE_LABEL = { week:    'Latest Week QA Score',
                      month:   'Latest Month QA Score',
                      quarter: 'Latest Quarter QA Score',
                      year:    'Latest Year QA Score' };
const TREND_N = { week: 12, month: 12, quarter: 8, year: 5 };

const PLOT_H   = 150;  // height of the plot row inside the trend box
const BAR_MAX  = 128;  // tallest a trend column may be
const PLOT_TOP = 64;   // top of the plot row, measured from the trend box
const MAX_CRIT = 3;    // rows in the Top markdowns table

const LAST_COL = 42;   // widest column the board reads

const COL = { CHANNEL: 28, SCORE: 32, WEEK: 35,
              MD_C: 37, MD_B: 38, MD_K: 39, P_C: 40, P_B: 41, P_K: 42 };

const CRITERIA = [
  ['Opening & Greeting', 2], ['Acknowledgment & Assurance', 3], ['Active Listening', 4],
  ['Proactive Guidance', 5], ['Alignment & Restoring', 6], ['Closure', 7],
  ['Pace, Tone & Volume', 8], ['Professionalism', 9], ['Syntax & Grammar', 10],
  ['Communicate Clearly', 11], ['Hold', 12], ['Readiness', 14],
  ['Escalation', 15], ['Accuracy & Adherence', 16], ['Probing Questions', 17],
  ['Logging', 18], ['Internal Note', 19], ['Data Privacy & Security', 21],
  ['Verification', 22], ['Call recording', 23]
];

const C = {
  bg:'#0E1116', surf:'#161B23', line:'#242B36', line2:'#1B212A', track:'#20262F',
  fg:'#E8EDF4', mut:'#7E8899', dim:'#59626F',
  acc:'#4C8DFF', good:'#3EBD6B', warn:'#E0B33C', bad:'#FF8A73', flat:'#4A5464'
};

/* ------------------------------ helpers ------------------------------ */
function p2_(n){ return n < 10 ? '0' + n : '' + n; }
function pct_(x){ return (x * 100).toFixed(1) + '%'; }
function esc_(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;'); }
function clamp_(n, lo, hi){ return Math.max(lo, Math.min(hi, n)); }
function scale_(s){ return clamp_(Math.round((s - SCALE_FLOOR)/(1 - SCALE_FLOOR)*100), 3, 100); }
/* > Stretch -> green | <= Stretch -> yellow | <= Budget -> light red
   Each Targets row therefore carries exactly the colour its own value earns. */
function tone_(s){ return s > TARGET_STRETCH ? C.good : (s > TARGET_BUDGET ? C.warn : C.bad); }

function monday_(v){
  if (v instanceof Date) return v;
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)){
    const p = v.split('-');
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  return null;
}
function isoWeek_(d){
  const t = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  t.setDate(t.getDate() + 3 - ((t.getDay() + 6) % 7));
  const w1 = new Date(t.getFullYear(), 0, 4);
  return { y: t.getFullYear(),
           w: 1 + Math.round(((t - w1)/86400000 - 3 + ((w1.getDay() + 6) % 7)) / 7) };
}
function key_(d, mode){
  const y = d.getFullYear(), m = d.getMonth() + 1;
  if (mode === 'week')    return y + '-' + p2_(m) + '-' + p2_(d.getDate());
  if (mode === 'quarter') return y + '-Q' + Math.ceil(m / 3);
  if (mode === 'year')    return '' + y;
  return y + '-' + p2_(m);
}
const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function bigLabel_(k, mode, d){
  if (mode === 'month')   { const p = k.split('-'); return MON[+p[1]-1] + ' ' + p[0]; }
  if (mode === 'quarter') { const p = k.split('-'); return p[0] + ' · ' + p[1]; }
  if (mode === 'week' && d){
    const i = isoWeek_(d);
    return 'Week ' + i.w + ' · ' + d.getDate() + ' ' + MON[d.getMonth()] + ' ' + i.y;
  }
  return k;
}
function tick_(k, mode){
  if (mode === 'month')   { const p = k.split('-'); return MON[+p[1]-1] + ' ' + p[0]; }
  if (mode === 'week')    return k.slice(5).replace('-', '/');
  if (mode === 'quarter') return k.split('-')[1] + ' ' + k.slice(2,4);
  return k;
}

/* ------------------------------- data -------------------------------- */
/* The sheet is opened by id when SHEET_ID is set, otherwise by container.
   Either way the read happens as the deployment owner, so anonymous
   viewers never need access to the spreadsheet itself. */
function sheet_(){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
  if (!ss) throw new Error('No spreadsheet available. Bind this script to the sheet, ' +
                           'or set SHEET_ID at the top of the script.');
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Tab "' + SHEET_NAME + '" not found in "' + ss.getName() + '".');
  if (sh.getMaxColumns() < LAST_COL)
    throw new Error('Tab "' + SHEET_NAME + '" has ' + sh.getMaxColumns() +
                    ' columns; the board reads up to column ' + LAST_COL + '.');
  return sh;
}

/* One sheet read per request, memoised, rather than one per view. */
let ROWS_ = null;
function rows_(){
  if (ROWS_) return ROWS_;
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return (ROWS_ = []);
  const v = sh.getRange(1, 1, last, LAST_COL).getValues();
  const all = [];
  for (let i = 1; i < v.length; i++){
    const r = v[i], s = r[COL.SCORE - 1], d = monday_(r[COL.WEEK - 1]);
    if (typeof s === 'number' && d) all.push({ r: r, d: d, s: s });
  }
  return (ROWS_ = all);
}

function build_(mode){
  const all = rows_();
  if (!all.length) return null;

  const bucket = {};
  all.forEach(function(o){
    const k = key_(o.d, mode);
    if (!bucket[k]) bucket[k] = [];
    bucket[k].push(o);
  });
  const keys = Object.keys(bucket).sort();
  const lag = (mode === 'week' || mode === 'month') && keys.length > 1;
  const cur = lag ? keys[keys.length - 2] : keys[keys.length - 1];
  const sel = bucket[cur];

  const md  = function(o){ return (o.r[COL.MD_C-1]||0) + (o.r[COL.MD_B-1]||0) + (o.r[COL.MD_K-1]||0); };
  const avg = function(f){ return sel.reduce(function(a,o){ return a + f(o); }, 0) / sel.length; };

  const curIdx = keys.indexOf(cur);
  const trend = keys.slice(0, curIdx + 1).slice(-TREND_N[mode]).map(function(k){
    const b = bucket[k];
    return { k: tick_(k, mode), s: b.reduce(function(a,o){ return a + o.s; }, 0) / b.length };
  });

  const H = {};
  sel.forEach(function(o){
    const c = String(o.r[COL.CHANNEL-1] || '').toUpperCase().trim(); if (!c) return;
    if (!H[c]) H[c] = { name: c, sum: 0, n: 0 };
    H[c].sum += o.s; H[c].n++;
  });
  const channels = Object.keys(H).map(function(k){ return H[k]; })
                         .sort(function(a,b){ return b.n - a.n; });

  const crit = CRITERIA.map(function(c){
                 return { name: c[0],
                          n: sel.filter(function(o){ return o.r[c[1]-1] === 'Needs Improvement'; }).length };
               }).filter(function(c){ return c.n > 0; })
                 .sort(function(a,b){ return b.n - a.n; })
                 .slice(0, MAX_CRIT);

  return {
    label:   bigLabel_(cur, mode, sel[0].d),
    score:   avg(function(o){ return o.s; }),
    audits:  sel.length,
    md:      sel.reduce(function(a,o){ return a + md(o); }, 0),
    pillars: [['Customer',   avg(function(o){ return o.r[COL.P_C-1] || 0; })],
              ['Business',   avg(function(o){ return o.r[COL.P_B-1] || 0; })],
              ['Compliance', avg(function(o){ return o.r[COL.P_K-1] || 0; })]],
    channels: channels, crit: crit, trend: trend
  };
}

function stats_(mode){
  const cache = CacheService.getScriptCache();
  const hit = cache.get('wb2_' + mode);
  if (hit) return JSON.parse(hit);
  const d = build_(mode);
  if (d) cache.put('wb2_' + mode, JSON.stringify(d), CACHE_SECONDS);
  return d;
}

/* ------------------------------ pieces ------------------------------- */
function barCell_(widthPct, color){
  return '<td class="bw"><div class="track"><div class="fill" style="width:'
       + widthPct + '%;background:' + color + '"></div></div></td>';
}
function kpi_(label, value, color, small, cls){
  return '<td class="kt ' + (cls || '') + '"><div class="kbox" style="border-left-color:'
       + (color || C.line) + '">'
       + '<div class="eb">' + label + '</div>'
       + '<div class="kn' + (small ? ' ksm' : '') + '"'
       + (color ? ' style="color:' + color + '"' : '') + '>' + value + '</div></div></td>';
}

/* the Quality Targets legend that sits beside QA Score */
function targets_(){
  const rows = [['Rockstar', 1,               C.good],
                ['Stretch',  TARGET_STRETCH,  C.warn],
                ['Budget',   TARGET_BUDGET,   C.bad]];
  let t = '';
  rows.forEach(function(r){
    t += '<tr><td class="tgn" style="color:' + r[2] + '">' + r[0] + '</td>'
       + '<td class="tgv" style="color:' + r[2] + '">' + Math.round(r[1] * 100) + '%</td></tr>';
  });
  return '<td class="kt k2"><div class="kbox tgbox">'
       + '<div class="eb">Targets</div><table class="tg">' + t + '</table></div></td>';
}

function view_(mode, stamp, rot){
  const d = stats_(mode);
  if (!d) return '<div class="rail"><div class="ttl">No data</div></div>';

  let band = 'Below Budget', col = C.bad;
  if (d.score >= 1)                    { band = 'Rockstar';      col = C.good; }
  else if (d.score > TARGET_STRETCH)   { band = 'Above Stretch'; col = C.good; }
  else if (d.score > TARGET_BUDGET)    { band = 'Above Budget';  col = C.warn; }

  let h = '';

  h += '<table class="rail"><tr>'
     + '<td class="who"><div class="org">' + esc_(ORG_LINE) + '</div>'
     + '<div class="ttl">' + esc_(BOARD_TITLE) + '</div></td>'
     + '<td class="pl">';
  ORDER.forEach(function(m){
    h += '<span class="pill' + (m === mode ? ' on big' : '') + '">' + NAMES[m] + '</span>';
  });
  h += '</td></tr></table>';

  h += '<table class="kpi"><tr>'
     + kpi_(SCORE_LABEL[mode], pct_(d.score), col, false, 'k1')
     + targets_()
     + kpi_('Against Target', band, col, true, 'k3')
     + kpi_('Audits', d.audits, null, false, 'k4')
     + kpi_('Markdowns', d.md, null, false, 'k5')
     + '</tr></table>';

  const guideTop = PLOT_TOP + PLOT_H
                 - Math.round((TARGET_STRETCH - SCALE_FLOOR) / (1 - SCALE_FLOOR) * BAR_MAX);
  h += '<div class="trend"><div class="eb">Score trend</div>'
     + '<div class="guide" style="top:' + guideTop + 'px"></div>'
     + '<table class="cols"><tr>';
  d.trend.forEach(function(t){
    h += '<td class="pc"><div class="cv">' + pct_(t.s) + '</div>'
       + '<div class="col" style="height:' + Math.round(scale_(t.s) / 100 * BAR_MAX)
       + 'px;background:' + tone_(t.s) + '"></div></td>';
  });
  h += '</tr><tr>';
  d.trend.forEach(function(t){ h += '<td class="cx">' + esc_(t.k) + '</td>'; });
  h += '</tr></table></div>';

  h += '<table class="body"><tr><td class="cL">'
     + '<div class="eb">Top markdowns</div><table class="tbl hd">'
     + '<tr><th class="l nm2">Attribute</th><th class="bw2"></th>'
     + '<th class="vl">Of audits</th><th class="mt">Count</th></tr>';
  d.crit.forEach(function(c){
    const share = d.audits ? c.n / d.audits : 0;   // share of audits that hit this
    h += '<tr><td class="nm2">' + esc_(c.name) + '</td>'
       + '<td class="bw2"><div class="track"><div class="fill" style="width:'
       + Math.max(3, Math.round(share * 100)) + '%;background:' + C.flat + '"></div></div></td>'
       + '<td class="vl">' + pct_(share) + '</td>'
       + '<td class="mt">' + c.n + '</td></tr>';
  });
  h += '</table></td><td class="cR">';

  h += '<table class="tbl hd top">'
     + '<tr><th class="l nm">Pillar</th><th class="bw"></th><th class="vl">Score</th></tr>';
  d.pillars.forEach(function(p){
    h += '<tr><td class="nm">' + p[0] + '</td>' + barCell_(scale_(p[1]), tone_(p[1]))
       + '<td class="vl">' + pct_(p[1]) + '</td></tr>';
  });
  h += '</table>';

  h += '<table class="tbl hd gap">'
     + '<tr><th class="l nm">Channel</th><th class="bw"></th>'
     + '<th class="vl">Score</th><th class="mt">Audits</th></tr>';
  d.channels.forEach(function(c){
    const cs = c.sum / c.n;
    h += '<tr><td class="nm">' + esc_(c.name) + '</td>' + barCell_(scale_(cs), tone_(cs))
       + '<td class="vl">' + pct_(cs) + '</td><td class="mt">' + c.n + '</td></tr>';
  });
  h += '</table></td></tr></table>';

  h += '<table class="foot"><tr>'
     + '<td>View changes every ' + rot + 's</td>'
     + '<td class="rt">Last updated ' + stamp + '</td></tr></table>';

  return h;
}

/* ------------------------------- entry ------------------------------- */
function stamp_(){
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy · HH:mm');
}

function body_(only, rot){
  const stamp = stamp_();
  if (only && ORDER.indexOf(only) >= 0){
    return '<div class="v v1 solo">' + view_(only, stamp, rot) + '</div>';
  }
  let body = '';
  ORDER.forEach(function(mode, i){
    body += '<div class="v v' + (i + 1) + '">' + view_(mode, stamp, rot) + '</div>';
  });
  return body;
}

/* Called from the page every REFRESH_MINUTES. Returns just the four
   views, which the client drops into #root — no navigation, so nothing
   depends on the sandbox iframe being allowed to reload itself. */
function refreshBody(only, rot){
  ROWS_ = null;
  CacheService.getScriptCache().removeAll(ORDER.map(function(m){ return 'wb2_' + m; }));
  const r = Number(rot) > 0 ? Number(rot) : ROTATE_SECONDS;
  return body_(only || '', r);
}

function doGet(e){
  const q = (e && e.parameter) || {};
  const only = ORDER.indexOf(q.only) >= 0 ? q.only : '';
  const rot  = Number(q.rot) > 0 ? Number(q.rot) : ROTATE_SECONDS;
  const ref  = Number(q.ref) > 0 ? Number(q.ref) : REFRESH_MINUTES;
  try {
    return page_(body_(only, rot), !!only, rot, ref, only);
  } catch (err) {
    return errPage_(err && err.message ? err.message : String(err));
  }
}

/* A public viewer should never meet a stack trace. */
function errPage_(msg){
  const html =
      '<style>html,body{margin:0;height:100%;background:' + C.bg + ';color:' + C.fg + ';'
    + 'font-family:"IBM Plex Sans",Arial,sans-serif}'
    + '.w{padding:8vh 8vw}h1{font-size:34px;margin:0 0 14px}'
    + 'p{font-size:20px;line-height:1.5;color:' + C.mut + ';max-width:60ch}'
    + 'code{font-family:"IBM Plex Mono","Courier New",monospace;color:' + C.warn + '}</style>'
    + '<div class="w"><h1>' + esc_(BOARD_TITLE) + ' — not available</h1>'
    + '<p>The board could not read its data.</p><p><code>' + esc_(msg) + '</code></p>'
    + '<p>This page retries on its own every few minutes.</p></div>'
    + '<script>setTimeout(function(){location.reload()},300000);</scr' + 'ipt>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('QA Wall Board')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------- shell ------------------------------- */
function page_(body, solo, rot, ref, only){
  const F  = '"IBM Plex Sans",Arial,Helvetica,sans-serif';
  const FM = '"IBM Plex Mono","Courier New",monospace';
  const total = ORDER.length * rot;               // full cycle, seconds
  const vis   = Math.round(100 / ORDER.length);   // % of cycle each view is up

  let anim = '';
  for (let i = 0; i < ORDER.length; i++){
    anim += '.v' + (i + 1) + '{animation-delay:' + (i * rot) + 's;'
          + '-webkit-animation-delay:' + (i * rot) + 's}';
  }
  if (solo) anim = '.solo{opacity:1;visibility:visible;animation:none;-webkit-animation:none}';
  const frames =
        '0%{opacity:1;visibility:visible}'
      + (vis - 1) + '%{opacity:1;visibility:visible}'
      + vis + '%{opacity:0;visibility:hidden}'
      + '99.9%{opacity:0;visibility:hidden}'
      + '100%{opacity:1;visibility:visible}';

  const css = [
    '*{box-sizing:border-box}',
    'html,body{margin:0;padding:0;height:100%;background:' + C.bg + ';overflow:hidden}',
    'body{color:' + C.fg + ';font-family:' + F + ';position:relative}',
    'table{border-collapse:collapse;width:100%}',

    /* stacked views, cycled by CSS only */
    '.v{position:absolute;top:0;left:0;right:0;bottom:0;overflow:hidden;'
      + 'background:' + C.bg + ';padding:26px 34px 0;opacity:0;visibility:hidden;'
      + 'animation:cyc ' + total + 's step-end infinite;'
      + '-webkit-animation:cyc ' + total + 's step-end infinite}',
    '.v1{opacity:1;visibility:visible}',
    anim,
    '@keyframes cyc{' + frames + '}',
    '@-webkit-keyframes cyc{' + frames + '}',

    '.eb{font-family:' + FM + ';font-size:16px;letter-spacing:.16em;text-transform:uppercase;color:' + C.mut + ';white-space:nowrap}',
    '.sp{margin-top:16px}',

    '.rail{border-bottom:2px solid ' + C.line + ';padding-bottom:18px}',
    '.rail td{vertical-align:bottom}',
    '.org{font-family:' + FM + ';font-size:16px;letter-spacing:.24em;text-transform:uppercase;color:' + C.dim + '}',
    '.ttl{font-size:40px;font-weight:700;line-height:1.05;padding-top:4px}',
    '.pl{width:1%;white-space:nowrap;padding-bottom:8px}',
    '.pill{display:inline-block;vertical-align:middle;font-family:' + FM + ';font-size:14px;letter-spacing:.12em;color:' + C.dim + ';border:1px solid ' + C.line + ';border-radius:999px;padding:6px 12px;margin-left:7px}',
    '.pill.big{font-size:20px;font-weight:700;letter-spacing:.16em;padding:9px 20px}',
    '.pill.on{background:' + C.acc + ';border-color:' + C.acc + ';color:#08111F;font-weight:600}',

    '.kpi{margin-top:22px}',
    '.kt{padding:0 7px;vertical-align:top}',
    '.kt{width:20%}',
    '.kt:first-child{padding-left:0}.kt:last-child{padding-right:0}',
    '.kbox{height:140px;overflow:hidden;text-align:center;background:' + C.surf + ';border:1px solid ' + C.line + ';border-left:4px solid ' + C.line + ';border-radius:12px;padding:18px 22px 20px}',
    '.kn{font-size:60px;font-weight:700;line-height:1.05;padding-top:6px;white-space:nowrap}',
    '.ksm{font-size:40px;padding:15px 0 14px}',
    '.tgbox{padding:15px 18px 12px}',
    '.tg{width:auto;margin:7px auto 0}',
    '.tg td{font-family:' + FM + ';font-size:19px;line-height:1.25;padding:1px 0;white-space:nowrap}',
    '.tg td.tgn{text-align:left;letter-spacing:.04em;padding:1px 26px 1px 0}',
    '.tg td.tgv{text-align:right;font-weight:600;padding:1px 0}',

    '.trend{position:relative;margin-top:22px;background:' + C.surf + ';border:1px solid ' + C.line + ';border-radius:12px;padding:14px 20px 8px;height:246px}',
    '.guide{position:absolute;left:22px;right:22px;height:0;border-top:1px dashed rgba(224,179,60,.45)}',
    '.cols{margin-top:26px}',
    '.cols td{text-align:center;padding:0 5px}',
    '.cols td.pc{vertical-align:bottom;height:' + PLOT_H + 'px}',
    '.cv{font-family:' + FM + ';font-size:15px;color:' + C.mut + ';padding-bottom:5px}',
    '.col{width:92px;margin:0 auto;border-radius:4px 4px 0 0}',
    '.cx{font-family:' + FM + ';font-size:16px;font-weight:600;color:' + C.mut + ';padding:9px 5px 0}',

    '.body{margin-top:20px}',
    '.body>tbody>tr>td{vertical-align:top}',
    '.cL{width:54%;padding-right:32px}',
    '.cR{width:46%}',
    '.tbl{margin-top:8px}',
    '.tbl.hd th{font-size:16px;letter-spacing:.16em;color:' + C.mut + '}',
    '.tbl.top{margin-top:27px}',
    '.tbl.gap{margin-top:30px}',
    '.tbl td{padding:5px 6px;border-bottom:1px solid ' + C.line2 + ';font-size:21px}',
    '.tbl th{font-family:' + FM + ';font-weight:500;font-size:14px;letter-spacing:.18em;text-transform:uppercase;color:' + C.dim + ';padding:0 6px 7px;border-bottom:1px solid ' + C.line + ';text-align:right;white-space:nowrap}',
    '.tbl th.l{text-align:left}',
    '.nm{width:44%;white-space:nowrap;overflow:hidden}',
    '.nm2{width:34%;white-space:nowrap;overflow:hidden}',
    '.bw2{width:41%}',
    '.bw{width:36%}',
    '.vl{width:13%;text-align:right;font-weight:600;white-space:nowrap;font-family:' + FM + ';font-size:22px}',
    '.mt{width:10%;text-align:right;color:' + C.mut + ';font-family:' + FM + ';font-size:19px}',
    '.track{background:' + C.track + ';height:16px;border-radius:8px}',
    '.fill{height:16px;border-radius:8px}',

    '.foot{margin-top:16px;border-top:1px solid ' + C.line2 + '}',
    '.foot td{font-family:' + FM + ';font-size:15px;color:' + C.dim + ';padding:12px 0 16px}',
    '.foot .ct{text-align:center}.foot .rt{text-align:right}'
  ].join('');

  /* Data refresh, in order of preference:
       1. google.script.run — repaints #root in place, no navigation.
       2. top-level navigation to the /exec URL, for browsers where the
          client API is unavailable (very old Smart-TV engines).
     A failed pull retries once after a minute before falling back, so a
     transient network blip does not send the TV to a Google error page. */
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (err) { url = ''; }

  const js =
      'var ONLY=' + JSON.stringify(only || '') + ',ROT=' + rot + ',URL=' + JSON.stringify(url) + ';'
    + 'var MS=' + (ref * 60000) + ',tries=0;'
    + 'function hard(){ if(!URL){return;} try{ if(window.top&&window.top!==window){window.top.location.href=URL;}'
    + 'else{window.location.href=URL;} }catch(e){ try{window.open(URL,"_top");}catch(e2){} } }'
    + 'function ok(h){ tries=0; var r=document.getElementById("root");'
    + 'if(h&&r){ r.innerHTML=h; } setTimeout(pull,MS); }'
    + 'function bad(){ tries++; if(tries<2){ setTimeout(pull,60000); } else { hard(); } }'
    + 'function pull(){ try{ if(window.google&&google.script&&google.script.run){'
    + 'google.script.run.withSuccessHandler(ok).withFailureHandler(bad).refreshBody(ONLY,ROT);'
    + '} else { hard(); } }catch(e){ hard(); } }'
    + 'setTimeout(pull,MS);';

  const html =
      '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?'
    + 'family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap">'
    + '<style>' + css + '</style>'
    + '<div id="root">' + body + '</div>'
    + '<script>' + js + '</scr' + 'ipt>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('QA Wall Board')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
