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

/* Applied AI lockup, white on transparent, embedded so the board stays a
   single self-contained response with no external image request. Rendered
   at 2x (459x76) and displayed at LOGO_W x LOGO_H, which keeps it crisp on
   a 1080p panel and on a laptop. Set LOGO to '' to drop it entirely. */
const LOGO_W = 230;
const LOGO_H = 38;
const LOGO = 'data:image/png;base64,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';

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

  /* Three cells at 34/32/34 mean the middle one spans 34%-66%, so the
     logo sits on the page's exact centre line without absolute
     positioning — plain table maths every browser already agrees on. */
  h += '<table class="rail"><tr>'
     + '<td class="who"><div class="org">' + esc_(ORG_LINE) + '</div>'
     + '<div class="ttl">' + esc_(BOARD_TITLE) + '</div></td>'
     + '<td class="lg">'
     + (LOGO ? '<img src="' + LOGO + '" width="' + LOGO_W + '" height="' + LOGO_H
               + '" alt="Applied AI">' : '')
     + '</td>'
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
    '.who{width:34%}',
    '.lg{width:32%;text-align:center;padding-bottom:9px}',
    '.lg img{display:inline-block;vertical-align:middle;border:0}',
    '.pl{width:34%;text-align:right;white-space:nowrap;padding-bottom:8px}',
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
