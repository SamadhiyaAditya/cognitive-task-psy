/* =====================================================================
   HRR task – Google Apps Script backend (v2)

   SETUP (one time)
   1. Open the Google Sheet > Extensions > Apps Script. Replace everything with this file. Save.
   2. Deploy > Manage deployments > (pencil icon on the existing deployment) > Version: "New version" > Deploy.
      Editing the existing deployment keeps the same /exec URL. If you make a NEW deployment instead,
      copy its URL into SCRIPT_URL in index.html and report.html.
      Settings must be:  Execute as: Me   |   Who has access: Anyone
      (NOT "Anyone within nst.rishihood.edu.in" – outside people are silently blocked.)
   3. Open  <your /exec URL>?action=ping  in an incognito window. You must see {"ok":true,...}.
      If you see a Google login page, the access setting is wrong.

   SHEETS (created automatically)
     participants       one row per session: ID, condition, demographics, email, status
     trials             one row per trial / probe (no email)
     participants_test, trials_test   same, for test runs (index.html?test=1)
   The old "data" tab is never touched.
   ===================================================================== */

const VERSION = '2.0';
const N_TRIALS = 160, N_PROBES = 8;           // expected rows for a complete session
const RESERVE_MIN = 30;                        // a started-but-unfinished session holds its condition slot this long
const CONDITIONS = [[85,'A'],[85,'B'],[70,'A'],[70,'B']];

const PART_COLS = ['participant_id','session_id','run_mode','status','validity_pct','mapping',
  'age','gender','course','email','noticed_change','trials_saved','probes_saved','tab_hidden_count','practice_rounds','practice_correct','timeout_warnings',
  'started_at','completed_at','user_agent','screen','refresh_hz','app_version'];
const TRIAL_COLS = ['row_key','session_id','participant_id','run_mode','age','gender','course',
  'validity_pct','mapping','row_type','trial','phase','phase_name','trial_in_phase','cue','rule_side',
  'target_side','validity','response','correct','RT','timeout','anticipation','soa_actual_ms',
  'hidden_during_trial','probe_expected_side','explicit_matches_rule','probe_confidence','RT_probe_ms','timestamp'];

/* ---------- helpers ---------- */
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const isTest = mode => mode === 'test';
function sheet(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
const partSheet  = mode => sheet(isTest(mode) ? 'participants_test' : 'participants', PART_COLS);
const trialSheet = mode => sheet(isTest(mode) ? 'trials_test' : 'trials', TRIAL_COLS);
function readObjects(sh, cols) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2,1,n,cols.length).getValues().map((r,i) => { const o = {_row:i+2}; cols.forEach((c,j)=>o[c]=r[j]); return o; });
}
function findSession(mode, sid) {
  const sh = partSheet(mode);
  return { sh, p: readObjects(sh, PART_COLS).find(p => p.session_id === sid) };
}

/* ---------- web endpoints ---------- */
function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || 'report';
  if (a === 'ping') return out({ok:true, version:VERSION});
  // report data: complete, real sessions only; email removed
  const P = readObjects(partSheet('real'), PART_COLS).filter(p => p.status === 'complete');
  const ids = new Set(P.map(p => p.session_id));
  const T = readObjects(trialSheet('real'), TRIAL_COLS).filter(t => ids.has(t.session_id));
  P.forEach(p => { delete p.email; delete p._row; });
  T.forEach(t => delete t._row);
  return out({ok:true, participants:P, trials:T});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return out({ok:false, error:'busy, retry'});
  try {
    const b = JSON.parse(e.postData.contents);
    const mode = b.run_mode === 'test' ? 'test' : 'real';
    if (b.action === 'start')    return out(start(b, mode));
    if (b.action === 'rows')     return out(saveRows(b, mode));
    if (b.action === 'complete') return out(complete(b, mode));
    return out({ok:false, error:'unknown action'});
  } catch (err) {
    return out({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

/* start: duplicate check, sequential ID, balanced condition */
function start(b, mode) {
  const sh = partSheet(mode), P = readObjects(sh, PART_COLS);
  const email = String(b.email || '').trim().toLowerCase();
  if (!isTest(mode) && P.some(p => String(p.email).toLowerCase() === email && (p.status === 'complete' || String(p.status).startsWith('excluded'))))
    return {ok:false, error:'already_participated'};
  const now = Date.now();
  const active = P.filter(p => p.status === 'complete' || (p.status === 'started' && now - new Date(p.started_at).getTime() < RESERVE_MIN*60000));
  const counts = CONDITIONS.map(([v,m]) => active.filter(p => +p.validity_pct === v && p.mapping === m).length);
  const min = Math.min(...counts), pool = CONDITIONS.filter((c,i) => counts[i] === min);
  const [validity_pct, mapping] = pool[Math.floor(Math.random()*pool.length)];
  const participant_id = P.reduce((mx,p) => Math.max(mx, +p.participant_id || 0), 0) + 1;
  const session_id = Utilities.getUuid();
  const row = {participant_id, session_id, run_mode:mode, status:'started', validity_pct, mapping,
    age:b.age, gender:b.gender, course:b.course, email, started_at:new Date().toISOString(),
    user_agent:b.user_agent, screen:b.screen, refresh_hz:b.refresh_hz, app_version:b.app_version};
  sh.appendRow(PART_COLS.map(c => row[c] === undefined ? '' : row[c]));
  return {ok:true, participant_id, session_id, validity_pct, mapping};
}

/* rows: append only rows whose row_key is not already in the sheet (safe to resend) */
function saveRows(b, mode) {
  if (!findSession(mode, b.session_id).p) return {ok:false, error:'unknown session'};
  const sh = trialSheet(mode), n = sh.getLastRow() - 1;
  const have = new Set(n > 0 ? sh.getRange(2,1,n,1).getValues().map(r => r[0]) : []);
  const fresh = (b.rows || []).filter(r => r.session_id === b.session_id && r.row_key && !have.has(r.row_key) && have.add(r.row_key));
  if (fresh.length) sh.getRange(sh.getLastRow()+1, 1, fresh.length, TRIAL_COLS.length)
                      .setValues(fresh.map(r => TRIAL_COLS.map(c => r[c] === undefined || r[c] === null ? '' : r[c])));
  return {ok:true, written:fresh.length, ...countRows(mode, b.session_id)};
}
function countRows(mode, sid) {
  const sh = trialSheet(mode), n = sh.getLastRow() - 1;
  if (n < 1) return {trials_saved:0, probes_saved:0};
  const v = sh.getRange(2,1,n,10).getValues();      // cols: row_key, session_id ... row_type (10th)
  const mine = v.filter(r => r[1] === sid);
  return {trials_saved: mine.filter(r => r[9] === 'trial').length, probes_saved: mine.filter(r => r[9] === 'probe').length};
}

/* complete: mark session finished only if every row is in the sheet */
function complete(b, mode) {
  const {sh, p} = findSession(mode, b.session_id);
  if (!p) return {ok:false, error:'unknown session'};
  const c = countRows(mode, b.session_id);
  const full = c.trials_saved === N_TRIALS && c.probes_saved === N_PROBES;
  const set = (col, val) => sh.getRange(p._row, PART_COLS.indexOf(col)+1).setValue(val);
  set('noticed_change', b.noticed_change ?? '');
  ['tab_hidden_count','practice_rounds','practice_correct','timeout_warnings'].forEach(k => set(k, b[k] ?? ''));
  set('trials_saved', c.trials_saved);
  set('probes_saved', c.probes_saved);
  if (full) { set('status','complete'); set('completed_at', new Date().toISOString()); }
  return {ok:full, complete:full, ...c};
}
