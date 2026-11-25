/**
 * Toast Strength Log — Google Sheets receiver + run importer
 *
 * WHAT THIS DOES
 *  - doPost: receives strength sessions, meals, weights and workout edits from the app.
 *  - doGet:  returns tabs (Strength, Runs, Weight, Nutrition, Workouts, recap) as JSON,
 *            and runs the Claude actions (recommend, estimate, build).
 *  - importRuns: reads HealthFit's run CSVs from a Google Drive folder, digests each
 *    into summary + interval + split rows, and writes them to a "Runs" tab in this
 *    same workbook. Runs on an hourly trigger and skips files it has already imported.
 *
 * SECRETS (Project Settings ▸ Script Properties) — never paste them into this code:
 *   TOKEN              the token the app sends (same value as in the app's Settings)
 *   ANTHROPIC_API_KEY  your key from console.anthropic.com
 * If TOKEN is missing every request is refused.
 *
 * ONE-TIME SETUP FOR RUNS (after pasting this in):
 *  1. Run the function `installRunTrigger` once (pick it from the toolbar dropdown ▸ Run).
 *     It asks for Drive permission the first time — approve it. This also does an
 *     immediate first import so your existing run files get picked up.
 *  2. Re-deploy: Deploy ▸ Manage deployments ▸ edit ▸ New version ▸ Deploy.
 */

const SHEET_ID = '1zZM2gOqpmLtARR_ckhV4p2a4i9oQ5Y9WO8qvD8yjPsk';
const SHEET_NAME = 'Strength';

function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}
function tokenOk_(t) {
  const want = prop_('TOKEN');
  return !!want && t === want;
}
function apiKey_() { return prop_('ANTHROPIC_API_KEY'); }
const NO_KEY_MSG = 'No ANTHROPIC_API_KEY in Script Properties yet.';

// ----- Runs config -----
const RUN_FOLDER_ID = '1DRgXMBSDG-iyZSPxXhx0hpPBrKfyF8wt'; // HealthFit auto-export folder
const RUN_SHEET_NAME = 'Runs';
const RUN_HEADER = [
  'run_id', 'date', 'run_type', 'section', 'seq', 'label',
  'distance', 'duration', 'pace', 'avg_hr', 'avg_power', 'cadence', 'imported_at'
];

const HEADER = [
  'session_id', 'date', 'session', 'exercise', 'target', 'set',
  'weight', 'reps', 'set_note', 'exercise_feel', 'exercise_note',
  'session_feel', 'session_note', 'logged_at'
];

// ----- Recap config -----
// The daily task writes recap .txt files (named YYYY-MM-DD.txt) into this Drive
// folder; doGet ?tab=recap returns the newest one.
const RECAP_FOLDER_ID = '1RkVNs0gFMt11C0UqRoVBZXcykWF5u8-V';

// ----- Workout recommendations (Claude) -----
const RECOMMEND_MODEL = 'claude-haiku-4-5-20251001'; // fast + cheap; 'claude-sonnet-5' is smarter but ~20s (too slow for the app)

// ----- Nutrition + Weight config -----
const NUTRITION_SHEET_NAME = 'Nutrition';
const NUTRITION_HEADER = ['meal_id', 'date', 'name', 'amount', 'unit', 'cal', 'protein', 'carbs', 'fat', 'logged_at', 'meal_type'];
const WEIGHT_SHEET_NAME = 'Weight';
const WEIGHT_HEADER = ['date', 'lb', 'logged_at'];

// ----- Workouts (the app's workout picker) -----
// One row per movement. Rows sharing a workout_id are one workout, in row order.
// Only the title (workout) is needed for workouts you add by hand; the app keys
// those by title until it gives them an id.
const WORKOUT_SHEET_NAME = 'Workouts';
const WORKOUT_HEADER = ['workout_id', 'workout', 'exercise', 'target', 'sets', 'per_side'];

// ----- Plan (day-by-day schedule; written by planning chats + the morning email) -----
const PLAN_SHEET_NAME = 'Plan';
const PLAN_HEADER = ['date', 'weekday', 'session', 'day_type', 'prescription', 'status', 'note', 'updated_by', 'updated_at'];

// =====================================================================
//  STRENGTH
// =====================================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!tokenOk_(body.token)) {
      return json({ ok: false, error: 'bad token' });
    }
    if (body.meal) { return saveMeal(body.meal); }
    if (body.deleteMeal) { return deleteMealRow(body.deleteMeal.id); }
    if (body.weight) { return saveWeight(body.weight); }
    if (body.saveWorkout) { return saveWorkout(body.saveWorkout); }
    if (body.renameWorkout) { return renameWorkout(body.renameWorkout); }
    if (body.deleteWorkout) { return deleteWorkout(body.deleteWorkout.id); }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) sh = ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(HEADER);

    // upsert: drop any existing rows for this session_id so re-syncing never duplicates
    const sid = body.session && body.session.id;
    if (sid) removeSessionRows(sh, sid);

    const rows = flatten(body.session);
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);
    }
    return json({ ok: true, rows: rows.length });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// doGet: ?token=...               -> Strength tab (default)
//        ?token=...&tab=runs      -> Runs tab
//        ?token=...&tab=weight    -> Weight tab
//        ?token=...&tab=workouts  -> Workouts tab
//        ?token=...&tab=plan      -> Plan tab (day-by-day schedule)
//        ?token=...&tab=nutrition -> Nutrition tab (optionally &date=YYYY-MM-DD)
//        ?token=...&tab=recap     -> newest recap file
function doGet(e) {
  const token = e && e.parameter && e.parameter.token;
  if (!tokenOk_(token)) {
    return json({ ok: false, error: 'bad token' });
  }
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'recommend') { return recommend(e); }
  if (action === 'estimate') { return estimateMeal(e); }
  if (action === 'build') { return buildWorkout(e); }
  const p = (e && e.parameter && e.parameter.tab) || '';
  if (p === 'recap') { return getRecap(); }
  if (p === 'nutrition') { return getMeals(e); }
  let tab = SHEET_NAME, header = HEADER;
  if (p === 'runs') { tab = RUN_SHEET_NAME; header = RUN_HEADER; }
  else if (p === 'weight') { tab = WEIGHT_SHEET_NAME; header = WEIGHT_HEADER; }
  else if (p === 'workouts') { tab = WORKOUT_SHEET_NAME; header = WORKOUT_HEADER; }
  else if (p === 'plan') { tab = PLAN_SHEET_NAME; header = PLAN_HEADER; }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return json({ ok: true, rows: [] });

    const tz = ss.getSpreadsheetTimeZone();
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, header.length).getValues();
    const rows = values.map(function (row) {
      const obj = {};
      header.forEach(function (key, i) {
        let v = row[i];
        if (Object.prototype.toString.call(v) === '[object Date]') {
          v = (key === 'date') ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v.getTime();
        }
        obj[key] = v;
      });
      return obj;
    });
    return json({ ok: true, rows: rows });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Return the newest recap from the recap Drive folder.
//  - a YYYY-MM-DD.json file -> { date, kind:'json', recap:<object>, saved_at }
//  - a YYYY-MM-DD.txt file  -> { date, kind:'text', text, saved_at }
// For the same date a .json is preferred over a .txt.
function getRecap() {
  try {
    const folder = DriveApp.getFolderById(RECAP_FOLDER_ID);
    const files = folder.getFiles();
    var newest = null, newestKey = '', newestJson = false;
    while (files.hasNext()) {
      var f = files.next();
      var name = f.getName();
      var isJson = /\.json$/i.test(name), isTxt = /\.txt$/i.test(name);
      if (!isJson && !isTxt) continue;
      var m = name.match(/(\d{4}-\d{2}-\d{2})/);
      var key = m ? m[1] : Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var better = false;
      if (!newest) better = true;
      else if (key > newestKey) better = true;
      else if (key === newestKey) {
        if (isJson && !newestJson) better = true;
        else if (isJson === newestJson && f.getLastUpdated() > newest.getLastUpdated()) better = true;
      }
      if (better) { newest = f; newestKey = key; newestJson = isJson; }
    }
    if (!newest) return json({ ok: true, rows: [] });
    var content = newest.getBlob().getDataAsString();
    var stamp = newest.getLastUpdated().getTime();
    if (newestJson) {
      try { return json({ ok: true, rows: [{ date: newestKey, kind: 'json', recap: JSON.parse(content), saved_at: stamp }] }); }
      catch (e) { return json({ ok: true, rows: [{ date: newestKey, kind: 'text', text: content, saved_at: stamp }] }); }
    }
    return json({ ok: true, rows: [{ date: newestKey, kind: 'text', text: content, saved_at: stamp }] });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// ===== Nutrition + Weight =====
// (named apart from the runs `_num` below, which returns null — a second
// function with the same name silently replaces the first in Apps Script)
function _numOr0(x) { var n = parseFloat(x); return isNaN(n) ? 0 : n; }

// Save (or update, by meal_id) one logged food/meal in the Nutrition tab.
function saveMeal(m) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(NUTRITION_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(NUTRITION_SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(NUTRITION_HEADER);
  if (sh.getRange(1, NUTRITION_HEADER.length).getValue() === '') {
    sh.getRange(1, 1, 1, NUTRITION_HEADER.length).setValues([NUTRITION_HEADER]);
  }
  const row = [m.id || '', m.date || '', m.name || '', _numOr0(m.amount), m.unit || '', _numOr0(m.cal), _numOr0(m.p), _numOr0(m.c), _numOr0(m.f), new Date(), m.mealType || ''];
  const last = sh.getLastRow();
  if (last >= 2 && m.id) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(m.id)) { sh.getRange(i + 2, 1, 1, NUTRITION_HEADER.length).setValues([row]); return json({ ok: true, meal: 'updated' }); }
    }
  }
  sh.appendRow(row);
  return json({ ok: true, meal: 'added' });
}

function deleteMealRow(id) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(NUTRITION_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return json({ ok: true, deleted: false });
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); return json({ ok: true, deleted: true }); }
  }
  return json({ ok: true, deleted: false });
}

function saveWeight(w) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(WEIGHT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(WEIGHT_SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(WEIGHT_HEADER);
  const tz = ss.getSpreadsheetTimeZone();
  const date = String(w.date || '');
  const lb = parseFloat(w.lb);
  if (!(lb > 0)) return json({ ok: false, error: 'bad weight' });
  const last = sh.getLastRow();
  if (last >= 2) {
    const ds = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ds.length; i++) {
      var c = ds[i][0];
      var cs = Object.prototype.toString.call(c) === '[object Date]' ? Utilities.formatDate(c, tz, 'yyyy-MM-dd') : String(c);
      if (cs === date) { sh.getRange(i + 2, 1, 1, WEIGHT_HEADER.length).setValues([[date, lb, new Date()]]); return json({ ok: true, weight: 'updated' }); }
    }
  }
  sh.appendRow([date, lb, new Date()]);
  return json({ ok: true, weight: 'added' });
}

// GET ?tab=nutrition&date=YYYY-MM-DD  -> meals for that date (or all if no date)
function getMeals(e) {
  try {
    const date = e && e.parameter && e.parameter.date;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(NUTRITION_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return json({ ok: true, rows: [] });
    const tz = ss.getSpreadsheetTimeZone();
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, NUTRITION_HEADER.length).getValues();
    const rows = [];
    values.forEach(function (row) {
      const o = {};
      NUTRITION_HEADER.forEach(function (k, i) {
        var v = row[i];
        if (Object.prototype.toString.call(v) === '[object Date]') { v = (k === 'date') ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v.getTime(); }
        o[k] = v;
      });
      if (!date || String(o.date) === String(date)) rows.push(o);
    });
    return json({ ok: true, rows: rows });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}

// ===== Workouts =====
function workoutSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(WORKOUT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(WORKOUT_SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(WORKOUT_HEADER);
  return sh;
}
// Same rule as the app: a row without an id belongs to the workout keyed by its title.
function rowWorkoutId_(row) {
  const id = String(row[0] || '').trim();
  return id || ('sheet:' + String(row[1] || '').trim());
}
function workoutRowIndexes_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  const out = [];
  vals.forEach(function (r, i) { if (rowWorkoutId_(r) === String(id)) out.push(i + 2); });
  return out;
}

// POST {saveWorkout:{id, name, exercises:[{name,target,sets,perSide}]}} — replaces that workout's rows.
function saveWorkout(w) {
  if (!w || !w.id) return json({ ok: false, error: 'workout id required' });
  const sh = workoutSheet_();
  workoutRowIndexes_(sh, w.id).reverse().forEach(function (r) { sh.deleteRow(r); });
  var rows = (w.exercises || []).filter(function (x) { return x && x.name; }).map(function (x) {
    return [w.id, w.name || '', x.name, x.target || '', parseInt(x.sets, 10) || 3, !!x.perSide];
  });
  if (!rows.length) rows = [[w.id, w.name || '', '', '', '', false]];
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, WORKOUT_HEADER.length).setValues(rows);
  return json({ ok: true, rows: rows.length });
}

// POST {renameWorkout:{id, name}} — retitles every row of that workout.
function renameWorkout(r) {
  if (!r || !r.id || !r.name) return json({ ok: false, error: 'id and name required' });
  const sh = workoutSheet_();
  const idx = workoutRowIndexes_(sh, r.id);
  idx.forEach(function (row) { sh.getRange(row, 1, 1, 2).setValues([[String(r.id), r.name]]); });
  return json({ ok: true, rows: idx.length });
}

// POST {deleteWorkout:{id}}
function deleteWorkout(id) {
  if (!id) return json({ ok: false, error: 'id required' });
  const sh = workoutSheet_();
  const idx = workoutRowIndexes_(sh, id);
  idx.reverse().forEach(function (r) { sh.deleteRow(r); });
  return json({ ok: true, deleted: idx.length });
}

// GET ?action=estimate&meal=<description> -> Claude estimates macros as JSON
function estimateMeal(e) {
  try {
    const key = apiKey_();
    if (!key) return json({ ok: false, error: NO_KEY_MSG });
    const meal = ((e.parameter && e.parameter.meal) || '').trim();
    if (!meal) return json({ ok: false, error: 'No meal text.' });
    const system = 'You estimate nutrition macros of a described meal for an athlete tracking a cut. '
      + 'US units: calories (kcal), protein/carbs/fat in grams. '
      + 'Assume COOKED food weights unless the text explicitly says raw. '
      + 'Return STRICT JSON ONLY: {"name":"<short label>","amount":<number>,"unit":"<concrete unit>","cal":<int>,"p":<int>,"c":<int>,"f":<int>}. '
      + 'NEVER use "serving". Prefer OUNCES ("oz") as the unit for meats and mixed or composite meals, using the cooked weight in oz (e.g. amount 8, unit "oz"); '
      + 'use a natural unit only for clearly discrete or volume foods (cup, slice, piece, tbsp). '
      + 'Be accurate for cooked weights, do not underestimate: for reference, 8 oz cooked lean ground chicken is about 340 kcal and 46 g protein, and 6 oz cooked chicken breast is about 280 kcal and 53 g protein. '
      + 'Give the macros for the exact amount and unit you return.';
    const user = 'Meal: ' + meal + '\nReturn strict JSON.';
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: RECOMMEND_MODEL, max_tokens: 400, system: system, messages: [{ role: 'user', content: user }] }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return json({ ok: false, error: 'Claude API ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 160) });
    const data = JSON.parse(resp.getContentText());
    var text = '';
    if (data && data.content) { for (var i = 0; i < data.content.length; i++) { var b = data.content[i]; if (b && b.type === 'text' && b.text) text += b.text; } }
    var mm = text.match(/\{[\s\S]*\}/);
    if (!mm) return json({ ok: false, error: 'Claude did not return JSON.' });
    return json({ ok: true, estimate: JSON.parse(mm[0]) });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}

// Conversational workout builder: chat back and forth to design/refine one session.
// GET ?action=build&token=...&msg=<latest user message>&current=<current workout JSON, optional>
function buildWorkout(e) {
  try {
    const key = apiKey_();
    if (!key) return json({ ok: false, error: NO_KEY_MSG });
    const msg = ((e.parameter && e.parameter.msg) || (e.parameter && e.parameter.desc) || '').trim();
    if (!msg) return json({ ok: false, error: 'No message.' });
    const current = ((e.parameter && e.parameter.current) || '').trim();
    const system = 'You are the coach for Stephen "Toast" Loper, a hybrid athlete (Australian Rules footballer, half-marathoner, lifter) in a gentle cut who is carefully reintroducing lower-body work after a knee issue. '
      + 'You are having a short back-and-forth chat to design and refine ONE training session. Talk like a coach: warm, brief, plain. '
      + 'In "reply" (1 to 4 sentences), say what you built or changed and WHY, in his terms. If he only asked a question, answer it and keep the workout as-is. '
      + 'Keep sessions realistic and time-appropriate (about 45-55 minutes / 5 to 8 exercises unless he says otherwise). '
      + 'If the request is rehab, prehab, or reintroduction, keep it cautious, controlled, and knee-friendly. Never use em dashes anywhere. '
      + 'Return STRICT JSON ONLY, no prose outside it: '
      + '{"reply":"<your short chat reply>","workout":{"name":"<short workout title>","exercises":[{"name":"<exercise>","target":"<e.g. 3×10 or 3×10 each>","sets":<int>,"perSide":<true if per-side/each-leg else false>}]}}. '
      + 'Always include the full current workout in "workout" (updated if you changed it, unchanged if you did not). Use "each" in target and perSide:true for unilateral movements.';
    const user = (current ? ('Current workout JSON: ' + current + '\n\n') : 'No workout yet.\n\n')
      + 'Stephen says: ' + msg + '\n\nReturn strict JSON.';
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: RECOMMEND_MODEL, max_tokens: 1500, system: system, messages: [{ role: 'user', content: user }] }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return json({ ok: false, error: 'Claude API ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 160) });
    const data = JSON.parse(resp.getContentText());
    var text = '';
    if (data && data.content) { for (var i = 0; i < data.content.length; i++) { var b = data.content[i]; if (b && b.type === 'text' && b.text) text += b.text; } }
    var mm = text.match(/\{[\s\S]*\}/);
    if (!mm) return json({ ok: false, error: 'Claude did not return JSON.' });
    var parsed = JSON.parse(mm[0]);
    return json({ ok: true, reply: parsed.reply || '', workout: parsed.workout || null });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}

// ----- movement-name matching for recommendations -----
// Names logged by hand drift ("Db rdl", "Incline db press", "Single arm cable row"),
// so history is matched on a normalized token set, and similar movements are
// offered as a guide when an exercise has little history of its own.
const MOVE_STOP = { each: 1, side: 1, sides: 1, x: 1, with: 1, and: 1, the: 1, a: 1, of: 1, on: 1, to: 1, for: 1, w: 1 };
const MOVE_HEADS = ['row', 'press', 'squat', 'deadlift', 'rdl', 'curl', 'raise', 'extension', 'pull', 'pulldown', 'pullup',
  'chinup', 'pushup', 'lunge', 'bridge', 'thrust', 'carry', 'crunch', 'slam', 'throw', 'pass', 'swing', 'fly', 'dip',
  'plank', 'bug', 'walk', 'step', 'jump', 'clean', 'snatch', 'jerk', 'thruster', 'bench', 'copenhagen', 'hinge', 'twist'];
// Words that describe how a movement is done rather than what it is; ignored when
// judging similarity (so "Single Leg Calf Raise" ~ "Standing Calf Raise", not "Straight-Leg Raise").
const MOVE_GENERIC = { single: 1, one: 1, leg: 1, arm: 1, alt: 1, alternating: 1, standing: 1, seated: 1, banded: 1 };
const SIMILAR_MIN = 0.34;

function moveTokens_(name) {
  const s = String(name || '').toLowerCase()
    .replace(/[‘’]/g, "'")
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[-\/_,.+&:]/g, ' ')
    .replace(/\bdumb ?bells?\b/g, 'db')
    .replace(/\bkettle ?bells?\b/g, 'kb')
    .replace(/\bbar ?bells?\b/g, 'bb')
    .replace(/\bromanian dead ?lifts?\b/g, 'rdl')
    .replace(/\bmedball\b/g, 'med ball')
    .replace(/\b(pull|push|chin) ups?\b/g, '$1up');
  return s.split(/\s+/).filter(function (t) { return t && !MOVE_STOP[t] && !/^\d/.test(t); }).map(function (t) {
    if (/(ss|sh|ch|x)es$/.test(t)) return t.slice(0, -2);
    if (t.length > 3 && /s$/.test(t) && !/ss$/.test(t)) return t.slice(0, -1);
    return t;
  });
}
function moveKey_(name) { return moveTokens_(name).slice().sort().join(' '); }
function moveSimilarity_(a, b) {
  const A = {}, B = {};
  a.forEach(function (t) { if (!MOVE_GENERIC[t]) A[t] = 1; });
  b.forEach(function (t) { if (!MOVE_GENERIC[t]) B[t] = 1; });
  var inter = 0, union = 0, sharedHead = false, anyHead = false;
  Object.keys(A).forEach(function (t) { if (B[t]) { inter++; if (MOVE_HEADS.indexOf(t) >= 0) sharedHead = true; } });
  Object.keys(A).concat(Object.keys(B)).forEach(function (t) { if (MOVE_HEADS.indexOf(t) >= 0) anyHead = true; });
  union = Object.keys(A).length + Object.keys(B).filter(function (t) { return !A[t]; }).length;
  if (!union || (anyHead && !sharedHead)) return 0;
  return inter / union;
}

// History lines for one movement: newest `n` dates, one session per date (re-synced
// duplicates of a session would otherwise double the sets shown).
function historyLines_(rows, n, indent) {
  var byDate = {};
  rows.forEach(function (x) {
    var d = byDate[x.date] = byDate[x.date] || { sid: x.sid, rows: [], seen: {} };
    if (x.sid !== d.sid) { d.sid = x.sid; d.rows = []; d.seen = {}; } // later rows win: the sheet is append-only
    if (x.set !== '' && d.seen[x.set]) return; // the same set written twice
    d.seen[x.set] = 1;
    d.rows.push(x);
  });
  return Object.keys(byDate).sort().slice(-n).map(function (dt) {
    var rs = byDate[dt].rows;
    var sets = rs.map(function (x) { return (x.weight || 'BW') + '×' + (x.reps || '?'); }).join(', ');
    var feel = rs[0].exFeel ? (' (felt ' + rs[0].exFeel + '/5)') : '';
    var note = '';
    for (var i = 0; i < rs.length; i++) { if (rs[i].note) { note = ' | note: ' + rs[i].note; break; } }
    var sn = rs[0].sessNote ? (' | session note: ' + rs[0].sessNote) : '';
    return indent + dt + ': ' + sets + feel + note + sn;
  });
}

// Recommend weights/reps for a workout by sending recent history to Claude.
// GET ?action=recommend&token=...&workout=<name>&ex=Name~sets|Name~sets|...
function recommend(e) {
  try {
    const key = apiKey_();
    if (!key) return json({ ok: false, error: NO_KEY_MSG });
    const workout = e.parameter.workout || 'Workout';
    const exParam = e.parameter.ex || '';
    const exList = exParam.split('|').filter(String).map(function (s) {
      var parts = s.split('~');
      return { name: parts[0], sets: parseInt(parts[1] || '3', 10) || 3 };
    });
    if (!exList.length) return json({ ok: false, error: 'No exercises provided.' });

    // Group the Strength tab by normalized movement name.
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    const byKey = {}; // moveKey -> { names: {loggedName: count}, rows: [] }
    if (sh && sh.getLastRow() >= 2) {
      const tz = ss.getSpreadsheetTimeZone();
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEADER.length).getValues();
      // cols: session_id,date,session,exercise,target,set,weight,reps,set_note,exercise_feel,exercise_note,session_feel,session_note,logged_at
      vals.forEach(function (r) {
        var name = String(r[3] || '').trim();
        if (!name || name.indexOf('Push/Pull Ladder') === 0 || name.indexOf('WOD — ') === 0) return;
        var k = moveKey_(name);
        if (!k) return;
        var d = r[1];
        var dateStr = Object.prototype.toString.call(d) === '[object Date]' ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d);
        var g = byKey[k] = byKey[k] || { names: {}, rows: [] };
        g.names[name] = (g.names[name] || 0) + 1;
        g.rows.push({ sid: String(r[0]), set: String(r[5]), date: dateStr, weight: r[6], reps: r[7], exFeel: r[9], note: r[8] || r[10] || '', sessNote: r[12] || '' });
      });
    }
    const label_ = function (g) { return Object.keys(g.names).sort(function (a, b) { return g.names[b] - g.names[a]; })[0]; };
    const dateCount_ = function (g) { var s = {}; g.rows.forEach(function (x) { s[x.date] = 1; }); return Object.keys(s).length; };

    var histText = exList.map(function (ex) {
      var k = moveKey_(ex.name);
      var own = byKey[k];
      var lines = own ? historyLines_(own.rows, 4, '  ') : [];
      var out = ex.name + ' — target ' + ex.sets + ' sets:\n' + (lines.length ? lines.join('\n') : '  no history of its own');
      if (!own || dateCount_(own) < 2) {
        var toks = moveTokens_(ex.name);
        var sims = Object.keys(byKey).filter(function (k2) { return k2 !== k; }).map(function (k2) {
          return { g: byKey[k2], score: moveSimilarity_(toks, k2.split(' ')) };
        }).filter(function (s) { return s.score >= SIMILAR_MIN; })
          .sort(function (a, b) { return b.score - a.score || dateCount_(b.g) - dateCount_(a.g); })
          .slice(0, 2);
        sims.forEach(function (s) {
          out += '\n  similar movement "' + label_(s.g) + '" (a different exercise; use only as a guide):\n'
            + historyLines_(s.g.rows, 3, '    ').join('\n');
        });
      }
      return out;
    }).join('\n\n');

    const system = 'You are the strength coach for Stephen "Toast" Loper, a hybrid athlete (Australian Rules footballer, half-marathoner, lifter) in a gentle cut who is carefully reintroducing lower-body work after a knee issue. '
      + 'Recommend weights and reps for today\'s workout using the recent history, how sessions felt (feel 1=rough to 5=peak), AND his logged notes. '
      + 'Weigh the notes heavily: they carry the real cues (pain, knee or tendon sensations, "start at X next time", form, "felt different"). If a note says to start at a specific number, use it. If a note flags pain, swelling, or a knee/tendon sensation, hold or reduce rather than progress. '
      + 'When the workout name includes "Reintroduction", "Bulletproofing", or "Prehab" (or the exercise is clearly rehab/prehab), progress in small cautious steps and prioritize staying pain-free and controlled over adding load. '
      + 'General progression otherwise: hit cleanly at feel 4-5 -> nudge up; feel 3 -> small bump or repeat; feel 1-2 or missed reps or a pain note -> hold or slightly reduce. '
      + 'If an exercise has little or no history of its own but similar movements are listed, base the recommendation on them and adjust for the differences (single-arm vs two-arm, cable vs dumbbell, incline vs flat, unilateral vs bilateral), and say so briefly in "why". '
      + 'Only when there is neither its own history nor a similar movement, give a sensible conservative starting point and say so briefly. '
      + 'ALL WEIGHTS ARE IN POUNDS (lb), never kilograms. Progress in real plate increments: about 5 lb for main lifts, 2.5 lb for small or rehab bumps. Never use em dashes in the "why". '
      + 'Return STRICT JSON ONLY, no prose, exactly: '
      + '{"recommendations":[{"name":"<exact name given>","sets":[{"weight":"<e.g. 285 or BW>","reps":"<e.g. 5>"}],"why":"<one short sentence>"}]}. '
      + 'Use the exact exercise names provided and give the indicated number of sets per exercise (ramping sets are fine).';
    const user = 'Today\'s workout: ' + workout + '\n\nRecent history and notes:\n' + histText + '\n\nReturn strict JSON.';

    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: RECOMMEND_MODEL, max_tokens: 3000, system: system, messages: [{ role: 'user', content: user }] }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      return json({ ok: false, error: 'Claude API ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 200) });
    }
    const data = JSON.parse(resp.getContentText());
    var textOut = '';
    if (data && data.content) {
      for (var ci = 0; ci < data.content.length; ci++) {
        var blk = data.content[ci];
        if (blk && blk.type === 'text' && blk.text) textOut += blk.text;
      }
    }
    var m = textOut.match(/\{[\s\S]*\}/);
    if (!m) return json({ ok: false, error: 'No JSON (stop=' + (data.stop_reason || '?') + '): ' + (textOut || JSON.stringify(data)).slice(0, 160) });
    var parsed = JSON.parse(m[0]);
    return json({ ok: true, recommendations: parsed.recommendations || [] });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function removeSessionRows(sh, sid) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(sid)) sh.deleteRow(i + 2);
  }
}

function flatten(s) {
  const out = [];
  const stamp = new Date();
  const id = s.id || '';
  let sFirst = true;
  (s.exercises || []).forEach(function (e) {
    const done = (e.sets || []).filter(function (st) { return st.weight || st.reps || st.note; });
    let eFirst = true;
    done.forEach(function (st, i) {
      out.push([
        id, s.date, s.type, e.name, e.target || '', i + 1,
        st.weight || '', st.reps || '', st.note || '',
        eFirst ? (e.feel || '') : '', eFirst ? (e.note || '') : '',
        sFirst ? (s.feel || '') : '', sFirst ? (s.notes || '') : '', stamp
      ]);
      eFirst = false; sFirst = false;
    });
  });
  if (s.finisher) {
    const f = s.finisher;
    const rounds = (f.rounds && f.rounds.length) ? f.rounds : [{ round: 1, push: f.totalPush, pull: f.totalPull }];
    let fFirst = true;
    rounds.forEach(function (r) {
      out.push([
        id, s.date, s.type, 'Push/Pull Ladder (finisher)', 'descending', r.round,
        r.push, r.pull, fFirst ? ('duration ' + fmt(f.seconds || 0)) : '',
        '', '', sFirst ? (s.feel || '') : '', sFirst ? (s.notes || '') : '', stamp
      ]);
      fFirst = false; sFirst = false;
    });
  }
  return out;
}

// =====================================================================
//  RUNS — import HealthFit CSVs from Drive into the "Runs" tab
// =====================================================================

// Run this ONCE from the editor. Creates the hourly trigger and does a first import.
function installRunTrigger() {
  const existing = ScriptApp.getProjectTriggers();
  const already = existing.some(function (t) { return t.getHandlerFunction() === 'importRuns'; });
  if (!already) {
    ScriptApp.newTrigger('importRuns').timeBased().everyHours(1).create();
  }
  const n = importRuns();
  Logger.log('Trigger ready. Imported ' + n + ' run rows on first pass.');
}

function importRuns() {
  const folder = DriveApp.getFolderById(RUN_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(RUN_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(RUN_SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(RUN_HEADER);

  const files = folder.getFiles();
  let imported = 0;
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (!/\.csv$/i.test(name)) continue;
    const key = 'run_done_' + f.getId();
    if (props.getProperty(key)) continue;
    try {
      const rows = parseRunCsv(f.getBlob().getDataAsString(), name);
      if (rows.length) {
        sh.getRange(sh.getLastRow() + 1, 1, rows.length, RUN_HEADER.length).setValues(rows);
        imported += rows.length;
      }
      props.setProperty(key, '1'); // mark done so it never re-imports
    } catch (err) {
      Logger.log('Run import failed for ' + name + ': ' + err); // leave unmarked to retry
    }
  }
  return imported;
}

const MI = 1609.344;
function _mmss(s) { s = Math.round(s); return _pad(Math.floor(s / 60)) + ':' + _pad(s % 60); }
function _hms(s) { s = Math.round(s); return Math.floor(s / 3600) + ':' + _pad(Math.floor((s % 3600) / 60)) + ':' + _pad(s % 60); }
function _pad(n) { return (n < 10 ? '0' : '') + n; }
function _pace(sec, mi) { return mi > 0 ? _mmss(sec / mi) : ''; }
function _disp(m) { return m >= MI * 0.95 ? (m / MI).toFixed(2) + ' mi' : (Math.round(m / 10) * 10) + ' m'; }
function _num(x) { const v = parseFloat(x); return isNaN(v) ? null : v; }
function _avg(a) { return a.length ? a.reduce(function (s, x) { return s + x; }, 0) / a.length : null; }
function _hr(v) { return v ? Math.round(v) + '' : ''; }
function _cad(v) { return v ? Math.round(v * 2) + '' : ''; } // CSV cadence is per-leg; x2 = spm

// Parse one HealthFit per-second CSV into Runs-tab rows (summary + intervals + splits).
// Reads columns by NAME, so it tolerates files that are missing Power / Intensity / etc.
function parseRunCsv(content, name) {
  const table = Utilities.parseCsv(content);
  if (!table || table.length < 2) return [];
  const head = table[0].map(function (h) { return String(h).trim(); });
  const idx = {};
  head.forEach(function (h, i) { idx[h] = i; });
  function cell(row, colName) {
    return (colName in idx) ? _num(row[idx[colName]]) : null;
  }
  function raw(row, colName) {
    return (colName in idx) ? String(row[idx[colName]]).trim() : '';
  }
  const hasIntensity = 'Intensity' in idx;
  const data = table.slice(1);

  // run_id / date / type from filename: YYYY-MM-DD-HHMMSS-<Type ...>-<device>.csv
  const m = String(name).match(/(\d{4}-\d{2}-\d{2})-(\d{6})-(.+?)-[^-]*\.csv$/i);
  const date = m ? m[1] : '';
  const rid = m ? (m[1] + '-' + m[2]) : name.replace(/\.csv$/i, '');
  const rtype = m ? m[3].trim() : 'Run';
  const stamp = new Date();

  const out = [];
  function add(section, seq, label, distM, dur, pc, hrV, pwV, cadV) {
    out.push([
      rid, date, rtype, section, seq, label,
      distM == null ? '' : _disp(distM), dur, pc,
      _hr(hrV), _hr(pwV), _cad(cadV), stamp
    ]);
  }

  // ---- totals for summary
  let totM = 0, totS = 0;
  const allHr = [], allPw = [], allCd = [];
  data.forEach(function (r) {
    const d = cell(r, 'Distance (meter)'); if (d != null && d > totM) totM = d;
    const s = cell(r, 'Since start (second)'); if (s != null && s > totS) totS = s;
    const h = cell(r, 'Heart Rate (bpm)'); if (h) allHr.push(h);
    const p = cell(r, 'Power (watt)'); if (p) allPw.push(p);
    const c = cell(r, 'Cadence (count/min)'); if (c) allCd.push(c);
  });
  const mi = totM / MI;
  add('summary', '', rtype, totM, _hms(totS), _pace(totS, mi), _avg(allHr), _avg(allPw), _avg(allCd));

  // ---- intervals by Lap (labelled from Intensity when present)
  const laps = {}; const order = [];
  data.forEach(function (r) {
    const lap = raw(r, 'Lap'); if (lap === '') return;
    if (!laps[lap]) { laps[lap] = { inten: hasIntensity ? raw(r, 'Intensity') : '', dmin: null, dmax: null, smin: null, smax: null, hr: [], pw: [], cd: [] }; order.push(lap); }
    const L = laps[lap];
    const d = cell(r, 'Distance (meter)'), s = cell(r, 'Since start (second)');
    const h = cell(r, 'Heart Rate (bpm)'), p = cell(r, 'Power (watt)'), c = cell(r, 'Cadence (count/min)');
    if (d != null) { L.dmin = L.dmin == null ? d : Math.min(L.dmin, d); L.dmax = L.dmax == null ? d : Math.max(L.dmax, d); }
    if (s != null) { L.smin = L.smin == null ? s : Math.min(L.smin, s); L.smax = L.smax == null ? s : Math.max(L.smax, s); }
    if (h) L.hr.push(h); if (p) L.pw.push(p); if (c) L.cd.push(c);
  });
  if (order.length > 1) {
    order.forEach(function (lap, i) {
      const L = laps[lap];
      const dm = (L.dmax - L.dmin), ss = (L.smax - L.smin);
      add('interval', i + 1, L.inten || ('lap ' + lap), dm, _mmss(ss), _pace(ss, dm / MI), _avg(L.hr), _avg(L.pw), _avg(L.cd));
    });
  }

  // ---- splits per mile (computed from cumulative distance + time)
  let cur = 1, last = 0; let H = [], P = [], C = [], lastD = 0, lastT = 0;
  data.forEach(function (r) {
    const d = cell(r, 'Distance (meter)'), s = cell(r, 'Since start (second)');
    if (d == null || s == null) return;
    lastD = d; lastT = s;
    const h = cell(r, 'Heart Rate (bpm)'), p = cell(r, 'Power (watt)'), c = cell(r, 'Cadence (count/min)');
    if (h) H.push(h); if (p) P.push(p); if (c) C.push(c);
    if (d >= cur * MI) {
      add('split', cur, 'Mile ' + cur, MI, _mmss(s - last), _pace(s - last, 1), _avg(H), _avg(P), _avg(C));
      last = s; cur += 1; H = []; P = []; C = [];
    }
  });
  const frac = (lastD - (cur - 1) * MI) / MI;
  if (frac > 0.05) {
    add('split', cur, 'Mile ' + cur + ' (' + frac.toFixed(2) + ')', lastD - (cur - 1) * MI, _mmss(lastT - last), _pace(lastT - last, frac), _avg(H), _avg(P), _avg(C));
  }

  return out;
}

// =====================================================================
function fmt(t) {
  const m = Math.floor(t / 60), s = t % 60;
  return (m < 10 ? '0' : '') + m + ':' + (s < 10 ? '0' : '') + s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
