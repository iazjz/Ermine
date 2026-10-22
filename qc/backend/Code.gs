/**
 * Ermine QC backend (Google Apps Script, bound to a Google Sheet)
 * سكربت الخلفية لفحص الجودة اليومي: يستقبل التقارير ويحفظها في الشيت.
 *
 * Setup: see qc/README.md. Change the two codes below, run setup() once,
 * then Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 */
var TEAM_CODE = 'CHANGE-ME-TEAM';        // الرمز اللي يكتبه الموظفون مرة واحدة على جهازهم
var MANAGER_CODE = 'CHANGE-ME-MANAGER';  // الرمز اللي يفتح شاشة السجل والتصدير
var TZ = 'Asia/Riyadh';

var REPORT_HEADERS = ['id', 'submitted_at', 'date', 'shift', 'employee', 'workers', 'done', 'total', 'issues', 'notes', 'lang', 'items_json'];
var ITEM_HEADERS = ['report_id', 'date', 'time', 'shift', 'employee', 'section', 'item_ar', 'item_en', 'value', 'limit', 'status', 'note'];
var CONFIG_HEADERS = ['kind', 'id', 'ar', 'en', 'type', 'limit', 'hide'];

var DEFAULT_CONFIG = [
  ['unit', 'u1', 'ثلاجة التحضير', 'Prep fridge', 'fridge', 5],
  ['unit', 'u2', 'ثلاجة الألبان والكريمة', 'Dairy & cream fridge', 'fridge', 5],
  ['unit', 'u3', 'ثلاجة المنتج النهائي', 'Finished product fridge', 'fridge', 5],
  ['unit', 'u4', 'مجمد 1', 'Freezer 1', 'freezer', -18],
  ['unit', 'u5', 'مجمد المخزن', 'Storage freezer', 'freezer', -18],
  ['zone', 'z1', 'منطقة الاستلام', 'Receiving area', '', ''],
  ['zone', 'z2', 'منطقة التحضير', 'Prep area', '', ''],
  ['zone', 'z3', 'منطقة الأفران والخبز', 'Oven & baking area', '', ''],
  ['zone', 'z4', 'منطقة التعبئة والتغليف', 'Packing area', '', ''],
  ['zone', 'z5', 'المخزن الجاف', 'Dry storage', '', ''],
  ['zone', 'z6', 'الأرضيات والمصارف', 'Floors & drains', '', ''],
  ['zone', 'z7', 'الأسطح والأدوات والمعدات', 'Surfaces, tools & equipment', '', ''],
  ['zone', 'z8', 'دورات المياه ومغاسل الأيدي', 'Restrooms & handwash stations', '', ''],
  ['hygiene', 'h1', 'الزي موحّد ونظيف', 'Uniform clean and complete', '', ''],
  ['hygiene', 'h2', 'غطاء الشعر والقفازات والكمامة', 'Hair cover, gloves and mask', '', ''],
  ['hygiene', 'h3', 'أظافر قصيرة ونظيفة وبدون مجوهرات', 'Short clean nails, no jewelry', '', ''],
  ['hygiene', 'h4', 'غسل الأيدي عند الدخول', 'Handwashing on entry', '', ''],
  ['hygiene', 'h5', 'لا توجد أعراض مرضية على العمال', 'No illness symptoms among workers', '', '']
];

/** Run once from the editor: creates the tabs and asks for permission. */
function setup() {
  ensureSheets_();
}

function doGet() {
  return ContentService.createTextOutput('Ermine QC backend is running.');
}

function doPost(e) {
  var out;
  try {
    var body = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var req = JSON.parse(body);
    if (setupIncomplete_()) throw new Error('setup_incomplete');
    out = route_(req);
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  switch (req.action) {
    case 'signIn':
      return signIn_(req.employee, req.code);
    case 'config':
      requireEmployee_(req.employee, req.code);
      return { ok: true, config: readConfig_() };
    case 'who':
      /* The sign-in screen needs the list of names BEFORE anyone has proved who they
       * are, so this one action is open. It returns names and job titles only - never a
       * code, and never which checklist anyone fills. A name on a factory door is not a
       * secret; the code beside it is. */
      return { ok: true, people: whoList_() };
    case 'doneToday':
      requireEmployee_(req.employee, req.code);
      return { ok: true, done: doneToday_() };
    case 'submit':
      requireEmployee_(req.report && req.report.employeeId, req.code);
      return submit_(req.report);
    case 'list':
      requireCode_(req.managerCode, MANAGER_CODE);
      return list_(req);
    case 'otGet':
      requireCode_(req.managerCode, MANAGER_CODE);
      return { ok: true, record: otGet_(req.date) };
    case 'otDecide':
      requireCode_(req.managerCode, MANAGER_CODE);
      return otDecide_(req);
    case 'tasksList':
      requireCode_(req.managerCode, MANAGER_CODE);
      return { ok: true, tasks: tasksList_(req) };
    case 'tasksSave':
      requireCode_(req.managerCode, MANAGER_CODE);
      return tasksSave_(req.task);
    case 'tasksDone':
      requireCode_(req.managerCode, MANAGER_CODE);
      return tasksDone_(req.id, req.done);
    default:
      throw new Error('unknown_action');
  }
}

function setupIncomplete_() {
  return TEAM_CODE.length < 6 || MANAGER_CODE.length < 6 ||
    TEAM_CODE.indexOf('CHANGE-ME') === 0 || MANAGER_CODE.indexOf('CHANGE-ME') === 0;
}

function requireCode_(given, expected) {
  if (String(given || '') !== expected) throw new Error('bad_code');
}

/* Every person has their own code, held in column I of Config and never sent to any
 * browser. The client posts an id and a code; the server answers yes or no. That is the
 * whole point of doing it here rather than in the page: a code that is never downloaded
 * cannot be read out of the page by the next person to pick up the tablet.
 *
 * TEAM_CODE still works as a master key so a forgotten code never stops the day. */
function findPerson_(id) {
  if (!id) return null;
  var sh = ensureSheets_().getSheetByName('Config');
  var last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, 9).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[0]).trim().toLowerCase() !== 'employee') continue;
    if (String(r[1]).trim() !== String(id).trim()) continue;
    var hide = r[6] === true || String(r[6]).toUpperCase() === 'TRUE';
    if (hide) return null;
    return {
      id: String(r[1]).trim(),
      ar: String(r[2]).trim() || String(r[3]).trim(),
      en: String(r[3]).trim() || String(r[2]).trim(),
      title: String(r[4] || '').trim(),
      role: String(r[7] || '').trim(),
      code: String(r[8] == null ? '' : r[8]).trim()
    };
  }
  return null;
}

function requireEmployee_(id, code) {
  var given = String(code || '');
  var p = findPerson_(id);
  if (!p) throw new Error('unknown_employee');
  if (given && given === TEAM_CODE) return p;
  if (!p.code) throw new Error('no_code_set');
  if (given !== p.code) throw new Error('bad_code');
  return p;
}

function signIn_(id, code) {
  var p = requireEmployee_(id, code);
  return {
    ok: true,
    me: { id: p.id, ar: p.ar, en: p.en, title: p.title, role: p.role },
    config: readConfig_()
  };
}

/* Who has already handed in a checklist today, by role.
 *
 * A checklist belongs to a job, not to a person: two CDPs on the same shift share one
 * list. Without this they either both fill it - two records of one round, and no way to
 * know which is true - or each assumes the other did and nobody does. So the app tells
 * you, before you start, whether your role is already covered for this shift. It does
 * not stop you: a second opinion is sometimes exactly what is wanted. It only removes
 * the excuse of not knowing.
 */
function doneToday_() {
  var ss = ensureSheets_();
  var rep = ss.getSheetByName('Reports');
  var last = rep.getLastRow();
  if (last < 2) return [];
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var rows = rep.getRange(2, 1, last - 1, REPORT_HEADERS.length).getValues();
  var people = {};
  (readConfig_().employees || []).forEach(function (e) {
    people[String(e.ar).trim()] = e; people[String(e.en).trim()] = e;
  });
  var out = [];
  rows.forEach(function (r) {
    if (dstr_(r[2]) !== today) return;
    var who = String(r[4] || '').trim();
    var p = people[who];
    out.push({ employee: who, shift: String(r[3] || ''),
               role: p ? p.role : '', title: p ? p.title : '',
               at: String(r[1] || '').slice(11, 16) });
  });
  return out;
}

function whoList_() {
  var sh = ensureSheets_().getSheetByName('Config');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 9).getValues();
  var out = [];
  rows.forEach(function (r) {
    if (String(r[0]).trim().toLowerCase() !== 'employee') return;
    var hide = r[6] === true || String(r[6]).toUpperCase() === 'TRUE';
    if (hide) return;
    var ar = String(r[2]).trim(), en = String(r[3]).trim();
    if (!ar && !en) return;
    out.push({
      id: String(r[1]).trim(),
      ar: ar || en,
      en: en || ar,
      title: String(r[4] || '').trim(),
      hasCode: !!String(r[8] == null ? '' : r[8]).trim()
    });
  });
  return out;
}

/* ---------- sheets ---------- */

function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var made = false;

  var rep = ss.getSheetByName('Reports');
  if (!rep) {
    rep = ss.insertSheet('Reports'); made = true;
    rep.appendRow(REPORT_HEADERS);
    rep.setFrozenRows(1);
    rep.getRange(1, 1, 1, REPORT_HEADERS.length).setFontWeight('bold');
  }
  // Asserted every time, not only at creation. Applying it once left any sheet that
  // predates the rule - or that somebody reformatted - storing real dates again.
  rep.getRange('C:D').setNumberFormat('@');
  var items = ss.getSheetByName('Items');
  if (!items) {
    items = ss.insertSheet('Items'); made = true;
    items.appendRow(ITEM_HEADERS);
    items.setFrozenRows(1);
    items.getRange(1, 1, 1, ITEM_HEADERS.length).setFontWeight('bold');
    items.getRange('B:C').setNumberFormat('@');
  }
  var cfg = ss.getSheetByName('Config');
  if (!cfg) {
    cfg = ss.insertSheet('Config'); made = true;
    cfg.appendRow(CONFIG_HEADERS);
    cfg.setFrozenRows(1);
    cfg.getRange(1, 1, 1, CONFIG_HEADERS.length).setFontWeight('bold');
    cfg.getRange(2, 1, DEFAULT_CONFIG.length, 6).setValues(DEFAULT_CONFIG);
    cfg.getRange(2, 7, 300, 1).insertCheckboxes();
    cfg.getRange('I1').setValue('code');
  cfg.getRange('H1').setValue('kind: employee | unit | zone | hygiene.  type: fridge | freezer.  limit: max °C.  hide: tick to hide a row.');
  }
  if (made) {
    ['Sheet1', 'ورقة1'].forEach(function (n) {
      var s = ss.getSheetByName(n);
      if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
    });
  }
  return ss;
}

function readConfig_() {
  var ss = ensureSheets_();
  var sh = ss.getSheetByName('Config');
  var cfg = { employees: [], units: [], zones: [], hyg: [], checks: [], roles: [], sections: [] };
  var last = sh.getLastRow();
  if (last < 2) return cfg;
  var rows = sh.getRange(2, 1, last - 1, 9).getValues();   // 8 = roles, 9 = personal code
  rows.forEach(function (r, i) {
    var kind = String(r[0]).trim().toLowerCase();
    if (!kind) return;
    var id = String(r[1]).trim();
    if (!id) {
      id = kind + '-' + (i + 2);
      sh.getRange(i + 2, 2).setValue(id);
    }
    var hide = r[6] === true || String(r[6]).toUpperCase() === 'TRUE';
    if (hide) return;
    var ar = String(r[2]).trim();
    var en = String(r[3]).trim();
    if (!ar && !en) return;
    var item = { id: id, ar: ar || en, en: en || ar };
    if (kind === 'check') {
      // The section rides in the `type` column and the roles in a new 8th column, so a
      // sheet people already have open did not need restructuring.
      item.section = String(r[4] || '').trim();
      item.roles = String(r[7] || '').trim();
      cfg.checks.push(item);
    } else if (kind === 'role') {
      cfg.roles.push(item);
    } else if (kind === 'section') {
      cfg.sections.push(item);
    } else if (kind === 'employee') {
      item.title = String(r[4] || '').trim();
      item.role = String(r[7] || '').trim();
      // column 9 holds their code and is deliberately NOT copied into the reply
      cfg.employees.push(item);
    } else if (kind === 'unit') {
      var type = String(r[4]).trim().toLowerCase() === 'freezer' ? 'freezer' : 'fridge';
      var lim = parseFloat(r[5]);
      item.type = type;
      item.limit = isNaN(lim) ? (type === 'freezer' ? -18 : 5) : lim;
      cfg.units.push(item);
    } else if (kind === 'zone') {
      cfg.zones.push(item);
    } else if (kind === 'hygiene') {
      cfg.hyg.push(item);
    }
  });
  return cfg;
}

/* ---------- submit ---------- */

function clip_(v, n) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);
}

// Stops a typed note from being read as a formula when it opens in Sheets or Excel.
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function cleanStatus_(s) {
  return s === 'ok' || s === 'bad' ? s : '';
}

function numOrBlank_(v) {
  var n = parseFloat(v);
  return isNaN(n) ? '' : n;
}

function submit_(rep) {
  if (!rep || typeof rep !== 'object' || !rep.items) throw new Error('invalid_report');
  var id = String(rep.id || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) throw new Error('invalid_report');
  var employee = clip_(rep.employee, 80);
  if (!employee) throw new Error('invalid_report');

  var now = new Date();
  var date = /^\d{4}-\d{2}-\d{2}$/.test(rep.date) ? rep.date : Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  var time = /^\d{2}:\d{2}$/.test(rep.time) ? rep.time : Utilities.formatDate(now, TZ, 'HH:mm');
  var shift = rep.shift === 'evening' ? 'evening' : 'morning';

  var sections = [
    ['temperature', 'units', rep.items.units],
    ['cleanliness', 'zones', rep.items.zones],
    ['hygiene', 'hyg', rep.items.hyg],
    ['checklist', 'checks', rep.items.checks]
  ];
  var clean = { units: [], zones: [], hyg: [], checks: [] };
  var itemRows = [];
  var done = 0, total = 0, issues = 0;

  sections.forEach(function (s) {
    var list = Array.isArray(s[2]) ? s[2].slice(0, 200) : [];
    list.forEach(function (it) {
      var status = cleanStatus_(it.status);
      var temp = s[0] === 'temperature' ? numOrBlank_(it.temp) : '';
      var limit = s[0] === 'temperature' ? numOrBlank_(it.limit) : '';
      var note = clip_(s[0] === 'temperature' ? it.action : it.note, 300);
      var ar = clip_(it.ar, 120), en = clip_(it.en, 120);
      total++;
      if (status) done++;
      if (status === 'bad') issues++;
      clean[s[1]].push({
        id: clip_(it.id, 40), ar: ar, en: en, type: s[0] === 'temperature' ? (it.type === 'freezer' ? 'freezer' : 'fridge') : undefined,
        limit: limit === '' ? null : limit, temp: temp === '' ? null : temp, status: status || null, note: note
      });
      itemRows.push([id, date, time, shift, safe_(employee), s[0], safe_(ar), safe_(en), temp, limit, status, safe_(note)]);
    });
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = ensureSheets_();
    var rs = ss.getSheetByName('Reports');
    if (rs.getLastRow() > 1) {
      var hit = rs.getRange(2, 1, rs.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
      if (hit) return { ok: true, duplicate: true };
    }
    var workers = numOrBlank_(rep.workers);
    rs.appendRow([
      id, Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"), date, shift, safe_(employee), workers,
      done, total, issues, safe_(clip_(rep.notes, 1000)), rep.lang === 'en' ? 'en' : 'ar', JSON.stringify(clean)
    ]);
    if (itemRows.length) {
      var is = ss.getSheetByName('Items');
      is.getRange(is.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- history ---------- */

/* A date column that is sometimes text and sometimes a real date.
 *
 * Dates are written as TEXT on purpose - a sheet that reads 10-11 as 10 November in one
 * locale and 11 October in another is a sheet that moves your deadlines. But the format
 * was only applied when the sheet was first created, so a row written before that rule,
 * or after somebody reformatted the column, comes back as a Date object. String() then
 * gives "Wed Oct 21 2026 00:00:00 GMT+0300", which falls outside every yyyy-mm-dd range
 * it is compared against - and the row vanishes from its own history.
 *
 * Read the value; do not assume its type.
 */
function dstr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).slice(0, 10);
}

function list_(req) {
  var from = /^\d{4}-\d{2}-\d{2}$/.test(req.from) ? req.from : '0000-00-00';
  var to = /^\d{4}-\d{2}-\d{2}$/.test(req.to) ? req.to : '9999-99-99';
  var ss = ensureSheets_();
  var rs = ss.getSheetByName('Reports');
  var out = [];
  var last = rs.getLastRow();
  if (last > 1) {
    var rows = rs.getRange(2, 1, last - 1, REPORT_HEADERS.length).getValues();
    rows.forEach(function (r) {
      var d = dstr_(r[2]);
      if (d < from || d > to) return;
      var items = {};
      try { items = JSON.parse(r[11]); } catch (e) { items = {}; }
      out.push({
        id: String(r[0]), submittedAt: String(r[1]), date: d, shift: String(r[3]), employee: String(r[4]),
        workers: r[5], done: r[6], total: r[7], issues: r[8], notes: String(r[9]), lang: String(r[10]), items: items
      });
    });
  }
  out.sort(function (a, b) { return a.submittedAt < b.submittedAt ? 1 : a.submittedAt > b.submittedAt ? -1 : 0; });
  return { ok: true, reports: out.slice(0, 500), sheetUrl: ss.getUrl() };
}

/* ---------- tasks ---------- *
 * The manager's task list, kept in a Tasks tab of this same sheet so there is
 * one store and one code rather than two of each. Ticking a task archives it;
 * a repeating task rolls to its next date instead.
 */

/* ---------- overtime, asked for and answered ----------------------------------
 * The plan is built in the morning and some days it does not fit a standard shift.
 * Until now that simply happened: the hours were worked and the cost turned up at the
 * end of the month. This makes it a question with a name on it - one row per day, what
 * was decided, and when.
 *
 * The REQUEST itself is not stored. It is read from the workbook each time, because
 * the plan changes during the morning and a stored copy would go stale the moment the
 * chef edits a line. Only the DECISION is kept, because a decision is a fact about a
 * person and has to survive.
 */
var OT_HEADERS = ['date', 'extra_staff_hours', 'extra_cost', 'status',
                  'decided_at', 'note'];

function otSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Overtime');
  if (!sh) {
    sh = ss.insertSheet('Overtime');
    sh.appendRow(OT_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, OT_HEADERS.length).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
    sh.setColumnWidth(6, 320);
  }
  return sh;
}

function otRow_(date) {
  var sh = otSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { sh: sh, row: 0, values: null };
  var vals = sh.getRange(2, 1, last - 1, OT_HEADERS.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).slice(0, 10) === date) {
      return { sh: sh, row: i + 2, values: vals[i] };
    }
  }
  return { sh: sh, row: 0, values: null };
}

function otGet_(date) {
  date = String(date || '').slice(0, 10);
  if (!date) return null;
  var r = otRow_(date);
  if (!r.row) return null;
  return { date: date, hours: r.values[1], cost: r.values[2],
           status: String(r.values[3] || ''), at: String(r.values[4] || ''),
           note: String(r.values[5] || '') };
}

function otDecide_(req) {
  var date = String(req.date || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('bad_date');
  var status = String(req.status || '');
  if (['approved', 'declined'].indexOf(status) < 0) throw new Error('bad_status');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var r = otRow_(date);
    var now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
    var row = [date, Number(req.hours) || 0, Number(req.cost) || 0,
               status, now, clip_(req.note, 300)];
    if (r.row) r.sh.getRange(r.row, 1, 1, OT_HEADERS.length).setValues([row]);
    else r.sh.appendRow(row);
    return { ok: true, record: otGet_(date) };
  } finally {
    lock.releaseLock();
  }
}

var TASK_HEADERS = ['id', 'created_at', 'title', 'owner', 'section', 'due_date',
                    'due_time', 'priority', 'status', 'repeat', 'notes',
                    'done_at', 'calendar_event'];

function tasksSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Tasks');
  if (!sh) {
    sh = ss.insertSheet('Tasks');
    sh.appendRow(TASK_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, TASK_HEADERS.length).setFontWeight('bold');
    // Dates and times kept as TEXT on purpose. A sheet that helpfully reinterprets
    // 10-11 as 10 November in one locale and 11 October in another is a sheet that
    // quietly moves your deadlines.
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('F:G').setNumberFormat('@');
    sh.getRange('L:L').setNumberFormat('@');
    sh.setColumnWidth(3, 320);
    sh.setColumnWidth(11, 280);
  }
  return sh;
}

function tasksRows_() {
  var sh = tasksSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { sh: sh, rows: [] };
  return { sh: sh, rows: sh.getRange(2, 1, last - 1, TASK_HEADERS.length).getValues() };
}

function taskObj_(r) {
  return {
    id: String(r[0]), createdAt: String(r[1]), title: String(r[2]),
    owner: String(r[3]), section: String(r[4]), due: String(r[5]),
    time: String(r[6]), priority: String(r[7]), status: String(r[8]),
    repeat: String(r[9]), notes: String(r[10]), doneAt: String(r[11]),
    event: String(r[12])
  };
}

/* Open tasks by default; the finished ones when asked. Ticking a task is what sends
   it to the archive - nothing is deleted and nothing changes sheet. */
function tasksList_(req) {
  var d = tasksRows_();
  var want = req && req.archived ? true : false;
  var out = [];
  d.rows.forEach(function (r) {
    if (!String(r[0])) return;
    var o = taskObj_(r);
    if ((o.status === 'Done') === want) out.push(o);
  });
  out.sort(function (a, b) {
    var ad = a.due || '9999-99-99', bd = b.due || '9999-99-99';
    if (ad !== bd) return ad < bd ? -1 : 1;
    return a.createdAt < b.createdAt ? 1 : -1;
  });
  return out;
}

function tasksSave_(task) {
  if (!task || !task.title || !String(task.title).trim()) throw new Error('no_title');
  var d = tasksRows_();
  var now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  var id = String(task.id || '').trim();
  var row = [
    id || ('K' + new Date().getTime().toString(36).toUpperCase()),
    now,
    clip_(task.title, 200),
    clip_(task.owner, 60),
    clip_(task.section, 60),
    /^\d{4}-\d{2}-\d{2}$/.test(task.due) ? task.due : '',
    /^\d{2}:\d{2}$/.test(task.time) ? task.time : '',
    ['High', 'Medium', 'Low'].indexOf(task.priority) >= 0 ? task.priority : 'Medium',
    ['To do', 'Doing', 'Waiting', 'Done'].indexOf(task.status) >= 0 ? task.status : 'To do',
    clip_(task.repeat, 20),
    clip_(task.notes, 600),
    '',
    clip_(task.event, 120)
  ];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (id) {
      for (var i = 0; i < d.rows.length; i++) {
        if (String(d.rows[i][0]) === id) {
          row[1] = String(d.rows[i][1]) || now;   // keep the original created_at
          row[11] = String(d.rows[i][11]);        // and any done_at already recorded
          d.sh.getRange(i + 2, 1, 1, TASK_HEADERS.length).setValues([row]);
          return { ok: true, task: taskObj_(row) };
        }
      }
    }
    d.sh.appendRow(row);
    return { ok: true, task: taskObj_(row) };
  } finally {
    lock.releaseLock();
  }
}

/* Tick a task off, or put it back.
   A repeating task does not close: its date moves to the next occurrence AFTER today
   and it returns to the open list, because "every week" means the next one is due
   however late this one was finished. done_at records when it was actually ticked. */
function tasksDone_(id, done) {
  if (!id) throw new Error('no_id');
  var d = tasksRows_();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    for (var i = 0; i < d.rows.length; i++) {
      if (String(d.rows[i][0]) !== String(id)) continue;
      var r = d.rows[i].slice();
      if (done === false) {
        r[8] = 'To do';
        r[11] = '';
      } else {
        var rep = String(r[9] || '').trim();
        var due = String(r[5] || '');
        var next = (rep && due) ? nextDue_(due, rep, today) : '';
        r[11] = today;
        if (next) { r[5] = next; r[8] = 'To do'; r[12] = ''; }
        else { r[8] = 'Done'; }
      }
      d.sh.getRange(i + 2, 1, 1, TASK_HEADERS.length).setValues([r]);
      return { ok: true, task: taskObj_(r) };
    }
    throw new Error('not_found');
  } finally {
    lock.releaseLock();
  }
}

function nextDue_(due, repeat, today) {
  var p = due.split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  var t = today.split('-');
  var now = new Date(Number(t[0]), Number(t[1]) - 1, Number(t[2]));
  var step = { 'Daily': 1, 'Weekly': 7, 'Every 2 weeks': 14 }[repeat];
  var guard = 0;
  if (step) {
    while (d <= now && guard++ < 2000) d.setDate(d.getDate() + step);
  } else if (repeat === 'Monthly' || repeat === 'Quarterly') {
    var n = (repeat === 'Monthly') ? 1 : 3;
    while (d <= now && guard++ < 400) {
      var day = d.getDate();
      d.setDate(1);
      d.setMonth(d.getMonth() + n);
      var last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
      d.setDate(Math.min(day, last));
    }
  } else { return ''; }
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}


/* ---------- the role-based checklist ---------- *
 * The five paper checklists - Chef de Partie, Head Chef, Storekeeper, Maintenance and
 * the manager's weekly - merged into one catalogue. Every item names the roles it
 * belongs to, so one list of 127 serves five people without any of them seeing the
 * other four's work. An item naming no role belongs to everyone, which is how the
 * original temperature, zone and hygiene items keep working untouched.
 *
 * Run seedChecklist() ONCE from the editor to write the catalogue into Config.
 * After that Config is the source: edit the wording there, tick `hide` to retire a
 * line, and change the roles column to move an item between people.
 */
var CHECK_ROLES = [
  ['cdp', 'شيف دي بارتي', 'Chef de Partie'],
  ['head', 'الشيف التنفيذي', 'Head Chef'],
  ['store', 'أمين المخزن', 'Storekeeper'],
  ['maint', 'الصيانة', 'Maintenance'],
  ['qpm', 'مدير الجودة والإنتاج', 'Quality & Production Manager']
];
var CHECK_SECTIONS = [
  ['before', 'قبل الوردية', 'Before shift'],
  ['after', 'بعد الوردية', 'After shift'],
  ['receiving', 'استلام المواد الخام', 'Raw material receiving'],
  ['cold', 'غرفة التبريد', 'Cold store'],
  ['freezer', 'المجمد', 'Freezer'],
  ['safety', 'السلامة والتحضير', 'Safety & preparation'],
  ['preventive', 'الصيانة الوقائية', 'Preventive maintenance'],
  ['facility', 'فحص المبنى', 'Facility checks'],
  ['weekly', 'المراجعة الأسبوعية', 'Weekly review']
];
var CHECK_CATALOGUE = [
  ['check', 'b01', 'النظافة الشخصية سليمة — اليدان والزي وغطاء الشعر', 'Personal hygiene OK — hands, uniform, hair covering', 'before', '', false, 'cdp,head'],
  ['check', 'b02', 'الحالة الصحية سليمة — لا مرض ولا جروح', 'Health condition OK — no illness or wounds', 'before', '', false, 'cdp,head,store'],
  ['check', 'b03', 'منطقة الإنتاج المكلَّف بها محدَّدة', 'Assigned production area confirmed', 'before', '', false, 'cdp'],
  ['check', 'b04', 'محطة العمل نظيفة وخالية من التلوث', 'Workstation clean and free from contamination', 'before', '', false, 'cdp'],
  ['check', 'b05', 'التحقق من نظافة الوردية السابقة', 'Previous cleaning verified', 'before', '', false, 'cdp,head'],
  ['check', 'b06', 'أدوات ومواد التنظيف متوفرة', 'Cleaning tools and chemicals available', 'before', '', false, 'cdp,head'],
  ['check', 'b07', 'خطة الإنتاج مُراجَعة', 'Production and batch plan reviewed', 'before', '', false, 'cdp,head'],
  ['check', 'b08', 'الوصفات المعتمدة متَّبَعة', 'Approved recipes followed', 'before', '', false, 'cdp'],
  ['check', 'b09', 'المواد الخام متوفرة وضمن الصلاحية', 'Raw materials available and within expiry', 'before', '', false, 'cdp,head'],
  ['check', 'b10', 'جودة المكونات مُتحقَّق منها', 'Ingredient quality verified', 'before', '', false, 'head'],
  ['check', 'b11', 'التغليف سليم ونظيف', 'Packaging intact and clean', 'before', '', false, 'cdp'],
  ['check', 'b12', 'المعدات والقوالب والأدوات نظيفة', 'Equipment, moulds and utensils clean', 'before', '', false, 'cdp'],
  ['check', 'b13', 'المعايرة مفحوصة بصريًا (إن وُجدت)', 'Calibration visually checked (if applicable)', 'before', '', false, 'cdp'],
  ['check', 'b14', 'الأفران وأجهزة التمبرنج جاهزة', 'Ovens and tempering machines ready', 'before', '', false, 'cdp'],
  ['check', 'b15', 'الملصقات والتواريخ وأكواد التشغيلة متوفرة', 'Labels, dates and batch codes available', 'before', '', false, 'cdp'],
  ['check', 'b16', 'الحضور والتغطية مؤكَّدان', 'Staff attendance and coverage confirmed', 'before', '', false, 'head'],
  ['check', 'b17', 'نقاط التحكم الحرجة والبرامج التمهيدية مُراجَعة مع الفريق', 'CCPs / PRPs / SSOPs reviewed with the team', 'before', '', false, 'head'],
  ['check', 'b18', 'المهام موزَّعة بوضوح', 'Tasks assigned clearly', 'before', '', false, 'head'],
  ['check', 'b19', 'حالة المعدات مؤكَّدة', 'Equipment condition confirmed', 'before', '', false, 'head'],
  ['check', 'b20', 'ملاحظات الصيانة الوقائية مُراجَعة', 'Preventive maintenance issues reviewed', 'before', '', false, 'head'],
  ['check', 'b21', 'سجلات حرارة التبريد والتجميد مُراجَعة', 'Chiller and freezer temperature logs reviewed', 'before', '', false, 'head'],
  ['check', 'b22', 'موازين الحرارة متوفرة', 'Thermometers available', 'before', '', false, 'head'],
  ['check', 'b23', 'الطلبات الخاصة والانحرافات مُراجَعة', 'Special orders or deviations reviewed', 'before', '', false, 'head'],
  ['check', 'a01', 'المنطقة نُظِّفت وعُقِّمت', 'Area cleaned and sanitised', 'after', '', false, 'cdp'],
  ['check', 'a02', 'المعدات والقوالب والصواني نُظِّفت وخُزِّنت', 'Equipment, moulds and trays cleaned and stored', 'after', '', false, 'cdp'],
  ['check', 'a03', 'المواد الكيميائية استُخدمت وخُزِّنت بشكل صحيح ومنفصل', 'Chemicals used correctly and stored separately', 'after', '', false, 'cdp'],
  ['check', 'a04', 'المنتجات مخزَّنة على الحرارة الصحيحة', 'Products stored at the correct temperature', 'after', '', false, 'cdp,store'],
  ['check', 'a05', 'مبدأ الوارد أولًا صادر أولًا مطبَّق', 'FIFO applied', 'after', '', false, 'cdp,store'],
  ['check', 'a06', 'الهدر والمرفوضات مسجَّلة', 'Waste and rejected items recorded', 'after', '', false, 'cdp,head'],
  ['check', 'a07', 'الأرضيات والمصارف نُظِّفت', 'Floors and drains cleaned', 'after', '', false, 'cdp'],
  ['check', 'a08', 'الآلات مُطفأة بأمان', 'Machines switched off safely', 'after', '', false, 'cdp,maint'],
  ['check', 'a09', 'التبريد والتجميد يعملان بشكل طبيعي', 'Chillers and freezers operating normally', 'after', '', false, 'cdp,store'],
  ['check', 'a10', 'الانحرافات مُبلَّغ عنها', 'Deviations reported', 'after', '', false, 'cdp,head,store,maint'],
  ['check', 'a11', 'الإجراءات التصحيحية مدوَّنة', 'Corrective actions noted', 'after', '', false, 'cdp,head,store'],
  ['check', 'a12', 'كميات الإنتاج مسجَّلة', 'Production quantities recorded', 'after', '', false, 'cdp'],
  ['check', 'a13', 'المنطقة جاهزة للوردية التالية', 'Area ready for the next shift', 'after', '', false, 'cdp'],
  ['check', 'a14', 'الإنتاج أُنجز كما هو مخطَّط', 'Production completed as planned', 'after', '', false, 'head'],
  ['check', 'a15', 'جودة المنتج النهائي مُتحقَّق منها', 'Finished product quality verified', 'after', '', false, 'head'],
  ['check', 'a16', 'مراقبة نقاط التحكم الحرجة مُراجَعة', 'CCP monitoring reviewed', 'after', '', false, 'head'],
  ['check', 'a17', 'سجلات الحرارة مُراجَعة وموقَّعة', 'Temperature logs reviewed and signed', 'after', '', false, 'head'],
  ['check', 'a18', 'فعالية التنظيف مُتحقَّق منها', 'Cleaning effectiveness verified', 'after', '', false, 'head'],
  ['check', 'a19', 'التخزين ووضع الملصقات والوارد أولًا مؤكَّدة', 'Storage, labelling and FIFO confirmed', 'after', '', false, 'head,store'],
  ['check', 'a20', 'الإجراءات التصحيحية والوقائية محدَّدة', 'Corrective and preventive actions defined', 'after', '', false, 'head'],
  ['check', 'a21', 'الهدر وإعادة التشغيل مُراجَعان', 'Waste and rework reviewed', 'after', '', false, 'head'],
  ['check', 'a22', 'إطفاء المعدات مُتحقَّق منه', 'Equipment shutdown verified', 'after', '', false, 'head,maint'],
  ['check', 'a23', 'التزام الفريق مُراجَع', 'Staff compliance reviewed', 'after', '', false, 'head'],
  ['check', 'a24', 'تعليمات الوردية التالية مُجهَّزة', 'Instructions prepared for the next shift', 'after', '', false, 'head'],
  ['check', 'a25', 'النظافة العامة والإنتاج معتمدان', 'Overall hygiene and production approved', 'after', '', false, 'head'],
  ['check', 'r01', 'مركبة التوصيل نظيفة ومناسبة', 'Delivery vehicle clean and suitable', 'receiving', '', false, 'store'],
  ['check', 'r02', 'المورّد معتمد', 'Supplier approved', 'receiving', '', false, 'store'],
  ['check', 'r03', 'مستندات التوصيل متوفرة', 'Delivery documents available', 'receiving', '', false, 'store'],
  ['check', 'r04', 'الصنف والكمية مطابقان لأمر الشراء', 'Product and quantity match the purchase order', 'receiving', '', false, 'store'],
  ['check', 'r05', 'التغليف سليم ونظيف', 'Packaging intact and clean', 'receiving', '', false, 'store'],
  ['check', 'r06', 'لا تلوث ولا أجسام غريبة', 'No contamination or foreign matter', 'receiving', '', false, 'store'],
  ['check', 'r07', 'تواريخ الصلاحية مقبولة', 'Expiry dates acceptable', 'receiving', '', false, 'store'],
  ['check', 'r08', 'حرارة المبرَّد مقبولة', 'Chilled temperature acceptable', 'receiving', '', false, 'store'],
  ['check', 'r09', 'حرارة المجمَّد مقبولة', 'Frozen temperature acceptable', 'receiving', '', false, 'store'],
  ['check', 'r10', 'البضائع الجافة خالية من الرطوبة', 'Dry goods free from moisture', 'receiving', '', false, 'store'],
  ['check', 'r11', 'المرفوضات معزولة', 'Rejected items isolated', 'receiving', '', false, 'store'],
  ['check', 'r12', 'المقبولات موسومة', 'Accepted items labelled', 'receiving', '', false, 'store'],
  ['check', 'r13', 'منطقة الاستلام نظيفة', 'Receiving area clean', 'receiving', '', false, 'store'],
  ['check', 'c01', 'غرفة التبريد نظيفة وبلا روائح', 'Chiller clean and odour-free', 'cold', '', false, 'store'],
  ['check', 'c02', 'أختام الأبواب سليمة', 'Door seals OK', 'cold', '', false, 'store'],
  ['check', 'c03', 'الأصناف مرفوعة عن الأرض وبعيدة عن الجدران', 'Items off the floor and away from walls', 'cold', '', false, 'store'],
  ['check', 'c04', 'الخام ونصف المصنَّع والنهائي مفصولة', 'Raw, semi-finished and finished separated', 'cold', '', false, 'store'],
  ['check', 'c05', 'الملصقات مكتملة', 'Labels complete', 'cold', '', false, 'store'],
  ['check', 'c06', 'لا أصناف منتهية الصلاحية', 'No expired items', 'cold', '', false, 'store'],
  ['check', 'c07', 'الحاويات مغطاة', 'Containers covered', 'cold', '', false, 'store'],
  ['check', 'c08', 'لا تحميل زائد', 'No overloading', 'cold', '', false, 'store'],
  ['check', 'c09', 'الانسكابات مُنظَّفة', 'Spills cleaned', 'cold', '', false, 'store'],
  ['check', 'c10', 'الباب مغلق', 'Door closed', 'cold', '', false, 'store'],
  ['check', 'f01', 'المجمد نظيف ومرتَّب', 'Freezer clean and organised', 'freezer', '', false, 'store'],
  ['check', 'f02', 'أختام الأبواب سليمة', 'Door seals OK', 'freezer', '', false, 'store'],
  ['check', 'f03', 'المنتجات مرصوصة بأمان', 'Products stacked safely', 'freezer', '', false, 'store'],
  ['check', 'f04', 'الأصناف موسومة', 'Items labelled', 'freezer', '', false, 'store'],
  ['check', 'f05', 'لا أصناف أُذيبت ثم جُمِّدت', 'No thawed or refrozen items', 'freezer', '', false, 'store'],
  ['check', 'f06', 'لا تغليف تالف', 'No damaged packaging', 'freezer', '', false, 'store'],
  ['check', 'f07', 'لا تراكم جليد', 'No ice build-up', 'freezer', '', false, 'store'],
  ['check', 'f08', 'المجمد غير محمَّل زيادة', 'Freezer not overloaded', 'freezer', '', false, 'store'],
  ['check', 'f09', 'الباب مغلق', 'Door closed', 'freezer', '', false, 'store'],
  ['check', 'f10', 'الأصناف المنتهية أو المرفوضة معزولة', 'Expired or rejected items segregated', 'freezer', '', false, 'store'],
  ['check', 'f11', 'الأصناف قريبة الانتهاء مُبلَّغ عنها', 'Near-expiry items communicated', 'freezer', '', false, 'store'],
  ['check', 'f12', 'المخزن مؤمَّن', 'Store secured', 'freezer', '', false, 'store'],
  ['check', 'm01', 'الأدوات متوفرة وبحالة جيدة', 'Tools available and in good condition', 'safety', '', false, 'maint'],
  ['check', 'm02', 'المنطقة مفحوصة من المخاطر', 'Area checked for hazards', 'safety', '', false, 'maint'],
  ['check', 'm03', 'لا أسلاك مكشوفة', 'No exposed wiring', 'safety', '', false, 'maint'],
  ['check', 'm04', 'المشرف مُبلَّغ قبل بدء العمل', 'Supervisor informed before work', 'safety', '', false, 'maint'],
  ['check', 'm05', 'الخلاطات تعمل بشكل طبيعي', 'Mixers working normally', 'preventive', '', false, 'maint'],
  ['check', 'm06', 'الأفران تعمل بشكل صحيح', 'Ovens functioning correctly', 'preventive', '', false, 'maint'],
  ['check', 'm07', 'أجهزة التمبرنج مستقرة', 'Tempering machines stable', 'preventive', '', false, 'maint'],
  ['check', 'm08', 'الثلاجات والمجمدات تعمل', 'Refrigerators and freezers operating', 'preventive', '', false, 'maint'],
  ['check', 'm09', 'المراوح والتهوية تعمل', 'Fans and ventilation working', 'preventive', '', false, 'maint'],
  ['check', 'm10', 'لا تسريبات ماء أو زيت أو غاز', 'No leaks — water, oil or gas', 'preventive', '', false, 'maint'],
  ['check', 'm11', 'الكابلات والقوابس سليمة', 'Cables and plugs intact', 'preventive', '', false, 'maint'],
  ['check', 'm12', 'الإضاءة تعمل', 'Lighting operational', 'facility', '', false, 'maint'],
  ['check', 'm13', 'إمداد المياه مستقر', 'Water supply stable', 'facility', '', false, 'maint'],
  ['check', 'm14', 'التصريف يجري', 'Drainage flowing', 'facility', '', false, 'maint'],
  ['check', 'm15', 'الأبواب والمفصلات تعمل', 'Doors and hinges functional', 'facility', '', false, 'maint'],
  ['check', 'm16', 'الرفوف ثابتة', 'Shelving stable', 'facility', '', false, 'maint'],
  ['check', 'm17', 'الجدران والأرضيات والأسقف بحالة جيدة', 'Walls, floors and ceilings in good condition', 'facility', '', false, 'maint'],
  ['check', 'm18', 'الشفاطات والتهوية تعمل', 'Exhaust and ventilation working', 'facility', '', false, 'maint'],
  ['check', 'm19', 'الأدوات نُظِّفت وخُزِّنت', 'Tools cleaned and stored', 'after', '', false, 'maint'],
  ['check', 'm20', 'المنطقة نُظِّفت', 'Area cleaned', 'after', '', false, 'maint'],
  ['check', 'm21', 'المعدات تُركت آمنة', 'Equipment left safe', 'after', '', false, 'maint'],
  ['check', 'm22', 'الملاحظات أُبلغت للمشرف', 'Issues communicated to the supervisor', 'after', '', false, 'maint'],
  ['check', 'w01', 'المخطَّط مقابل الفعلي مُراجَع', 'Plan versus actual reviewed', 'weekly', '', false, 'qpm'],
  ['check', 'w02', 'الانحرافات محدَّدة', 'Deviations identified', 'weekly', '', false, 'qpm'],
  ['check', 'w03', 'جودة المنتج مُتحقَّق منها', 'Product quality verified', 'weekly', '', false, 'qpm'],
  ['check', 'w04', 'الفاقد وإعادة التشغيل مُراجَعان', 'Losses and rework reviewed', 'weekly', '', false, 'qpm'],
  ['check', 'w05', 'الإجراءات التصحيحية محدَّدة', 'Corrective actions defined', 'weekly', '', false, 'qpm'],
  ['check', 'w06', 'قوائم الفحص اليومية مُراجَعة', 'Daily checklists reviewed', 'weekly', '', false, 'qpm'],
  ['check', 'w07', 'سجلات نقاط التحكم الحرجة مُراجَعة', 'CCP records reviewed', 'weekly', '', false, 'qpm'],
  ['check', 'w08', 'سجلات الحرارة مُتحقَّق منها', 'Temperature logs verified', 'weekly', '', false, 'qpm'],
  ['check', 'w09', 'سجلات التنظيف مُراجَعة', 'Cleaning records reviewed', 'weekly', '', false, 'qpm'],
  ['check', 'w10', 'الالتزام بالنظافة مُلاحَظ', 'Hygiene compliance observed', 'weekly', '', false, 'qpm'],
  ['check', 'w11', 'استخدام معدات الوقاية مُتحقَّق منه', 'PPE usage verified', 'weekly', '', false, 'qpm'],
  ['check', 'w12', 'احتياجات التدريب محدَّدة', 'Training needs identified', 'weekly', '', false, 'qpm'],
  ['check', 'w13', 'ظروف التخزين مُتحقَّق منها', 'Storage conditions verified', 'weekly', '', false, 'qpm'],
  ['check', 'w14', 'الوارد أولًا والأقرب انتهاءً أولًا متَّبعان', 'FIFO / FEFO followed', 'weekly', '', false, 'qpm'],
  ['check', 'w15', 'قريب الانتهاء مضبوط', 'Near-expiry controlled', 'weekly', '', false, 'qpm'],
  ['check', 'w16', 'الفصل بين الأصناف مؤكَّد', 'Segregation confirmed', 'weekly', '', false, 'qpm'],
  ['check', 'w17', 'سجلات الصيانة مُراجَعة', 'Maintenance logs reviewed', 'weekly', '', false, 'qpm'],
  ['check', 'w18', 'الملاحظات المفتوحة متابَعة', 'Outstanding issues tracked', 'weekly', '', false, 'qpm'],
  ['check', 'w19', 'حالة المبنى مقبولة', 'Facility condition acceptable', 'weekly', '', false, 'qpm'],
  ['check', 'w20', 'السجلات مكتملة', 'Records complete', 'weekly', '', false, 'qpm'],
  ['check', 'w21', 'موقَّعة ومؤرَّخة', 'Signed and dated', 'weekly', '', false, 'qpm'],
  ['check', 'w22', 'محفوظة بشكل صحيح', 'Filed correctly', 'weekly', '', false, 'qpm']
];


/* The people who actually fill a checklist, and the role each one defaults to.
 * Everybody else on the floor is hidden rather than deleted: a name that has signed a
 * report before must keep existing, or that report loses its author. */
/* id, Arabic name, English name, job title, which catalogue role they fill.
 * Title and role are different things on purpose: a Sous Chef and a Head Chef run the
 * same checks, and a Senior CDP runs the CDP's, so five job titles map onto three
 * checklists. Keeping them apart means a promotion changes a title without rewriting
 * anyone's checklist. */
var CHECK_PEOPLE = [
  ['emp-youssef', 'يوسف',  'Youssef',        'Head Chef',              'head'],
  ['emp-gihan',   'غيهان', 'Gihan Fernando', 'Sous Chef',              'head'],
  ['emp-nitish',  'نيتيش', 'Nitish Bisht',   'Senior CDP',             'cdp'],
  ['emp-mahesh',  'ماهيش', 'Mahesh Chanaka', 'CDP',                    'cdp'],
  ['emp-thamim',  'تميم',  'Thamim Ansari',  'CDP',                    'cdp'],
  ['emp-aziz',    'عزيز',  'Aziz',           'Operations Manager',     'qpm'],
  ['emp-ali',     'علي',   'Ali Nawaz',      'Storekeeper',            'store'],
  /* Asif Cheema, maintenance - NOT Sk Asif, who works in the kitchen and has no
   * checklist. Two different people whose first names match; the full name is
   * carried here so the row can never be read as the other man. */
  ['emp-asif',    'عاصف تشيما', 'Asif Cheema',  'Maintenance Technician', 'maint']
];

/** Put the seven on the list and hide the rest. Run from seedChecklist(). */
function seedPeople_(cfg) {
  var last = cfg.getLastRow();
  if (last < 2) return 0;
  var rows = cfg.getRange(2, 1, last - 1, 8).getValues();
  var keep = {}, touched = 0;
  CHECK_PEOPLE.forEach(function (p) { keep[p[0]] = p; });
  rows.forEach(function (r, i) {
    if (String(r[0]).trim().toLowerCase() !== 'employee') return;
    var id = String(r[1]).trim();
    var row = i + 2;
    if (keep[id]) {
      cfg.getRange(row, 3).setValue(keep[id][1]);
      cfg.getRange(row, 4).setValue(keep[id][2]);
      cfg.getRange(row, 5).setValue(keep[id][3]);   // job title
      cfg.getRange(row, 7).setValue(false);
      cfg.getRange(row, 8).setValue(keep[id][4]);   // checklist role
      delete keep[id];                               // column 9 is their code, left alone

    } else {
      cfg.getRange(row, 7).setValue(true);     // hidden, never deleted
    }
    touched++;
  });
  var add = [];
  CHECK_PEOPLE.forEach(function (p) {
    if (keep[p[0]]) add.push(['employee', p[0], p[1], p[2], p[3], '', false, p[4], '']);
  });
  if (add.length) cfg.getRange(cfg.getLastRow() + 1, 1, add.length, 9).setValues(add);
  if (cfg.getRange(1, 9).getValue() !== 'code') cfg.getRange(1, 9).setValue('code');
  return touched + add.length;
}

function seedChecklist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config') || ensureSheets_().getSheetByName('Config');
  var last = cfg.getLastRow();
  var have = {};
  if (last > 1) {
    cfg.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      have[String(r[0]).trim().toLowerCase() + '|' + String(r[1]).trim()] = true;
    });
  }
  if (cfg.getRange(1, 8).getValue() !== 'roles') cfg.getRange(1, 8).setValue('roles');
  var add = [];
  CHECK_ROLES.forEach(function (r) {
    if (!have['role|' + r[0]]) add.push(['role', r[0], r[1], r[2], '', '', false, '']);
  });
  CHECK_SECTIONS.forEach(function (x) {
    if (!have['section|' + x[0]]) add.push(['section', x[0], x[1], x[2], '', '', false, '']);
  });
  CHECK_CATALOGUE.forEach(function (c) {
    if (!have['check|' + c[1]]) add.push(c);
  });
  if (add.length) cfg.getRange(cfg.getLastRow() + 1, 1, add.length, 8).setValues(add);
  var people = seedPeople_(cfg);
  return 'added ' + add.length + ' row(s); people touched ' + people + '; skipped ' + (CHECK_CATALOGUE.length + CHECK_ROLES.length + CHECK_SECTIONS.length - add.length) + ' already there';
}
