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
var ACTION_HEADERS = ['id', 'created_at', 'updated_at', 'report_id', 'date', 'section', 'item_id', 'item_ar', 'item_en', 'employee', 'problem', 'owner', 'action', 'due', 'status', 'closed_at'];

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
    case 'config':
      requireCode_(req.code, TEAM_CODE);
      return { ok: true, config: readConfig_() };
    case 'submit':
      requireCode_(req.code, TEAM_CODE);
      return submit_(req.report);
    case 'list':
      requireCode_(req.managerCode, MANAGER_CODE);
      return list_(req);
    case 'actions':
      requireCode_(req.managerCode, MANAGER_CODE);
      return actions_();
    case 'action_save':
      requireCode_(req.managerCode, MANAGER_CODE);
      return actionSave_(req.entry);
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
    rep.getRange('C:D').setNumberFormat('@');
  }
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
    cfg.getRange('H1').setValue('kind: employee | unit | zone | hygiene.  type: fridge | freezer.  limit: max °C.  hide: tick to hide a row.');
  }
  var acts = ss.getSheetByName('Actions');
  if (!acts) {
    acts = ss.insertSheet('Actions');
    acts.appendRow(ACTION_HEADERS);
    acts.setFrozenRows(1);
    acts.getRange(1, 1, 1, ACTION_HEADERS.length).setFontWeight('bold');
    acts.getRange('B:C').setNumberFormat('@');
    acts.getRange('E:E').setNumberFormat('@');
    acts.getRange('N:N').setNumberFormat('@');
    acts.getRange('P:P').setNumberFormat('@');
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
  var cfg = { employees: [], units: [], zones: [], hyg: [] };
  var last = sh.getLastRow();
  if (last < 2) return cfg;
  var rows = sh.getRange(2, 1, last - 1, 7).getValues();
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
    if (kind === 'employee') {
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
    ['hygiene', 'hyg', rep.items.hyg]
  ];
  var clean = { units: [], zones: [], hyg: [] };
  var itemRows = [];
  var done = 0, total = 0, issues = 0;

  sections.forEach(function (s) {
    var list = Array.isArray(s[2]) ? s[2].slice(0, 60) : [];
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
      var d = cellStr_(r[2], 'yyyy-MM-dd');
      if (d < from || d > to) return;
      var items = {};
      try { items = JSON.parse(r[11]); } catch (e) { items = {}; }
      out.push({
        id: String(r[0]), submittedAt: cellStr_(r[1], "yyyy-MM-dd'T'HH:mm:ssXXX"), date: d, shift: String(r[3]), employee: String(r[4]),
        workers: r[5], done: r[6], total: r[7], issues: r[8], notes: String(r[9]), lang: String(r[10]), items: items
      });
    });
  }
  out.sort(function (a, b) { return a.submittedAt < b.submittedAt ? 1 : a.submittedAt > b.submittedAt ? -1 : 0; });
  return { ok: true, reports: out.slice(0, 500), sheetUrl: ss.getUrl() };
}

// Sheets sometimes turns a typed date into a real Date; read both forms as the same text.
function cellStr_(v, fmt) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, fmt);
  return String(v == null ? '' : v);
}

/* ---------- corrective actions ---------- */

function actionRow_(r) {
  return {
    id: String(r[0]), createdAt: cellStr_(r[1], "yyyy-MM-dd'T'HH:mm:ssXXX"), updatedAt: cellStr_(r[2], "yyyy-MM-dd'T'HH:mm:ssXXX"),
    reportId: String(r[3]), date: cellStr_(r[4], 'yyyy-MM-dd'), section: String(r[5]), itemId: String(r[6]),
    itemAr: String(r[7]), itemEn: String(r[8]), employee: String(r[9]), problem: String(r[10]),
    owner: String(r[11]), action: String(r[12]), due: cellStr_(r[13], 'yyyy-MM-dd'),
    status: String(r[14]) === 'done' ? 'done' : 'open', closedAt: cellStr_(r[15], "yyyy-MM-dd'T'HH:mm:ssXXX")
  };
}

function actions_() {
  var sh = ensureSheets_().getSheetByName('Actions');
  var last = sh.getLastRow(), out = [];
  if (last > 1) {
    sh.getRange(2, 1, last - 1, ACTION_HEADERS.length).getValues().forEach(function (r) { if (r[0]) out.push(actionRow_(r)); });
  }
  out.sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : a.createdAt > b.createdAt ? -1 : 0; });
  return { ok: true, actions: out.slice(0, 1000) };
}

function actionSave_(e) {
  if (!e || typeof e !== 'object') throw new Error('invalid_action');
  var section = e.section === 'units' || e.section === 'zones' || e.section === 'hyg' ? e.section : '';
  var reportId = String(e.reportId || '');
  var itemId = clip_(e.itemId, 60);
  var due = /^\d{4}-\d{2}-\d{2}$/.test(e.due) ? e.due : '';
  var status = e.status === 'done' ? 'done' : 'open';
  var owner = clip_(e.owner, 80), action = clip_(e.action, 500);
  if (!section || !itemId || !/^[A-Za-z0-9_-]{8,64}$/.test(reportId)) throw new Error('invalid_action');
  if (!owner || !action) throw new Error('invalid_action');
  var date = /^\d{4}-\d{2}-\d{2}$/.test(e.date) ? e.date : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var stamp = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ensureSheets_().getSheetByName('Actions');
    var last = sh.getLastRow(), rowNo = 0, old = null;
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, ACTION_HEADERS.length).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var same = e.id ? String(r[0]) === String(e.id) : (String(r[3]) === reportId && String(r[5]) === section && String(r[6]) === itemId);
        if (same) { rowNo = i + 2; old = r; break; }
      }
    }
    if (e.id && !rowNo) throw new Error('not_found');
    var closedAt = status === 'done' ? (old && String(old[14]) === 'done' && old[15] ? cellStr_(old[15], "yyyy-MM-dd'T'HH:mm:ssXXX") : stamp) : '';
    var row;
    if (rowNo) {
      row = [old[0], old[1], stamp, old[3], old[4], old[5], old[6], old[7], old[8], old[9], old[10], safe_(owner), safe_(action), due, status, closedAt];
      sh.getRange(rowNo, 1, 1, ACTION_HEADERS.length).setValues([row]);
    } else {
      var id = 'a' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.random().toString(36).slice(2, 7);
      row = [id, stamp, stamp, reportId, date, section, itemId, safe_(clip_(e.itemAr, 120)), safe_(clip_(e.itemEn, 120)),
        safe_(clip_(e.employee, 80)), safe_(clip_(e.problem, 300)), safe_(owner), safe_(action), due, status, closedAt];
      sh.appendRow(row);
    }
    return { ok: true, action: actionRow_(row) };
  } finally {
    lock.releaseLock();
  }
}
