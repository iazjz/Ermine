
/* ================= Notion -> Tasks, one way =================================
 * Aziz works in Notion. Keeping a second list here that drifts from it is worse than
 * having no list, so Notion is the source and this sheet is a mirror. Nothing written
 * here ever travels back: if a task is wrong, it is wrong in Notion, and that is the
 * only place to fix it. That is the whole point of choosing one direction.
 *
 * What a sync does, every hour:
 *   - reads the Tasks database from Notion
 *   - removes the rows this mirror put here last time
 *   - writes the current ones
 *   - LEAVES ALONE any row that did not come from Notion, so nothing you typed into
 *     the dashboard before the switch is thrown away by a sync you did not watch
 *
 * Setup, once:
 *   1. notion.so/my-integrations -> New integration -> internal -> copy the secret
 *   2. open the Tasks database in Notion -> ... -> Connections -> add that integration
 *   3. paste the secret into NOTION_TOKEN below and save
 *   4. run installNotionTrigger() once
 */

var NOTION_TOKEN = 'PASTE-YOUR-NOTION-SECRET-HERE';
var NOTION_DB    = '3d3dbbef-5168-8091-9469-cc5d8e30760e';   // Workspace / 2026 / Tasks
var NOTION_VER   = '2022-06-28';
var MIRROR_TAG   = 'notion.so/';     // how a mirrored row is recognised, in `notes`

function notionReady_() {
  return NOTION_TOKEN && NOTION_TOKEN.indexOf('PASTE-') !== 0 && NOTION_TOKEN.length > 20;
}

/* ---- reading ---- */

function notionQuery_() {
  var out = [], cursor = null, guard = 0;
  do {
    var body = { page_size: 100 };
    if (cursor) body.start_cursor = cursor;
    var res = UrlFetchApp.fetch(
      'https://api.notion.com/v1/databases/' + NOTION_DB + '/query',
      { method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + NOTION_TOKEN, 'Notion-Version': NOTION_VER },
        payload: JSON.stringify(body), muteHttpExceptions: true });
    var code = res.getResponseCode();
    if (code !== 200) {
      throw new Error('Notion said ' + code + ': ' + res.getContentText().slice(0, 300));
    }
    var j = JSON.parse(res.getContentText());
    out = out.concat(j.results || []);
    cursor = j.has_more ? j.next_cursor : null;
  } while (cursor && ++guard < 20);
  return out;
}

function txt_(prop) {
  if (!prop) return '';
  var a = prop.title || prop.rich_text;
  if (!a) return '';
  return a.map(function (t) { return t.plain_text || ''; }).join('').trim();
}

function notionRow_(page) {
  var p = page.properties || {};
  var date = (p.Date && p.Date.date) || null;
  var start = date ? String(date.start) : '';
  var day = start ? start.slice(0, 10) : '';
  var time = (start.length > 10) ? start.slice(11, 16) : '';
  var people = (p.Person && p.Person.multi_select) || [];
  // The checkbox property is literally named " " in that database - a single space.
  var doneProp = p[' '] || p['Done'] || null;
  var done = !!(doneProp && doneProp.checkbox);
  return {
    id: 'NOTION-' + String(page.id).replace(/-/g, '').slice(-10).toUpperCase(),
    title: txt_(p.Name),
    owner: people.map(function (x) { return x.name; }).filter(function (x) { return x !== '*'; }).join(', '),
    section: (p.Area && p.Area.select && p.Area.select.name) || '',
    due: day,
    time: time,
    priority: (p.Priority && p.Priority.select && p.Priority.select.name) || 'Medium',
    status: done ? 'Done' : 'To do',
    notes: MIRROR_TAG + String(page.id).replace(/-/g, ''),
    doneAt: done ? (String(page.last_edited_time || '').slice(0, 10)) : ''
  };
}

/* ---- writing ---- */

function syncNotionTasks() {
  if (!notionReady_()) {
    Logger.log('no Notion secret set yet - nothing done');
    return 'no token';
  }
  var pages = notionQuery_();
  var rows = pages.map(notionRow_).filter(function (r) { return r.title; });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = tasksSheet_();
    var last = sh.getLastRow();

    // Keep anything that is not a mirror row. A row typed into the dashboard has no
    // Notion link in its notes, and a sync must never be the reason it disappears.
    var keep = [];
    if (last > 1) {
      sh.getRange(2, 1, last - 1, TASK_HEADERS.length).getValues().forEach(function (r) {
        var id = String(r[0] || ''), notes = String(r[10] || '');
        if (!id) return;
        if (id.indexOf('NOTION-') === 0) return;             // last sync's work
        if (notes.indexOf(MIRROR_TAG) !== -1) return;        // the one-off import
        keep.push(r);
      });
      sh.getRange(2, 1, last - 1, TASK_HEADERS.length).clearContent();
    }

    var now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
    var out = keep.slice();
    rows.forEach(function (r) {
      out.push([r.id, now, r.title, r.owner, r.section, r.due, r.time,
                r.priority, r.status, '', r.notes, r.doneAt, '']);
    });
    if (out.length) {
      sh.getRange(2, 1, out.length, TASK_HEADERS.length).setValues(out);
    }
    var msg = 'from Notion: ' + rows.length + ', kept from the app: ' + keep.length;
    Logger.log(msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

/* ---- the hourly trigger ---- */

function installNotionTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncNotionTasks') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncNotionTasks').timeBased().everyHours(1).create();
  var msg = 'hourly sync installed';
  Logger.log(msg);
  return msg;
}
