
/* ---------- pulling the overdue work forward (run once, 20 Oct 2026) ----------
 * Ten of the twelve open tasks were already past their date when they came over from
 * Notion - the oldest by five weeks. Left alone they would have gone into the calendar
 * as ten entries in the past: an accurate record that warns nobody and prompts nothing.
 *
 * So each one gets a day in the week ahead, highest priority first, two a day, Friday
 * skipped. The original date is not erased - it is written into the calendar entry's
 * description, because how late a thing already is tells you something about it.
 *
 * The sheet and the calendar are set together here. Setting only one would leave two
 * answers to the same question, which is the fault this whole system exists to remove.
 * Row shape: title, new due date, calendar event id.
 */
var RESCHEDULE = [
  ["Samra testing prep",                                            "2026-10-21", "09:00", "dr65ps6dpdp2bppldn327sblq0"],
  ["Safety certificate",                                            "2026-10-21", "10:00", "g41gqgsdsolmghsg91np5mhulk"],
  ["Camera & water filter distribution, and the electricity ampere","2026-10-22", "09:00", "nddu12658ousjc6gv9msks8lg4"],
  ["Ermine new testing (details in Notion)",                        "2026-10-22", "10:00", "2iadvq5ve2adtotgmabm0mm09c"],
  ["Summary of production sheets by Ahmed",                         "2026-10-24", "09:00", "66iu81orr7qh1on1fq6qjr8kls"],
  ["Electricity share",                                             "2026-10-24", "10:00", "cun5jul0jltaaajh7135ieepks"],
  ["Fingers nutritional facts & Samra’s",                      "2026-10-25", "09:00", "jq4fh969g9dttj27vlq7qfktt4"],
  ["Think about the staff’s tasks next step with Adnan",        "2026-10-25", "10:00", "gv212nrihpsb3mi0n10f8ie6hs"],
  ["Chocolate logo, lighters, marshmallow fingers operations",      "2026-10-26", "09:00", "3q49bg83fahec6csckv9dpdrb8"],
  ["Organize WhatsApp groups",                                      "2026-10-27", "09:00", "fa8d13lbl7469pap87r76l53qk"]
];

function applyReschedule() {
  var sh = tasksSheet_();
  var last = sh.getLastRow();
  if (last < 2) return 'no tasks';
  var titles = sh.getRange(2, 3, last - 1, 1).getValues();
  var index = {};
  titles.forEach(function (r, i) {
    var t = String(r[0] || '').trim().toLowerCase();
    if (t) index[t] = i + 2;
  });
  var done = 0, missing = [];
  RESCHEDULE.forEach(function (r) {
    var row = index[String(r[0]).trim().toLowerCase()];
    if (!row) { missing.push(r[0]); return; }
    sh.getRange(row, 6).setValue(r[1]);   // due_date
    sh.getRange(row, 7).setValue(r[2]);   // due_time
    sh.getRange(row, 13).setValue(r[3]);  // calendar_event
    done++;
  });
  var msg = 'rescheduled ' + done + (missing.length ? ', not found: ' + missing.join(' | ') : '');
  Logger.log(msg);
  return msg;
}
