/* ---------- one-time import of the Notion board (Workspace / 2026 / Tasks) ----------
 * Aziz kept his tasks in Notion. They move here so that one list is the list, rather
 * than two that disagree. Everything comes across, finished work included: the finished
 * rows land straight in the archive, where they cost nothing to keep and answer "when
 * did we do that" later. Each row keeps a link back to its Notion page, so nothing is
 * orphaned by the move.
 *
 * Run importNotionTasks() ONCE. It is safe to run twice - a title that is already in the
 * sheet is skipped rather than duplicated - but there is no reason to.
 * Row shape: title, owner, section, due, time, priority, status, repeat, notes
 */
var IMPORT_ROWS = [
  ["Muath decision whether chef can go to Samra", "", "Operations", "2026-09-06", "21:00", "Medium", "Done", "", "notion.so/3d3dbbef51688051b3ebeec298162a1f"],
  ["Send Ali to reach Ismael", "", "Stores", "2026-09-07", "", "Medium", "Done", "", "notion.so/3d3dbbef516880919ab2ddcbaea177fa"],
  ["Praline cake testing with adjustments (deeper hole for more filling)", "", "Quality", "2026-09-07", "", "High", "Done", "", "notion.so/3d3dbbef516880f99a42fd289c55d9fb"],
  ["Ermine’s products testing", "Diloshan", "Quality", "2026-09-07", "", "High", "Done", "", "notion.so/3d4dbbef5168803f95b1c199191cc085"],
  ["Test half macarons in pizza boxes (quantity & quality?)", "", "Quality", "2026-09-07", "", "Medium", "Done", "", "notion.so/3d4dbbef516880b88fa0cbc99051b09d"],
  ["How to cut the brownie", "", "Operations", "2026-09-07", "", "Medium", "Done", "", "notion.so/3d4dbbef516880f39b1de5280c87b17e"],
  ["Code half macarons and fingers in pizza boxes", "", "Finance", "2026-09-07", "", "High", "Done", "", "notion.so/3d4dbbef516880fd9aaff120c12d466d"],
  ["Photo shoot at the factory", "", "Operations", "2026-09-07", "13:00", "Medium", "Done", "", "notion.so/3d4dbbef516880b2ab89f83299e451e4"],
  ["Production stages files by Ahmed + sheets summary", "", "Operations", "2026-09-07", "14:00", "Medium", "Done", "", "notion.so/3d3dbbef5168806db8f2d096724d143a"],
  ["Identify day and date of new Abdurrahman order", "", "Sales", "2026-09-07", "16:00", "Medium", "Done", "", "notion.so/3d4dbbef516880809e78f38e6098eda3"],
  ["Test old bites boxes", "", "Quality", "2026-09-08", "", "Medium", "Done", "", "notion.so/3d4dbbef5168808099d8fb80c3615a3f"],
  ["Ice cream pickup", "", "Operations", "2026-09-08", "", "Medium", "Done", "", "notion.so/3d4dbbef516880c0ba8be33879e52384"],
  ["Praline filling", "", "Purchasing", "2026-09-08", "", "High", "Done", "", "notion.so/3d4dbbef516880cfba85d2000b143c7e"],
  ["Decision on new cafe pcs", "", "Sales", "2026-09-08", "16:00", "High", "Done", "", "notion.so/3d5dbbef51688093a94bf40c609a0f23"],
  ["Brownie ring", "", "Purchasing", "2026-09-09", "", "Medium", "Done", "", "notion.so/3d4dbbef5168804a888efa4e1e85f50b"],
  ["اجتماع مع الشيف لمناقشة فروقات سمراء", "", "Operations", "2026-09-09", "10:00", "High", "Done", "", "notion.so/3d4dbbef5168800ab021ddcbdc241f1b"],
  ["Chef to compare the recipes between what we approved and what is actually produced", "", "Quality", "2026-09-10", "", "High", "Done", "", "notion.so/3d6dbbef5168802c87a3c8c464fa3c54"],
  ["Decision on new accommodation", "", "HR", "2026-09-14", "", "High", "Done", "", "notion.so/3dadbbef5168808d8de3cc4b4b6f5101"],
  ["Meet chef and Raqeeb and compare recipes — finalize the master file", "", "Operations", "2026-09-14", "", "High", "Done", "", "notion.so/3dadbbef5168809fbf25e048364f1be2"],
  ["Meeting w/ Adnan to discuss the final roles", "", "HR", "2026-09-14", "", "High", "Done", "", "notion.so/3dadbbef516880c6a65ddf51b3bab80c"],
  ["Updates to Muath on the production dept. development process", "", "Strategy", "2026-09-14", "", "High", "Done", "", "notion.so/3dadbbef516880cb9df8fa89108ba1aa"],
  ["Get Raqeeb’s production reports to see what he does", "", "Operations", "2026-09-14", "", "Medium", "Done", "", "notion.so/3dadbbef516880f5842dce5d7224bd29"],
  ["Dough sheeter maintenance", "", "Maintenance", "2026-09-14", "", "High", "Done", "", "notion.so/3dbdbbef516880b88f9bc7abc1e0407c"],
  ["Summary of production sheets by Ahmed", "Ahmed", "Operations", "2026-09-15", "", "Medium", "To do", "", "notion.so/3d5dbbef51688017914ee25d95d11cc6"],
  ["Raw materials for new finger flavors and new praline cake", "", "Purchasing", "2026-09-15", "", "High", "Done", "", "notion.so/3d5dbbef516880f7abc4e05ae6a3e44e"],
  ["Get the work stage tracker from Nitish for the packing stages", "", "Operations", "2026-09-15", "", "High", "Done", "", "notion.so/3dbdbbef516880bcb3fcc37a3bb1e005"],
  ["Meeting Shehab to add a door", "", "Operations", "2026-09-15", "", "High", "Done", "", "notion.so/3dcdbbef516880299b8dc28f833df2e0"],
  ["Who can go to Samra for testing instead of the chef", "", "Operations", "2026-09-15", "", "Medium", "Done", "", "notion.so/3dcdbbef516880699007e976598baa65"],
  ["Final check of the dept. yearly goals before the meeting", "", "Strategy", "2026-09-15", "", "High", "Done", "", "notion.so/3dcdbbef5168806b8793f92cd9a6dd65"],
  ["Revise sales pcs order (Dahoum’s message) — provide dates", "", "Sales", "2026-09-15", "", "High", "Done", "", "notion.so/3dcdbbef5168807d8020e4f7f7183eb1"],
  ["What percentage of the month’s produced pcs is each product?", "", "Operations", "2026-09-16", "", "Medium", "Done", "", "notion.so/3dbdbbef5168807f90a2fbc8808f59b6"],
  ["Shaikha insurance registration", "", "HR", "2026-09-16", "", "Medium", "Done", "", "notion.so/3dcdbbef516880e0a181c1a5a90602e7"],
  ["Organize WhatsApp groups", "", "Operations", "2026-09-16", "17:00", "Low", "To do", "", "notion.so/3d4dbbef516880a99c0ce82a9e71c058"],
  ["Enter the production stages on the major plan and test if it works", "", "Operations", "2026-09-17", "", "Medium", "Done", "", "notion.so/3dddbbef516880d3b528f152b1f4049d"],
  ["Mahesh iqama renewal", "", "HR", "2026-09-20", "", "Medium", "Done", "", "notion.so/3dbdbbef51688027af03c17633ca49b7"],
  ["Transfer Gihan to Samra for testing", "", "Operations", "2026-09-20", "", "Medium", "Done", "", "notion.so/3e1dbbef5168807da9f5f173e9a3a66c"],
  ["خذ المواد الناقصة من أسامة في سمراء للشهر الماضي بالأوزان", "", "Operations", "2026-09-21", "", "Medium", "Done", "", "notion.so/3e1dbbef516880059f2df446fa402d76"],
  ["اعتمد الرسبي وأرسله في إيميل", "", "Operations", "2026-09-21", "", "Medium", "Done", "", "notion.so/3e1dbbef51688038a336c031504bcb93"],
  ["عدد المنتجات في المصنع في سنة ٢٠٢٥", "", "Strategy", "2026-09-21", "", "Medium", "Done", "", "notion.so/3e1dbbef516880a281c4d671307f63a9"],
  ["من مسؤول رسبي سمراء، أحمد أم محمد — مع أسامة", "", "Quality", "2026-09-21", "", "Medium", "Done", "", "notion.so/3e1dbbef516880cfa2d8d5344892e895"],
  ["سداد قطع غيار الفرادة", "", "Maintenance", "2026-09-21", "", "Medium", "Done", "", "notion.so/3e2dbbef51688020ad9ac00874004578"],
  ["اطلب من عدنان البحث عن موظفين خبراء لسمراء", "Adnan", "HR", "2026-09-21", "", "Medium", "Done", "", "notion.so/3e2dbbef51688071a9e1e7a4e9f9827c"],
  ["تخطيط المكاتب وغرفة الموظفين — حل دخول المناديب", "", "Operations", "2026-09-22", "", "Medium", "Done", "", "notion.so/3e2dbbef5168806fb830f3f5b714919d"],
  ["Trolley QTY and stock ready to transfer to Logexa", "", "Stores", "2026-09-22", "", "Medium", "Done", "", "notion.so/3e3dbbef516880f3bb7be99b9f7d2da4"],
  ["Chef to test Yan’s", "", "Quality", "2026-09-26", "", "High", "Done", "", "notion.so/3d3dbbef516880fb9407e1f0eff63e77"],
  ["Samra testing prep", "Diloshan", "Quality", "2026-09-26", "", "High", "To do", "", "notion.so/3d4dbbef51688070bf40c8cdac971af7"],
  ["Deliver samples from Samra to the office with Ermine driver", "Hany", "Operations", "2026-10-01", "", "High", "Done", "", "notion.so/3eadbbef51688017972ded6f02685f33"],
  ["Charms barcodes", "Diloshan", "Operations", "2026-10-01", "", "High", "Done", "", "notion.so/3ebdbbef5168802b87b4e457cf25db76"],
  ["Electricity share", "Osama", "Finance", "2026-10-01", "", "Medium", "To do", "", "notion.so/3ebdbbef516880efa6bbd573d87442d4"],
  ["Master planning file meeting", "Diloshan", "Operations", "2026-10-01", "10:00", "High", "Done", "", "notion.so/3ebdbbef5168807092cae67c1bf62208"],
  ["Don’t wrap the master cartons — salesmen check then tape them off", "Ali", "Stores", "2026-10-04", "", "Low", "Done", "", "notion.so/3d4dbbef5168808f9ac1d2b9b136e569"],
  ["Makeel invoice", "", "Purchasing", "2026-10-06", "", "Low", "Done", "", "notion.so/3f0dbbef5168803b8da6d0c07c603aa0"],
  ["Payment to landlord whose apartment’s electricity is off", "Osama", "Finance", "2026-10-06", "", "High", "Done", "", "notion.so/3f0dbbef5168806096fec04b363e6281"],
  ["Praline cake nutritional facts", "", "Operations", "2026-10-06", "", "High", "Done", "", "notion.so/3f0dbbef51688075b76af3e34b1c3cd1"],
  ["Staff vacation requests", "Adnan", "HR", "2026-10-06", "", "Medium", "Done", "", "notion.so/3f0dbbef516880c88a38c504ed96ef3f"],
  ["Adding pcs stock and plans", "Diloshan", "Operations", "2026-10-06", "", "High", "Done", "", "notion.so/3f0dbbef516880cc9cb4e36faffbd696"],
  ["Fingers nutritional facts & Samra’s", "", "Operations", "2026-10-07", "", "Medium", "To do", "", "notion.so/3f0dbbef516880ed92bde078994a4645"],
  ["Camera transfer decision", "", "Operations", "2026-10-08", "", "High", "Done", "", "notion.so/3f3dbbef516880209679ef054b786cbe"],
  ["Safety certificate", "", "Operations", "2026-10-08", "", "High", "To do", "", "notion.so/3f3dbbef51688056bacbce1dcc25a70c"],
  ["Follow up increasing cakes and coins in the production plan", "", "Operations", "2026-10-13", "", "High", "Done", "", "notion.so/3f7dbbef51688017b411cf9b3751c331"],
  ["Camera & water filter distribution, and the electricity ampere", "", "Operations", "2026-10-13", "", "High", "To do", "", "notion.so/3f7dbbef5168801b9d28fed95daa5c09"],
  ["Think about the staff’s tasks next step with Adnan", "Adnan", "HR", "2026-10-13", "", "Medium", "To do", "", "notion.so/3f7dbbef51688059a694de10c127a842"],
  ["Ermine new testing (details in Notion)", "Diloshan", "R&D", "2026-10-18", "", "High", "To do", "", "notion.so/3f7dbbef516880de958bdcd571c36cf7"],
  ["Chocolate logo, lighters, marshmallow fingers operations", "Diloshan, Saif", "R&D", "2026-10-18", "", "Medium", "To do", "", "notion.so/3f9dbbef5168800081adcba91a607141"],
  ["Coins molds QTY", "", "R&D", "2026-10-19", "", "High", "Done", "", "notion.so/3f7dbbef5168801d8651e005bf77df73"],
  ["Nitish salary", "", "HR", "", "", "Medium", "To do", "", "notion.so/3e3dbbef5168800ba8ffef50752481a3"],
  ["Insurance issue with Adnan", "", "HR", "", "", "Medium", "Done", "", "notion.so/3e3dbbef5168800d8358f8fda256ea56"],
  ["Consumption for Veliche", "", "Purchasing", "", "", "Medium", "Done", "", "notion.so/3e3dbbef516880e9b8e9dc269297d3da"],
  ["Inform landlords of leaving rooms soon", "", "", "", "", "Medium", "To do", "", "notion.so/3fedbbef5168805eab20fa59a037ae09"]
];

function importNotionTasks() {
  var sh = tasksSheet_();
  var have = {};
  var last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 3, last - 1, 1).getValues().forEach(function (r) {
      var t = String(r[0] || '').trim();
      if (t) have[t.toLowerCase()] = true;
    });
  }
  var now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  var add = [], skipped = 0, n = 0;
  IMPORT_ROWS.forEach(function (r) {
    var title = String(r[0] || '').trim();
    if (!title) return;
    if (have[title.toLowerCase()]) { skipped++; return; }
    n++;
    add.push([
      'N' + (new Date().getTime() + n).toString(36).toUpperCase(),
      now, title, r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8],
      r[6] === 'Done' ? now.slice(0, 10) : '', ''
    ]);
  });
  if (add.length) {
    sh.getRange(sh.getLastRow() + 1, 1, add.length, TASK_HEADERS.length).setValues(add);
  }
  Logger.log('imported ' + add.length + ', skipped ' + skipped + ' already there');
  return 'imported ' + add.length + ', skipped ' + skipped;
}
