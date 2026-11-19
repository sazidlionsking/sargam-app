// Sargam reviews backend. Paste this whole file into Apps Script.
var SHEET = "Reviews";
function sh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) s.appendRow(["Time", "Name", "Stars", "Comment", "Show (yes/no)"]);
  return s;
}
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var stars = Math.max(1, Math.min(5, parseInt(d.stars, 10) || 0));
    if (!stars) return ContentService.createTextOutput("bad");
    var clean = function (v, n) { v = String(v || "").slice(0, n); return /^[=+\-@]/.test(v) ? "'" + v : v; };
    sh().appendRow([new Date(), clean(d.name, 30), stars, clean(d.text, 400), "yes"]);
    return ContentService.createTextOutput("ok");
  } catch (err) { return ContentService.createTextOutput("error"); }
}
function doGet(e) {
  var rows = sh().getDataRange().getValues().slice(1), out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[4]).toLowerCase() === "no" || !r[2]) continue;
    var dt = r[0] instanceof Date ? Utilities.formatDate(r[0], "Asia/Kolkata", "d MMM yyyy") : "";
    out.push({ name: r[1], stars: r[2], text: r[3], date: dt });
  }
  return ContentService.createTextOutput(JSON.stringify({ reviews: out.slice(-200) })).setMimeType(ContentService.MimeType.JSON);
}
