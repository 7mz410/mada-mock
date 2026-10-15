// Setup: create a Google Sheet > Extensions > Apps Script > paste this > Deploy > New deployment
// > Web app, Execute as: Me, Who has access: Anyone. Copy the URL into ENDPOINT in index.html.
const NOTIFY = "info@madapcc.org";
const FIELDS = ["name", "email", "country", "city", "type", "message", "lang"];

function doPost(e) {
  const p = e.parameter;
  if (p.website) return ContentService.createTextOutput("ok");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(["date", ...FIELDS]);
  // Prefix "'" so values starting with = + - @ are not run as formulas
  sheet.appendRow([new Date(), ...FIELDS.map(f => "'" + String(p[f] || "").slice(0, 2000))]);
  MailApp.sendEmail(NOTIFY, "New Mada sign-up: " + (p.type || ""), FIELDS.map(f => f + ": " + (p[f] || "")).join("\n"));
  return ContentService.createTextOutput("ok");
}
