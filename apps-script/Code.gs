/**
 * Inclusive · Actividades: saves each traveler selection to the "Pedidos" tab.
 * Paste into Extensions > Apps Script of the ACTIVIDADES sheet, then
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Send the /exec URL so it can be set as ORDER_URL in index.html.
 */
function doPost(e) {
  var data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) {}
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  sheet.appendRow([
    new Date(),
    clean_(data.viajero),
    clean_(data.actividades),
    clean_(data.totales),
    ''
  ]);
  return ContentService.createTextOutput('ok');
}

// Keeps text short and stops a value from being read as a formula.
function clean_(v) {
  var s = String(v || '').slice(0, 5000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
