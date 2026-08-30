function getNeracaDebug() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REKAPAN JAGO');
  var data = sheet.getDataRange().getValues();
  var log = [];
  var totalPengeluaran = 0;
  for (var i = 1; i < data.length; i++) {
    var amt = parseFloat(data[i][4]) || 0;
    var unit = data[i][6];
    if (amt < -10000000 || amt > 10000000) {
       log.push("Row " + (i+1) + ": " + amt + " | Unit: " + unit + " | Detail: " + data[i][2] + " | Notes: " + data[i][3]);
    }
    if (amt < 0) totalPengeluaran += Math.abs(amt);
  }
  return {log: log, totalPengeluaran: totalPengeluaran};
}
