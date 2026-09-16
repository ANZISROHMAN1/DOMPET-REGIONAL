// Mock calculateNeraca to debug
function parseSafeFloat(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  var str = val.toString().replace(/rp/ig, '').trim();
  if (str.indexOf('.') !== -1 && str.indexOf(',') !== -1) {
    str = str.replace(/\./g, '').replace(/,/g, '.');
  } else if (str.indexOf('.') !== -1) {
    var parts = str.split('.');
    if (parts[parts.length - 1].length === 3) str = str.replace(/\./g, '');
  } else if (str.indexOf(',') !== -1) {
    var parts = str.split(',');
    if (parts[parts.length - 1].length === 3) str = str.replace(/,/g, '');
    else str = str.replace(/,/g, '.');
  }
  str = str.replace(/[^0-9.-]/g, '');
  return parseFloat(str) || 0;
}

// Simulate jagoData where TOP is OLDEST
const jagoData = [
  ['Date & Time', 'Source/Destination', 'Transaction Details', 'Notes', 'Kas Masuk', 'Kas Keluar', 'Saldo', 'Unit', 'Kategori'],
  ['26 Sep 2026 16:20', 'A', 'B', '', '', '202000', '453297450', 'EASTERN', 'BODP'],
  ['26 Sep 2026 18:20', 'C', 'D', '', '', '1155000', '452142450', 'EASTERN', 'BODP']
];

const headers = jagoData[0];
const hMap = {};
for (var c = 0; c < headers.length; c++) hMap[String(headers[c]).toLowerCase().trim()] = c;
var isTwoColumnMode = (hMap['kas masuk'] !== undefined && hMap['kas keluar'] !== undefined);
var balColIdx = 6;
var masukColIdx = 4;
var keluarColIdx = 5;

function getRawAmount(row) {
  if (isTwoColumnMode) {
    var m = parseSafeFloat(row[masukColIdx]);
    var k = parseSafeFloat(row[keluarColIdx]);
    return (m > 0) ? m : ((k > 0) ? -k : 0);
  }
  return parseSafeFloat(row[4]);
}

const filteredData = [jagoData[1], jagoData[2]];

var isNewestToOldest = false;
for (var i = 0; i < filteredData.length - 1; i++) {
  var b0 = parseSafeFloat(filteredData[i][balColIdx]);
  var b1 = parseSafeFloat(filteredData[i+1][balColIdx]);
  var a0 = Math.abs(getRawAmount(filteredData[i]));
  var a1 = Math.abs(getRawAmount(filteredData[i+1]));
  
  if (Math.abs(a0 - a1) > 1) {
      if (a0 > 0 && Math.abs(Math.abs(b0 - b1) - a0) < 1) {
          isNewestToOldest = true; break;
      }
      if (a1 > 0 && Math.abs(Math.abs(b1 - b0) - a1) < 1) {
          isNewestToOldest = false; break;
      }
  }
}

var trueAmounts = new Array(filteredData.length).fill(0);
for (var i = 0; i < filteredData.length; i++) {
    var rawAmount = getRawAmount(filteredData[i]);
    var currentBalance = parseSafeFloat(filteredData[i][balColIdx]);
    var trueAmt = rawAmount; 
    var amountMag = Math.abs(rawAmount);
    
    if (isNewestToOldest && i + 1 < filteredData.length) {
        var olderBalance = parseSafeFloat(filteredData[i+1][balColIdx]);
        if (currentBalance > 0 && olderBalance > 0) {
            var diff = currentBalance - olderBalance;
            if (Math.abs(Math.abs(diff) - amountMag) < 1) trueAmt = diff;
            else if (diff < 0 && rawAmount > 0) trueAmt = -amountMag; 
            else if (diff > 0 && rawAmount < 0) trueAmt = amountMag;
        }
    } else if (!isNewestToOldest && i > 0) {
        var olderBalance = parseSafeFloat(filteredData[i-1][balColIdx]);
        if (currentBalance > 0 && olderBalance > 0) {
            var diff = currentBalance - olderBalance;
            if (Math.abs(Math.abs(diff) - amountMag) < 1) trueAmt = diff;
            else if (diff < 0 && rawAmount > 0) trueAmt = -amountMag;
            else if (diff > 0 && rawAmount < 0) trueAmt = amountMag;
        }
    }
    trueAmounts[i] = trueAmt;
}

var saldoAwal = 0;
for (var i = 0; i < filteredData.length; i++) {
    saldoAwal = parseSafeFloat(filteredData[i][balColIdx]) - trueAmounts[i];
    break;
}

console.log("isNewestToOldest:", isNewestToOldest);
console.log("trueAmounts:", trueAmounts);
console.log("saldoAwal:", saldoAwal);
