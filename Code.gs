// FUNGSI INI WAJIB DI-RUN SATU KALI UNTUK MEMAKSA IZIN PENUH
function paksaIzin() {
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getRootFolder().getFiles();
  DriveApp.createFile("test", "test", MimeType.PLAIN_TEXT).setTrashed(true);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('FORM USER') || ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var action = e.parameter.action;
  
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    row['row_index'] = i + 1;
    rows.push(row);
  }
  
  if (action === 'status') {
    var reqId = e.parameter.id;
    var result = rows.find(function(r) { return r['ID'].toString() === reqId.toString(); });
    if (result) return ContentService.createTextOutput(JSON.stringify({success: true, data: result})).setMimeType(ContentService.MimeType.JSON);
    else return ContentService.createTextOutput(JSON.stringify({success: false, message: 'ID tidak ditemukan'})).setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'list') {
    rows.reverse();
    return ContentService.createTextOutput(JSON.stringify({success: true, data: rows})).setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'jago_data') {
    var jagoSheet = ss.getSheetByName('REKAPAN JAGO');
    if (!jagoSheet) return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Sheet REKAPAN JAGO tidak ditemukan'})).setMimeType(ContentService.MimeType.JSON);
    
    var jagoData = jagoSheet.getDataRange().getValues();
    var summaryData = {};
    for (var k = 1; k < jagoData.length; k++) {
      var amount = parseFloat(jagoData[k][4]) || 0;
      var unit = jagoData[k][6] || 'Tanpa Unit';
      
      if (amount < 0) { // Only count expenses
        var absAmount = Math.abs(amount);
        if (!summaryData[unit]) summaryData[unit] = { count: 0, total: 0 };
        summaryData[unit].count += 1;
        summaryData[unit].total += absAmount;
      }
    }
    
    var resultData = [];
    for (var u in summaryData) {
      resultData.push({ unit: u, count: summaryData[u].count, total: summaryData[u].total });
    }
    return ContentService.createTextOutput(JSON.stringify({success: true, data: resultData})).setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'neraca_data') {
    var neracaSheet = ss.getSheetByName('Neraca Keuangan');
    if (!neracaSheet) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Sheet Neraca Keuangan belum ada.'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = neracaSheet.getDataRange().getValues();
    if (data.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Data kosong'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true, data: data})).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Action not found'})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var headers = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, GET, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" };
  if (typeof e !== 'undefined' && e.postData === undefined) return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.JSON);
  
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('FORM USER') || ss.getActiveSheet();
    var sheetJago = ss.getSheetByName('REKAPAN JAGO');
    
    // ID FOLDER SUDAH OTOMATIS SAYA MASUKKAN DI SINI! (TIDAK PERLU DIUBAH LAGI)
    var FOLDER_ID = '1GRHerfG8UMcQol4TY5HBvGS7NPXKYuL_'; 
    var folder = DriveApp.getFolderById(FOLDER_ID);
    
    if (action === 'submit') {
      var fileUrl = '';
      var fileHash = '';
      if (requestData.fileData && requestData.fileName && requestData.mimeType) {
        var base64Data = requestData.fileData.split(',')[1] || requestData.fileData;
        var byteData = Utilities.base64Decode(base64Data);
        
        // Cek duplikasi nota menggunakan MD5 hash
        var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, byteData);
        fileHash = digest.map(function(byte) { return ('0' + (byte & 0xFF).toString(16)).slice(-2); }).join('');
        
        var sheetData = sheet.getDataRange().getValues();
        for (var i = 1; i < sheetData.length; i++) {
          if (sheetData[i][12] === fileHash && fileHash !== '') {
            return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Nota sudah pernah diajukan'})).setMimeType(ContentService.MimeType.JSON);
          }
        }
        
        var blob = Utilities.newBlob(byteData, requestData.mimeType, requestData.fileName);
        var file = folder.createFile(blob);
        fileUrl = file.getUrl();
      }
      
      var newId = new Date().getTime().toString().slice(-6);
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      
      var rowData = [newId, now, requestData.nama, requestData.kegiatan, requestData.nominal, requestData.bank, requestData.rekening, 'Pending', requestData.unit, requestData.sub_unit, fileUrl, '', fileHash];
      sheet.appendRow(rowData);
      
      // Menambahkan data ke sheet REKAPAN JAGO dengan format khusus
      if (sheetJago) {
        var sourceDest = requestData.nama + "\n" + requestData.bank + " " + requestData.rekening;
        var transDetails = "Claim ID# " + newId;
        var nominalStr = requestData.nominal ? requestData.nominal.toString().replace(/[^0-9]/g, '') : "0";
        var amount = -Math.abs(parseFloat(nominalStr)); // Pengeluaran (minus)
        
        var jagoDataToInsert = [now, sourceDest, transDetails, requestData.kegiatan, amount, "", requestData.unit];
        
        // Mencari baris kosong pertama di kolom A (Menghindari bug appendRow jika ada ArrayFormula)
        var jagoColA = sheetJago.getRange("A:A").getValues();
        var jagoTargetRow = jagoColA.length + 1;
        for (var i = 0; i < jagoColA.length; i++) {
          if (jagoColA[i][0] === "" && i > 0) { // i > 0 untuk melewati header
            jagoTargetRow = i + 1;
            break;
          }
        }
        
        sheetJago.getRange(jagoTargetRow, 1, 1, 7).setValues([jagoDataToInsert]);
        formatTransactions(sheetJago);
      }
      
      return ContentService.createTextOutput(JSON.stringify({success: true, id: newId})).setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === 'update_status') {
      var tfUrl = '';
      if (requestData.fileData && requestData.fileName && requestData.mimeType) {
        var base64Data = requestData.fileData.split(',')[1] || requestData.fileData;
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), requestData.mimeType, "TF_" + requestData.fileName);
        var file = folder.createFile(blob);
        tfUrl = file.getUrl();
      }
      
      var rowIndex = parseInt(requestData.row_index);
      sheet.getRange(rowIndex, 8).setValue(requestData.status);
      if (tfUrl) sheet.getRange(rowIndex, 12).setValue(tfUrl);
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function formatTransactions(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 5, lastRow - 1, 1);
  var rules = sheet.getConditionalFormatRules();
  var newRules = rules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(function(r) { return r.getColumn() === 5; });
  });
  var incomeRule = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground("#d9ead3").setFontColor("#137333").setRanges([range]).build();
  var expenseRule = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground("#f4cccc").setFontColor("#990000").setRanges([range]).build();
  newRules.push(incomeRule);
  newRules.push(expenseRule);
  sheet.setConditionalFormatRules(newRules);
  range.setNumberFormat('"Rp" #,##0');
}

// --- TAMBAHAN UNTUK IMPORT PDF JAGO ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ KLAIM.ID')
      .addItem('📄 Import Rekapan Jago (PDF)', 'showImportDialog')
      .addToUi();
}

function showImportDialog() {
  var html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: sans-serif; padding: 20px; }
          .container { display: flex; flex-direction: column; gap: 15px; }
          button { background: #1a73e8; color: white; border: none; padding: 10px; border-radius: 4px; cursor: pointer; }
          button:hover { background: #1557b0; }
          #status { color: #555; font-size: 14px; margin-top: 10px; }
        </style>
        <!-- Load pdf.js -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
      </head>
      <body>
        <div class="container">
          <h3>Upload PDF Rekapan Jago</h3>
          <p>Pilih file PDF mutasi dari Bank Jago untuk di-import otomatis ke sheet REKAPAN JAGO.</p>
          <input type="file" id="fileInput" accept="application/pdf" />
          <button onclick="uploadFile()">Import Data</button>
          <div id="status"></div>
        </div>
        
        <script>
          pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
          
          async function uploadFile() {
            var file = document.getElementById('fileInput').files[0];
            if (!file) {
              alert('Silakan pilih file PDF terlebih dahulu!');
              return;
            }
            
            document.getElementById('status').innerText = "Membaca tabel PDF... (Harap tunggu)";
            
            try {
              var arrayBuffer = await file.arrayBuffer();
              var pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
              var parsedData = [];
              
              for (var i = 1; i <= pdf.numPages; i++) {
                var page = await pdf.getPage(i);
                var textContent = await page.getTextContent();
                var viewport = page.getViewport({scale: 1.0});
                var width = viewport.width; // standard A4 is ~595
                
                // Group by Y coordinate (tolerate 3px difference)
                var rows = {};
                textContent.items.forEach(function(item) {
                  var text = item.str.trim();
                  if (!text) return;
                  var y = Math.round(item.transform[5]);
                  var x = item.transform[4];
                  
                  var foundY = y;
                  for (var key in rows) {
                    if (Math.abs(key - y) <= 4) { foundY = key; break; }
                  }
                  if (!rows[foundY]) rows[foundY] = [];
                  rows[foundY].push({text: text, x: x});
                });
                
                // Sort Y descending (top to bottom of page)
                var sortedY = Object.keys(rows).sort(function(a, b) { return b - a; });
                
                sortedY.forEach(function(y) {
                  var items = rows[y];
                  var col1=[], col2=[], col3=[], col4=[], col5=[], col6=[];
                  
                  // Distribute to columns based on X percentage of width
                  items.forEach(function(it) {
                    var pct = it.x / width;
                    if (pct < 0.20) col1.push(it.text);
                    else if (pct < 0.44) col2.push(it.text);
                    else if (pct < 0.65) col3.push(it.text);
                    else if (pct < 0.82) col4.push(it.text);
                    else if (pct < 0.90) col5.push(it.text);
                    else col6.push(it.text);
                  });
                  
                  var rowData = [
                    col1.join(' '), col2.join(' '), col3.join(' '), 
                    col4.join(' '), col5.join(' '), col6.join(' ')
                  ];
                  
                  // Only push rows that have some text
                  if (rowData.join('').trim() !== '') {
                    parsedData.push(rowData);
                  }
                });
              }
              
              // Process multi-line transactions
              var finalTransactions = [];
              var currentTx = null;
              
              parsedData.forEach(function(row) {
                // If col1 matches a Date format (e.g., 11 Aug 2026 or 11:52)
                var hasDate = row[0].match(/\\d{1,2} [A-Z][a-z]{2} \\d{4}/);
                var hasTime = row[0].match(/\\d{2}:\\d{2}/);
                
                if (hasDate) {
                  if (currentTx) finalTransactions.push(currentTx);
                  currentTx = [...row];
                } else if (currentTx && (hasTime || row[1] || row[2])) {
                  // Merge multi-line into current transaction
                  for (var c = 0; c < 6; c++) {
                    if (row[c]) currentTx[c] += (currentTx[c] ? "\\n" : "") + row[c];
                  }
                }
              });
              if (currentTx) finalTransactions.push(currentTx);
              
              document.getElementById('status').innerText = "Menyimpan ke Spreadsheet...";
              
              google.script.run
                .withSuccessHandler(function(res) {
                  document.getElementById('status').innerHTML = "<span style='color:green'>" + res + "</span>";
                })
                .withFailureHandler(function(err) {
                  document.getElementById('status').innerHTML = "<span style='color:red'>Error: " + err.message + "</span>";
                })
                .processParsedJagoData(finalTransactions);
                
            } catch (e) {
              document.getElementById('status').innerHTML = "<span style='color:red'>Error PDF: " + e.message + "</span>";
            }
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(500)
  .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Rekapan Jago');
}

function processParsedJagoData(transactions) {
  if (!transactions || transactions.length === 0) {
    throw new Error("Tabel kosong atau gagal diproses dari PDF.");
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('REKAPAN JAGO');
  if (!sheet) throw new Error("Sheet bernama 'REKAPAN JAGO' tidak ditemukan!");
  
  var existingData = sheet.getDataRange().getValues();
  var existingIds = {};
  for (var i = 1; i < existingData.length; i++) {
    var details = existingData[i][2] ? existingData[i][2].toString() : "";
    var match = details.match(/ID#\\s*([A-Za-z0-9-]+)/);
    if (match) existingIds[match[1]] = true;
  }
  
  var addedCount = 0;
  
  for (var i = 0; i < transactions.length; i++) {
    var trx = transactions[i];
    
    // Combine Notes, Amount, and Balance columns to safely extract the last two numbers
    // This handles cases where pdf.js column boundaries shift slightly
    var notesAndMoney = (trx[3] + " " + trx[4] + " " + trx[5]).trim();
    var tokens = notesAndMoney.split(/\s+/);
    
    if (tokens.length < 2) continue; // Invalid row
    
    var balanceStr = tokens.pop();
    var amountStr = tokens.pop();
    var notes = tokens.join(" "); // Anything left over belongs to Notes
    
    if (amountStr.toLowerCase().includes('amount')) continue;
    
    var transIdMatch = trx[2].match(/ID#\s*([A-Za-z0-9-]+)/);
    var transId = transIdMatch ? transIdMatch[1] : null;
    
    if (transId && existingIds[transId]) continue; 
    
    // Clean numbers
    var amount = parseFloat(amountStr.replace(/\./g, '').replace(/,/g, '.'));
    var balance = parseFloat(balanceStr.replace(/\./g, '').replace(/,/g, '.'));
    
    sheet.appendRow([
      trx[0], // Date & Time
      trx[1], // Source/Dest
      trx[2], // Details
      notes,  // Corrected Notes
      amount,
      balance, // Actual Balance
      "" // Unit
    ]);
    addedCount++;
  }
  
  formatTransactions(sheet);
  
  try {
    updateNeracaKeuangan();
  } catch(e) {
    // Ignore error so it doesn't break import
  }
  
  return "Berhasil! " + addedCount + " transaksi baru telah di-import ke sheet REKAPAN JAGO.";
}

function updateNeracaKeuangan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var jagoSheet = ss.getSheetByName('REKAPAN JAGO');
  var neracaSheet = ss.getSheetByName('Neraca Keuangan');
  
  if (!jagoSheet) return;
  
  if (!neracaSheet) {
    neracaSheet = ss.insertSheet('Neraca Keuangan');
  }
  
  var data = jagoSheet.getDataRange().getValues();
  if (data.length < 2) return;
  
  var firstRow = data[1];
  var firstAmount = parseFloat(firstRow[4]) || 0;
  var firstBalance = parseFloat(firstRow[5]) || 0;
  var saldoAwal = firstBalance - firstAmount;
  
  var lastRow = data[data.length - 1];
  var saldoAkhir = parseFloat(lastRow[5]) || 0;
  
  var units = {};
  
  for (var i = 1; i < data.length; i++) {
    var amount = parseFloat(data[i][4]) || 0;
    var unit = data[i][6];
    if (!unit || unit.toString().trim() === '') {
      unit = 'Tanpa Unit';
    } else {
      unit = unit.toString().trim();
    }
    
    if (!units[unit]) {
      units[unit] = { debit: 0, kredit: 0 };
    }
    
    if (amount > 0) {
      units[unit].debit += amount;
    } else if (amount < 0) {
      units[unit].kredit += Math.abs(amount);
    }
  }
  
  neracaSheet.clear();
  
  var neracaData = [];
  neracaData.push(["Nama Akun", "Debit", "Kredit"]);
  neracaData.push(["Saldo Awal", saldoAwal, ""]);
  
  var totalDebit = 0;
  var totalKredit = 0;
  
  for (var u in units) {
    var prefix = (u === 'Tanpa Unit') ? "" : "Operasional / BODP ";
    neracaData.push([prefix + u, units[u].debit || "", units[u].kredit || ""]);
    totalDebit += units[u].debit;
    totalKredit += units[u].kredit;
  }
  
  neracaData.push(["", "", ""]);
  neracaData.push(["Total Nominal", totalDebit, totalKredit]);
  neracaData.push(["Saldo Akhir", saldoAkhir, ""]);
  
  neracaSheet.getRange(1, 1, neracaData.length, 3).setValues(neracaData);
  
  neracaSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f4f6");
  neracaSheet.getRange(neracaData.length - 1, 1, 2, 3).setFontWeight("bold").setBackground("#e5e7eb");
  neracaSheet.getRange("A2:C2").setFontWeight("bold");
  
  neracaSheet.getRange(2, 2, neracaData.length - 1, 2).setNumberFormat('"Rp" #,##0');
  neracaSheet.autoResizeColumns(1, 3);
}
