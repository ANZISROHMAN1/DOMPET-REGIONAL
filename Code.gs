// FUNGSI INI WAJIB DI-RUN SATU KALI UNTUK MEMAKSA IZIN PENUH
function paksaIzin() {
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getRootFolder().getFiles();
  DriveApp.createFile("test", "test", MimeType.PLAIN_TEXT).setTrashed(true);
}

// FUNGSI AUTO-UPDATE JIKA ADA PERUBAHAN MANUAL DI SHEET
function onEdit(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() === 'REKAPAN JAGO WEB') {
    var header = sheet.getRange(1, e.range.getColumn()).getValue();
    if (String(header).toLowerCase().trim() === 'unit') {
      try { updateNeracaKeuangan(); } catch(err) {}
    }
  } else if (sheet.getName() === 'NERACA KEUANGAN WEB' && e.range.getColumn() === 5 && e.range.getRow() === 2) {
    try { updateNeracaKeuangan(); } catch(err) {}
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('FORM USER WEB') || ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var action = e.parameter.action;
  
  var expectedHeaders = ['ID', 'Tanggal', 'Nama', 'Kegiatan', 'Nominal', 'Bank', 'Rekening', 'Status', 'Unit', 'Sub_Unit', 'Bukti_Path', 'Bukti_TF_Path', 'File_Hash', 'Kategori'];
  
  var hasHeader = (data.length > 0 && String(data[0][0]).toUpperCase() === 'ID');
  var headers = hasHeader ? data[0] : expectedHeaders;
  var startIndex = hasHeader ? 1 : 0;
  
  var rows = [];
  for (var i = startIndex; i < data.length; i++) {
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
    var jagoSheet = ss.getSheetByName('REKAPAN JAGO WEB');
    if (!jagoSheet) return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Sheet REKAPAN JAGO tidak ditemukan'})).setMimeType(ContentService.MimeType.JSON);
    
    var CUTOFF_DATE = new Date(2026, 6, 30);
    var filterMonth = e.parameter.month;
    var filterYear = e.parameter.year;
    
    function parseDateSafe(val) {
      if (!val) return null;
      if (val instanceof Date) return val;
      var dStr = val.toString();
      var dateMatch = dStr.match(/(\d{1,2}\s+[a-zA-Z]{3}\s+\d{4})/);
      if (dateMatch) {
        var englishDateStr = dateMatch[1].replace('Mei', 'May').replace('Agu', 'Aug').replace('Okt', 'Oct').replace('Des', 'Dec');
        var d = new Date(englishDateStr);
        if (!isNaN(d.getTime())) return d;
      }
      var d2 = new Date(dStr);
      if (!isNaN(d2.getTime())) return d2;
      var d3 = new Date(dStr.replace(' ', 'T'));
      if (!isNaN(d3.getTime())) return d3;
      return null;
    }

    var jagoData = jagoSheet.getDataRange().getValues();
    var hMap = {};
    if (jagoData.length > 0) {
      for (var c = 0; c < jagoData[0].length; c++) {
        hMap[String(jagoData[0][c]).toLowerCase().trim()] = c;
      }
    }
    
    // Default fallback if headers are missing
    var colDate = hMap['date & time'] !== undefined ? hMap['date & time'] : 0;
    var colMasuk = hMap['kas masuk'] !== undefined ? hMap['kas masuk'] : 4;
    var colKeluar = hMap['kas keluar'] !== undefined ? hMap['kas keluar'] : 5;
    var colUnit = hMap['unit'] !== undefined ? hMap['unit'] : 7;
    var colKategori = hMap['kategori'] !== undefined ? hMap['kategori'] : 8;
    var colDetails = hMap['transaction details'] !== undefined ? hMap['transaction details'] : 2;
    var colNotes = hMap['notes'] !== undefined ? hMap['notes'] : 3;
    
    var isTwoColumnMode = (hMap['kas masuk'] !== undefined && hMap['kas keluar'] !== undefined);
    
    var summaryData = {};
    var categorySummary = {};
    var trendData = {};
    var tagihanRutin = [];
    var monthsSet = {};
    var yearsSet = {};
    
    for (var k = 1; k < jagoData.length; k++) {
      var rowDate = parseDateSafe(jagoData[k][colDate]);
      if (!rowDate || rowDate < CUTOFF_DATE) continue;
      
      var rowMonthOnly = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MMMM");
      rowMonthOnly = rowMonthOnly.replace('January', 'Januari').replace('February', 'Februari').replace('March', 'Maret').replace('May', 'Mei').replace('June', 'Juni').replace('July', 'Juli').replace('August', 'Agustus').replace('October', 'Oktober').replace('December', 'Desember');
      var rowYearOnly = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy");
      
      var monthYear = rowMonthOnly + " " + rowYearOnly;
      monthsSet[rowMonthOnly] = true;
      yearsSet[rowYearOnly] = true;
      
      var amount = 0;
      var unit = 'Tanpa Unit';
      var kategori = 'Lainnya';
      
      if (isTwoColumnMode) {
        var masukVal = parseFloat(jagoData[k][colMasuk]) || 0;
        var keluarVal = parseFloat(jagoData[k][colKeluar]) || 0;
        amount = (masukVal > 0) ? masukVal : ((keluarVal > 0) ? -keluarVal : 0);
        unit = jagoData[k][colUnit] || 'Tanpa Unit';
        kategori = jagoData[k][colKategori] || 'Lainnya';
      } else {
        amount = parseFloat(jagoData[k][4]) || 0;
        unit = jagoData[k][6] || 'Tanpa Unit';
        kategori = jagoData[k][7] || 'Lainnya';
      }
      var details = jagoData[k][colDetails] || '';
      var notes = jagoData[k][colNotes] || '';
      
      if (filterMonth && filterMonth !== "Semua Bulan") {
          if (rowMonthOnly !== filterMonth) continue;
      }
      if (filterYear && filterYear !== "Semua Tahun") {
          if (rowYearOnly !== filterYear) continue;
      }

      // Hitung Trend Data (Sesuai dengan filter)
      var isSpecificMonth = (filterMonth && filterMonth !== "Semua Bulan");
      var trendKey;
      var sortKey;
      if (isSpecificMonth) {
          trendKey = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd MMM");
          sortKey = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyyMMdd");
      } else {
          trendKey = rowMonthOnly + " " + rowYearOnly;
          sortKey = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyyMM");
      }

      if (!trendData[trendKey]) trendData[trendKey] = { income: 0, expense: 0, sortKey: sortKey };
      if (amount > 0) trendData[trendKey].income += amount;
      if (amount < 0) trendData[trendKey].expense += Math.abs(amount);
      
      if (amount < 0) { // Only count expenses for summaries
        var absAmount = Math.abs(amount);
        
        // Unit summary
        if (!summaryData[unit]) summaryData[unit] = { count: 0, total: 0, transactions: [] };
        summaryData[unit].count += 1;
        summaryData[unit].total += absAmount;
        summaryData[unit].transactions.push({
            date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd MMM yyyy"),
            details: details + (notes ? ' - ' + notes : ''),
            amount: absAmount,
            kategori: kategori
        });
        
        // Category summary
        if (!categorySummary[kategori]) categorySummary[kategori] = { count: 0, total: 0 };
        categorySummary[kategori].count += 1;
        categorySummary[kategori].total += absAmount;
        
        // Deteksi Tagihan Rutin
        var combinedText = (details + ' ' + notes).toLowerCase();
        if (kategori.includes("Tagihan Rutin") || combinedText.includes("listrik") || combinedText.includes("internet") || combinedText.includes("wifi") || combinedText.includes("pdam") || combinedText.includes("sewa")) {
            tagihanRutin.push({ date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd MMM yyyy"), details: details + ' ' + notes, amount: absAmount, unit: unit, kategori: kategori });
        }
      }
    }
    
    var resultData = [];
    for (var u in summaryData) {
      resultData.push({ unit: u, count: summaryData[u].count, total: summaryData[u].total, transactions: summaryData[u].transactions });
    }
    
    var resultCategory = [];
    for (var c in categorySummary) {
      resultCategory.push({ kategori: c, count: categorySummary[c].count, total: categorySummary[c].total });
    }
    
    var resultTrend = [];
    for (var m in trendData) {
      resultTrend.push({ month: m, income: trendData[m].income, expense: trendData[m].expense, sortKey: trendData[m].sortKey });
    }
    resultTrend.sort(function(a, b) {
      return a.sortKey.localeCompare(b.sortKey);
    });
    
    var availableMonths = ["Semua Bulan"].concat(Object.keys(monthsSet));
    var availableYears = ["Semua Tahun"].concat(Object.keys(yearsSet));
    var currentMonth = filterMonth || "Semua Bulan";
    var currentYear = filterYear || "Semua Tahun";
    
    return ContentService.createTextOutput(JSON.stringify({
        success: true, 
        data: resultData,
        categoryData: resultCategory,
        trendData: resultTrend,
        tagihanRutin: tagihanRutin,
        months: availableMonths,
        years: availableYears,
        currentMonth: currentMonth,
        currentYear: currentYear
    })).setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'neraca_data') {
    var jagoSheet = ss.getSheetByName('REKAPAN JAGO WEB');
    if (!jagoSheet) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Sheet REKAPAN JAGO tidak ditemukan.'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    var filterMonth = e.parameter.month || "Semua Bulan";
    var jagoData = jagoSheet.getDataRange().getValues();
    
    try {
      var calcResult = calculateNeraca(jagoData, filterMonth);
      var neracaData = calcResult.data;
      var availableMonths = ["Semua Bulan"].concat(calcResult.availableMonths);
      
      if (neracaData.length === 0) {
        return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Data kosong'})).setMimeType(ContentService.MimeType.JSON);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
          success: true, 
          data: neracaData,
          months: availableMonths,
          currentMonth: filterMonth
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
          success: false, 
          message: 'Error di calculateNeraca: ' + err.message + ' at line ' + err.lineNumber
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } else if (action === 'debug_jago') {
    var jagoSheet = ss.getSheetByName('REKAPAN JAGO WEB');
    var jagoData = jagoSheet ? jagoSheet.getDataRange().getValues() : [];
    var filterMonth = e.parameter.month || "Semua Bulan";
    
    try {
      var calcResult = calculateNeraca(jagoData, filterMonth);
      return ContentService.createTextOutput(JSON.stringify({
          success: true, 
          data: jagoData.slice(0, 10),
          calcResult: calcResult
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
          success: false, 
          message: err.message
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } else if (action === 'tagihan_rutin_list') {
    var tagihanSheet = ss.getSheetByName('TAGIHAN RUTIN WEB');
    if (!tagihanSheet) return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Sheet TAGIHAN RUTIN WEB tidak ditemukan'})).setMimeType(ContentService.MimeType.JSON);
    
    var data = tagihanSheet.getDataRange().getValues();
    if (data.length < 2) return ContentService.createTextOutput(JSON.stringify({success: true, data: []})).setMimeType(ContentService.MimeType.JSON);
    
    var headerRowIndex = 0;
    for (var i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).trim() !== '' || String(data[i][1]).trim() !== '') {
        headerRowIndex = i;
        break;
      }
    }
    
    var headers = data[headerRowIndex];
    var rows = [];
    for (var i = headerRowIndex + 1; i < data.length; i++) {
      // Skip empty rows
      if (String(data[i][0]).trim() === '' && String(data[i][1]).trim() === '') continue;
      
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var headerName = String(headers[j]).trim();
        if (headerName === '') headerName = 'Column' + j;
        row[headerName] = data[i][j];
      }
      row['row_index'] = i + 1;
      rows.push(row);
    }
    rows.reverse(); // newest first
    return ContentService.createTextOutput(JSON.stringify({success: true, data: rows})).setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'get_planning') {
    var plans = getPlanningData();
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: plans })).setMimeType(ContentService.MimeType.JSON);
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
    var sheet = ss.getSheetByName('FORM USER WEB') || ss.getActiveSheet();
    var sheetJago = ss.getSheetByName('REKAPAN JAGO WEB');
    
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
      
      var rowData = [newId, now, requestData.nama, requestData.kegiatan, requestData.nominal, requestData.bank, requestData.rekening, 'Pending', requestData.unit, requestData.sub_unit, fileUrl, '', fileHash, requestData.kategori || ''];
      sheet.appendRow(rowData);
      
      // Menambahkan data ke sheet REKAPAN JAGO dengan format khusus
      if (sheetJago) {
        var sourceDest = requestData.nama + "\n" + requestData.bank + " " + requestData.rekening;
        var transDetails = "Claim ID# " + newId;
        var nominalStr = requestData.nominal ? requestData.nominal.toString().replace(/[^0-9]/g, '') : "0";
        var amount = -Math.abs(parseFloat(nominalStr)); // Pengeluaran (minus)
        
        // Dinamis mendapatkan kolom agar bebas ditaruh di mana saja
        var lastColJago = sheetJago.getLastColumn();
        var jagoHeaders = [];
        if (lastColJago > 0) {
          jagoHeaders = sheetJago.getRange(1, 1, 1, lastColJago).getValues()[0];
        } else {
          jagoHeaders = ['Date & Time', 'Source/Destination', 'Transaction Details', 'Notes', 'Kas Masuk', 'Kas Keluar', 'Saldo', 'Unit', 'Kategori'];
        }
        
        var jMap = {};
        for (var c = 0; c < jagoHeaders.length; c++) {
          jMap[String(jagoHeaders[c]).toLowerCase().trim()] = c;
        }
        
        var jagoDataToInsert = new Array(jagoHeaders.length).fill("");
        var setJagoCol = function(name, val) {
          var idx = jMap[name.toLowerCase()];
          if (idx !== undefined) jagoDataToInsert[idx] = val;
        };
        
        setJagoCol('date & time', now);
        setJagoCol('source/destination', sourceDest);
        setJagoCol('transaction details', transDetails);
        setJagoCol('notes', requestData.kegiatan);
        setJagoCol('kas keluar', Math.abs(amount)); // Pengeluaran (minus) masuk ke kas keluar
        setJagoCol('unit', requestData.unit);
        setJagoCol('kategori', requestData.kategori || '');
        
        // Mencari baris kosong pertama di kolom A (Menghindari bug appendRow jika ada ArrayFormula)
        var jagoColA = sheetJago.getRange("A:A").getValues();
        var jagoTargetRow = jagoColA.length + 1;
        for (var i = 0; i < jagoColA.length; i++) {
          if (jagoColA[i][0] === "" && i > 0) { // i > 0 untuk melewati header
            jagoTargetRow = i + 1;
            break;
          }
        }
        
        sheetJago.getRange(jagoTargetRow, 1, 1, jagoHeaders.length).setValues([jagoDataToInsert]);
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
    } else if (action === 'upload_bukti_tagihan') {
      var sheetTagihan = ss.getSheetByName('TAGIHAN RUTIN WEB');
      if (!sheetTagihan) return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Sheet tidak ditemukan'})).setMimeType(ContentService.MimeType.JSON);
      
      var rowIndex = parseInt(requestData.row_index);
      var fileUrl = '';
      if (requestData.fileData && requestData.fileName && requestData.mimeType) {
        var base64Data = requestData.fileData.split(',')[1] || requestData.fileData;
        var byteData = Utilities.base64Decode(base64Data);
        var blob = Utilities.newBlob(byteData, requestData.mimeType, requestData.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      }
      
      var allData = sheetTagihan.getDataRange().getValues();
      var headerRowIndex = 0;
      for (var i = 0; i < Math.min(5, allData.length); i++) {
        if (String(allData[i][0]).trim() !== '' || String(allData[i][1]).trim() !== '') {
          headerRowIndex = i;
          break;
        }
      }
      
      var headers = allData[headerRowIndex];
      var statusCol = -1;
      var buktiCol = -1;
      for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i]).toLowerCase();
        if (h.includes('status')) statusCol = i + 1;
        if (h.includes('bukti') || h.includes('transfer') || h.includes('tf')) buktiCol = i + 1;
      }
      
      // Jika kolom status dan bukti belum ada, tambahkan di sebelah kanan
      if (statusCol === -1) {
        statusCol = sheetTagihan.getLastColumn() + 1;
        sheetTagihan.getRange(headerRowIndex + 1, statusCol).setValue('Status');
      }
      if (buktiCol === -1) {
        buktiCol = statusCol + 1;
        sheetTagihan.getRange(headerRowIndex + 1, buktiCol).setValue('Bukti TF');
      }
      
      if (statusCol > 0) sheetTagihan.getRange(rowIndex, statusCol).setValue('Dibayar');
      if (buktiCol > 0 && fileUrl) sheetTagihan.getRange(rowIndex, buktiCol).setValue(fileUrl);
      
      return ContentService.createTextOutput(JSON.stringify({success: true, fileUrl: fileUrl})).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'save_planning') {
      var sheet = getOrCreatePlanningSheet();
      var id = 'PLN-' + new Date().getTime();
      sheet.appendRow([id, new Date(), requestData.planning, requestData.kategori, requestData.budget]);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Planning berhasil disimpan', id: id })).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'delete_planning') {
      var sheet = getOrCreatePlanningSheet();
      var data = sheet.getDataRange().getValues();
      var deleted = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === requestData.id) {
          sheet.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }
      if (deleted) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Planning berhasil dihapus' })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Planning tidak ditemukan' })).setMimeType(ContentService.MimeType.JSON);
      }
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function formatTransactions(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;
  
  // Format Header Kolom
  sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var masukCol = -1, keluarCol = -1, saldoCol = -1;
  
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toLowerCase().trim();
    if (h === 'kas masuk') masukCol = c + 1;
    if (h === 'kas keluar') keluarCol = c + 1;
    if (h === 'saldo') saldoCol = c + 1;
  }
  
  if (masukCol > 0) {
    sheet.getRange(1, masukCol).setValue("Kas Masuk").setBackground("#d9ead3").setFontColor("#137333");
  }
  if (keluarCol > 0) {
    sheet.getRange(1, keluarCol).setValue("Kas Keluar").setBackground("#f4cccc").setFontColor("#990000");
  }
  if (saldoCol > 0) {
    sheet.getRange(1, saldoCol).setValue("Saldo").setBackground("#cfe2f3").setFontColor("#0b5394");
  }
  
  if (lastRow < 2) return;
  
  var rangeMasuk = masukCol > 0 ? sheet.getRange(2, masukCol, lastRow - 1, 1) : null;
  var rangeKeluar = keluarCol > 0 ? sheet.getRange(2, keluarCol, lastRow - 1, 1) : null;
  var rangeSaldo = saldoCol > 0 ? sheet.getRange(2, saldoCol, lastRow - 1, 1) : null;
  
  if (rangeMasuk) rangeMasuk.setNumberFormat('"Rp" #,##0');
  if (rangeKeluar) rangeKeluar.setNumberFormat('"Rp" #,##0');
  if (rangeSaldo) rangeSaldo.setNumberFormat('"Rp" #,##0');
  
  var rules = sheet.getConditionalFormatRules();
  var newRules = rules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(function(r) { 
      var col = r.getColumn();
      return col === masukCol || col === keluarCol; 
    });
  });
  
  if (rangeMasuk) {
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#d9ead3")
      .setFontColor("#137333")
      .setRanges([rangeMasuk])
      .build());
  }
  if (rangeKeluar) {
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#f4cccc")
      .setFontColor("#990000")
      .setRanges([rangeKeluar])
      .build());
  }
  
  sheet.setConditionalFormatRules(newRules);
}

// --- TAMBAHAN UNTUK IMPORT PDF JAGO ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ FINEAST')
      .addItem('📄 Import Rekapan Jago (PDF)', 'showImportDialog')
      .addItem('🎨 Pisahkan Kolom Kas Masuk & Keluar', 'pisahkanKolomKasMasukKeluar')
      .addItem('🔄 Refresh Unit & Kategori (Deteksi DB)', 'refreshUnitFromDB')
      .addToUi();
}

function getMappings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var unitMap = {};
  var kategoriMap = {};
  var usedSheets = [];
  
  // 1. Coba ambil data Unit dari sheet 'UNIT' jika ada
  var unitSheet = ss.getSheetByName('UNIT') || ss.getSheetByName('UNIT ');
  if (unitSheet) {
    usedSheets.push(unitSheet.getName());
    var unitData = unitSheet.getDataRange().getValues();
    if (unitData.length > 0) {
      var uHeaders = unitData[0].map(function(c) { return String(c).toUpperCase().replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim(); });
      
      var uKeyIdx = -1, uValIdx = -1;
      for (var j = 0; j < uHeaders.length; j++) {
        var h = uHeaders[j];
        if (h.includes('TUJUAN') || h.includes('REK')) uKeyIdx = j;
        else if (h.includes('UNIT') || h.includes('CABANG')) uValIdx = j;
      }
      
      if (uKeyIdx === -1) uKeyIdx = 0;
      if (uValIdx === -1) uValIdx = 1;
      if (uKeyIdx === uValIdx) { uKeyIdx = 0; uValIdx = 1; }
      
      for (var r = 1; r < unitData.length; r++) {
        var uKey = String(unitData[r][uKeyIdx] !== undefined ? unitData[r][uKeyIdx] : "").trim();
        var uVal = String(unitData[r][uValIdx] !== undefined ? unitData[r][uValIdx] : "").trim();
        
        if (uKey.length > 1 && uVal && !uKey.toUpperCase().includes('KEYWORD')) {
           unitMap[uKey] = uVal;
        } else if (uKey.toLowerCase() === 'selainnya') {
           unitMap[uKey] = uVal;
        }
      }
    }
  }

  // 2. Coba ambil data Kategori dari sheet 'KATEGORI' jika ada
  var katSheet = ss.getSheetByName('KATEGORI') || ss.getSheetByName('KATEGORI ');
  if (katSheet) {
    usedSheets.push(katSheet.getName());
    var katData = katSheet.getDataRange().getValues();
    if (katData.length > 0) {
      var kHeaders = katData[0].map(function(c) { return String(c).toUpperCase().replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim(); });
      
      var kKeyIdx = -1, kValIdx = -1;
      for (var j = 0; j < kHeaders.length; j++) {
        var h = kHeaders[j];
        if (h.includes('TRANSAKSI') || (h.includes('KEYWORD') && !h.includes('REK'))) kKeyIdx = j;
        else if (h.includes('KATEGORI')) kValIdx = j;
      }
      
      if (kKeyIdx === -1) kKeyIdx = 0;
      if (kValIdx === -1) kValIdx = 1;
      if (kKeyIdx === kValIdx) { kKeyIdx = 0; kValIdx = 1; }
      
      for (var r = 1; r < katData.length; r++) {
        var kKey = String(katData[r][kKeyIdx] !== undefined ? katData[r][kKeyIdx] : "").trim();
        var kVal = String(katData[r][kValIdx] !== undefined ? katData[r][kValIdx] : "").trim();
        
        if (kKey.length > 1 && kVal && !kKey.toUpperCase().includes('KEYWORD')) {
           kategoriMap[kKey] = kVal;
        } else if (kKey.toLowerCase() === 'selainnya') {
           kategoriMap[kKey] = kVal;
        }
      }
    }
  }

  // 3. Fallback ambil dari 'DB' / 'MASTER' jika masih ada yang kosong
  if (!unitSheet || !katSheet) {
    var sheetNames = ['DB', 'DB ', 'MASTER', 'MASTER '];
    var dbSheet = null;
    for (var i = 0; i < sheetNames.length; i++) {
      var sh = ss.getSheetByName(sheetNames[i]);
      if (sh) {
        dbSheet = sh;
        break; 
      }
    }
    
    if (dbSheet) {
      if (!usedSheets.includes(dbSheet.getName())) {
        usedSheets.push(dbSheet.getName());
      }
      var dbData = dbSheet.getDataRange().getValues();
      var headerRowIdx = -1;
      var headers = [];
      
      for (var r = 0; r < Math.min(dbData.length, 100); r++) {
        var rowStrs = dbData[r].map(function(c) { return String(c).toUpperCase().replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim(); });
        if ((!katSheet && (rowStrs.includes('KEYWORD TRANSAKSI JAGO') || rowStrs.includes('KATEGORI'))) || (!unitSheet && rowStrs.includes('UNIT'))) {
          headerRowIdx = r;
          headers = rowStrs;
          break;
        }
      }
      
      if (headerRowIdx === -1) {
        headerRowIdx = 0;
        headers = dbData[0].map(function(c) { return String(c).toUpperCase().replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim(); });
      }
      
      var uKeyIdx2 = -1, uValIdx2 = -1, kKeyIdx2 = -1, kValIdx2 = -1;
      for (var j = 0; j < headers.length; j++) {
        var h = headers[j];
        if (!unitSheet && (h.includes('TUJUAN') || h.includes('REK'))) uKeyIdx2 = j;
        else if (!unitSheet && (h.includes('UNIT') || h.includes('CABANG'))) uValIdx2 = j;
        
        if (!katSheet && (h.includes('TRANSAKSI') || (h.includes('KEYWORD') && !h.includes('REK')))) kKeyIdx2 = j;
        else if (!katSheet && h.includes('KATEGORI')) kValIdx2 = j;
      }
      
      if (!unitSheet) {
        if (uKeyIdx2 === -1) uKeyIdx2 = 0;
        if (uValIdx2 === -1) uValIdx2 = 1;
        if (uKeyIdx2 === uValIdx2) { uKeyIdx2 = 0; uValIdx2 = 1; }
      }
      if (!katSheet) {
        if (kKeyIdx2 === -1) kKeyIdx2 = dbSheet.getLastColumn() >= 6 ? 4 : 2;
        if (kValIdx2 === -1) kValIdx2 = dbSheet.getLastColumn() >= 6 ? 5 : 3;
        if (kKeyIdx2 === kValIdx2) { kKeyIdx2 = 4; kValIdx2 = 5; }
      }
      
      for (var r = headerRowIdx + 1; r < dbData.length; r++) {
        if (!unitSheet) {
          var uKey = String(dbData[r][uKeyIdx2] !== undefined ? dbData[r][uKeyIdx2] : "").trim();
          var uVal = String(dbData[r][uValIdx2] !== undefined ? dbData[r][uValIdx2] : "").trim();
          if (uKey.length > 1 && uVal && !uKey.toUpperCase().includes('KEYWORD')) {
             unitMap[uKey] = uVal;
          } else if (uKey.toLowerCase() === 'selainnya') {
             unitMap[uKey] = uVal;
          }
        }
        
        if (!katSheet) {
          var kKey = String(dbData[r][kKeyIdx2] !== undefined ? dbData[r][kKeyIdx2] : "").trim();
          var kVal = String(dbData[r][kValIdx2] !== undefined ? dbData[r][kValIdx2] : "").trim();
          if (kKey.length > 1 && kVal && !kKey.toUpperCase().includes('KEYWORD')) {
             kategoriMap[kKey] = kVal;
          } else if (kKey.toLowerCase() === 'selainnya') {
             kategoriMap[kKey] = kVal;
          }
        }
      }
    }
  }
  
  var finalSheetName = usedSheets.length > 0 ? usedSheets.join(" & ") : "None";
  return { unitMap: unitMap, kategoriMap: kategoriMap, sheetName: finalSheetName };
}

function refreshUnitFromDB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('REKAPAN JAGO WEB');
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'REKAPAN JAGO WEB' tidak ditemukan!");
    return;
  }
  
  var maps = getMappings();
  var unitMap = maps.unitMap;
  var kategoriMap = maps.kategoriMap;
  
  if (maps.sheetName === "None") {
    SpreadsheetApp.getUi().alert("Gagal menemukan sheet yang berisi kolom KEYWORD TRANSAKSI JAGO atau KATEGORI.");
    return;
  }

  function guessUnit(details, notes) {
    var text = (details + " " + notes).toLowerCase();
    for (var key in unitMap) {
      if (key.toLowerCase() !== 'selainnya' && text.includes(key.toLowerCase())) {
        return unitMap[key];
      }
    }
    
    for (var key in unitMap) {
      if (key.toLowerCase() === 'selainnya') return unitMap[key];
    }
    return "EASTERN";
  }
  
  function guessKategori(details, notes) {
    var text = (details + " " + notes).toLowerCase();
    for (var key in kategoriMap) {
      if (key.toLowerCase() !== 'selainnya' && text.includes(key.toLowerCase())) {
        return kategoriMap[key];
      }
    }
    
    for (var key in kategoriMap) {
      if (key.toLowerCase() === 'selainnya') return kategoriMap[key];
    }
    return ""; // Kosongkan jika tidak ada keyword yang cocok dan tidak ada keyword 'selainnya'
  }

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 2) return;
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = {};
  for (var c = 0; c < headers.length; c++) {
    var hName = String(headers[c]).toLowerCase().trim();
    if (colMap[hName] === undefined) {
      colMap[hName] = c;
    }
  }
  
  var unitColIdx = colMap['unit'];
  var katColIdx = colMap['kategori'];
  var sourceDestIdx = colMap['source/destination'];
  var detailsIdx = colMap['transaction details'];
  var notesIdx = colMap['notes'];
  
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var updatedCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var sourceDest = sourceDestIdx !== undefined ? String(data[i][sourceDestIdx]) : "";
    var details = detailsIdx !== undefined ? String(data[i][detailsIdx]) : "";
    var notes = notesIdx !== undefined ? String(data[i][notesIdx]) : "";
    
    var isUpdated = false;
    
    if (unitColIdx !== undefined) {
      var currentUnit = String(data[i][unitColIdx]).trim();
      var autoUnit = guessUnit(sourceDest + " " + details, notes);
      if (autoUnit !== currentUnit) {
        data[i][unitColIdx] = autoUnit;
        isUpdated = true;
      }
    }
    
    if (katColIdx !== undefined) {
      var currentKat = String(data[i][katColIdx]).trim();
      var autoKat = guessKategori(sourceDest + " " + details, notes);
      if (autoKat !== currentKat) {
        data[i][katColIdx] = autoKat;
        isUpdated = true;
      }
    }
    
    if (isUpdated) updatedCount++;
  }
  
  if (updatedCount > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).setValues(data);
    try { updateNeracaKeuangan(); } catch(e) {}
    SpreadsheetApp.getUi().alert("✅ Berhasil! " + updatedCount + " baris telah diperbarui (Unit & Kategori) berdasarkan sheet DB terbaru.");
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Semua data sudah sesuai dengan Unit & Kategori dari DB.");
  }
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
                  
                  // Filter header & footer PDF Jago
                  var lowerText = text.toLowerCase();
                  if (lowerText === "date & time" || 
                      lowerText === "source/destination" || 
                      lowerText === "transaction details" || 
                      lowerText === "notes" || 
                      lowerText === "amount" || 
                      lowerText === "balance" || 
                      lowerText.match(/^page\s+\d+/) || 
                      lowerText.match(/^pockets transactions history/) ||
                      lowerText === "jago") {
                      return;
                  }
                  
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
                    if (pct < 0.135) col1.push(it.text); // Date & Time
                    else if (pct < 0.31) col2.push(it.text); // Source/Destination
                    else if (pct < 0.47) col3.push(it.text); // Transaction Details
                    else if (pct < 0.65) col4.push(it.text); // Notes
                    else if (pct < 0.82) col5.push(it.text); // Amount
                    else col6.push(it.text); // Balance
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
  var sheet = ss.getSheetByName('REKAPAN JAGO WEB');
  if (!sheet) throw new Error("Sheet bernama 'REKAPAN JAGO' tidak ditemukan!");
  
  var maps = getMappings();
  var unitMap = maps.unitMap;
  var kategoriMap = maps.kategoriMap;

  // HAPUS SEMUA DATA LAMA (Karena kita akan timpa / overwrite sepenuhnya dari PDF)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  function guessUnit(details, notes) {
    var text = (details + " " + notes).toLowerCase();
    for (var key in unitMap) {
      if (key.toLowerCase() !== 'selainnya' && text.includes(key.toLowerCase())) return unitMap[key];
    }
    
    for (var key in unitMap) {
      if (key.toLowerCase() === 'selainnya') return unitMap[key];
    }
    return "EASTERN";
  }
  
  function guessKategori(details, notes) {
    var text = (details + " " + notes).toLowerCase();
    for (var key in kategoriMap) {
      if (key.toLowerCase() !== 'selainnya' && text.includes(key.toLowerCase())) return kategoriMap[key];
    }
    
    for (var key in kategoriMap) {
      if (key.toLowerCase() === 'selainnya') return kategoriMap[key];
    }
    return ""; // Kosongkan jika tidak ada keyword yang cocok
  }
  
  // --- BATAS TANGGAL IMPORT ---
  // 30 Juli 2026 (Bulan di JavaScript dimulai dari 0, jadi 6 = Juli)
  var CUTOFF_DATE = new Date(2026, 6, 30); 
  
  // Baca header yang ada saat ini
  var lastColSheet = sheet.getLastColumn();
  var existingHeaders = [];
  if (lastColSheet > 0) {
    existingHeaders = sheet.getRange(1, 1, 1, lastColSheet).getValues()[0];
  }
  // Jika sheet kosong, gunakan default
  if (existingHeaders.length === 0 || existingHeaders.join('').trim() === '') {
    existingHeaders = ['Date & Time', 'Source/Destination', 'Transaction Details', 'Notes', 'Kas Masuk', 'Kas Keluar', 'Saldo', 'Unit', 'Kategori'];
  }
  
  var colMap = {};
  for (var c = 0; c < existingHeaders.length; c++) {
    var hName = String(existingHeaders[c]).toLowerCase().trim();
    if (colMap[hName] === undefined) {
      colMap[hName] = c;
    }
  }

  var newRows = [];
  var addedCount = 0;
  
  for (var i = 0; i < transactions.length; i++) {
    var trx = transactions[i];
    if (!trx || trx.length < 5) continue;
    
    var dateStr = trx[0];
    
    // Filter Tanggal (Ekstrak dari teks kotor)
    var dateMatch = dateStr.match(/(\d{1,2}\s+[a-zA-Z]{3}\s+\d{4})/);
    if (!dateMatch) continue; // Bukan baris transaksi (misal: header/footer)
    
    var englishDateStr = dateMatch[1].replace('Mei', 'May').replace('Agu', 'Aug').replace('Okt', 'Oct').replace('Des', 'Dec');
    var transDate = new Date(englishDateStr);
    if (!isNaN(transDate.getTime()) && transDate < CUTOFF_DATE) {
      continue; // Lewati transaksi sebelum 30 Juli 2026
    }
    
    // Format tanggal menjadi dd/MM/yyyy
    var dd = String(transDate.getDate()).padStart(2, '0');
    var mm = String(transDate.getMonth() + 1).padStart(2, '0');
    var yyyy = transDate.getFullYear();
    var formattedDate = dd + '/' + mm + '/' + yyyy;
    
    // Extract time if exists in original string to append it (optional, but good to preserve)
    var timeMatch = dateStr.match(/\d{2}:\d{2}/);
    if (timeMatch) {
      formattedDate += ' ' + timeMatch[0];
    }
    
    // Pastikan baris ini adalah transaksi asli (harus punya ID#)
    if (!trx[1].includes("ID#") && !trx[2].includes("ID#")) {
      continue; // Abaikan header tabel
    }
    
    var notesAndMoney = (trx[3] + " " + trx[4] + " " + trx[5]).trim();
    var tokens = notesAndMoney.split(/\s+/);
    
    if (tokens.length < 2) continue; // Invalid row
    
    var balanceStr = tokens.pop();
    var amountStr = tokens.pop();
    var notes = tokens.join(" "); // Anything left over belongs to Notes
    
    if (amountStr.toLowerCase().includes('amount')) continue;
    
    var isNegative = false;
    var dashRegex = /[-–—−]/;
    
    if (notes.trim().match(/[-–—−]$/)) {
      isNegative = true;
      notes = notes.replace(/[-–—−]$/, '').trim();
    } else if (amountStr.match(/^[-–—−]/)) {
      isNegative = true;
      amountStr = amountStr.replace(/^[-–—−]/, '');
    } else if (amountStr.match(/[-–—−]$/)) {
      isNegative = true;
      amountStr = amountStr.replace(/[-–—−]$/, '');
    } else if (amountStr.startsWith('(') && amountStr.endsWith(')')) {
      isNegative = true;
      amountStr = amountStr.slice(1, -1);
    }
    
    // Clean numbers
    var amount = parseFloat(amountStr.replace(/\./g, '').replace(/,/g, '.'));
    if (isNegative) amount = -Math.abs(amount);
    var balance = parseFloat(balanceStr.replace(/\./g, '').replace(/,/g, '.'));
    
    // Tebak Unit secara otomatis!
    var autoUnit = guessUnit(trx[1] + " " + trx[2], notes);
    var autoKategori = guessKategori(trx[1] + " " + trx[2], notes);
    
    var kasMasuk = (amount > 0) ? amount : '';
    var kasKeluar = (amount < 0) ? Math.abs(amount) : '';
    
    var rowData = new Array(existingHeaders.length).fill("");
    var setCol = function(name, val) {
      var idx = colMap[name.toLowerCase()];
      if (idx !== undefined) rowData[idx] = val;
    };
    
    setCol('date & time', formattedDate);
    setCol('source/destination', trx[1]);
    setCol('transaction details', trx[2]);
    setCol('notes', notes);
    setCol('kas masuk', kasMasuk);
    setCol('kas keluar', kasKeluar);
    setCol('saldo', balance);
    setCol('unit', autoUnit);
    setCol('kategori', autoKategori);
    
    newRows.push(rowData);
    addedCount++;
  }
  
  // Tulis Header (Tetap pertahankan urutan yang ada)
  sheet.getRange(1, 1, 1, existingHeaders.length).setValues([existingHeaders]);
  
  // Tulis sekaligus semua data ke Sheet (Jauh lebih cepat dari appendRow)
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, existingHeaders.length).setValues(newRows);
  }
  
  formatTransactions(sheet);
  
  try {
    updateNeracaKeuangan();
  } catch(e) {
    // Ignore error so it doesn't break import
  }
  
  return "Berhasil! " + addedCount + " transaksi baru telah di-import ke sheet REKAPAN JAGO.";
}

// Fungsi Otomatis untuk Mengonversi Sheet Lama Menjadi Format 2 Kolom (Kas Masuk & Kas Keluar)
function pisahkanKolomKasMasukKeluar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('REKAPAN JAGO WEB');
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'REKAPAN JAGO WEB' tidak ditemukan!");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("Sheet masih kosong!");
    return;
  }
  
  var h4 = String(data[0][4] || '').toLowerCase();
  var h5 = String(data[0][5] || '').toLowerCase();
  
  // Jika sudah 9 kolom dan kolom E sudah Kas Masuk, kita hanya re-format saja
  if (h4.includes('masuk') && h5.includes('keluar')) {
    formatTransactions(sheet);
    SpreadsheetApp.getUi().alert("Kolom sudah terpisah sebelumnya. Format warna dan angka telah diperbarui!");
    return;
  }
  
  // Konversi dari format 8 kolom (Amount gabungan di Kolom E) ke format 9 kolom (Kas Masuk di E, Kas Keluar di F)
  var newRows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var amount = parseFloat(row[4]) || 0;
    var kasMasuk = (amount > 0) ? amount : '';
    var kasKeluar = (amount < 0) ? Math.abs(amount) : '';
    var saldo = row[5];
    var unit = row[6] || '';
    var kategori = row[7] || '';
    
    newRows.push([
      row[0], // Date & Time
      row[1], // Source/Dest
      row[2], // Details
      row[3], // Notes
      kasMasuk, // Kas Masuk (E)
      kasKeluar, // Kas Keluar (F)
      saldo, // Saldo (G)
      unit, // Unit (H)
      kategori // Kategori (I)
    ]);
  }
  
  // Bersihkan sheet
  sheet.clearContents();
  
  // Tulis Header Baru
  var headers = ['Date & Time', 'Source/Destination', 'Transaction Details', 'Notes', 'Kas Masuk', 'Kas Keluar', 'Saldo', 'Unit', 'Kategori'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Tulis Data Baru
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
  }
  
  formatTransactions(sheet);
  
  try {
    updateNeracaKeuangan();
  } catch(e) {}
  
  SpreadsheetApp.getUi().alert("Berhasil memisahkan kolom Kas Masuk (Hijau) dan Kas Keluar (Merah)!");
}

function calculateNeraca(jagoData, filterMonth) {
  if (jagoData.length < 2) return { data: [], availableMonths: [] };
  
  var CUTOFF_DATE = new Date(2026, 6, 30); // 30 Juli 2026
  function parseDateSafe(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    var dStr = val.toString();
    var dateMatch = dStr.match(/(\d{1,2}\s+[a-zA-Z]{3}\s+\d{4})/);
    if (dateMatch) {
      var englishDateStr = dateMatch[1].replace('Mei', 'May').replace('Agu', 'Aug').replace('Okt', 'Oct').replace('Des', 'Dec');
      var d = new Date(englishDateStr);
      if (!isNaN(d.getTime())) return d;
    }
    var d2 = new Date(dStr);
    if (!isNaN(d2.getTime())) return d2;
    var d3 = new Date(dStr.replace(' ', 'T'));
    if (!isNaN(d3.getTime())) return d3;
    return null;
  }

  var allValidData = [];
  var monthsSet = {};
  
  for (var i = 1; i < jagoData.length; i++) {
    if (!jagoData[i][0] || jagoData[i][0] === "") continue;
    var rowDate = parseDateSafe(jagoData[i][0]);
    if (rowDate && rowDate >= CUTOFF_DATE) {
      allValidData.push({row: jagoData[i], date: rowDate});
      
      var monthYear = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MMMM yyyy");
      monthYear = monthYear.replace('January', 'Januari').replace('February', 'Februari').replace('March', 'Maret').replace('May', 'Mei').replace('June', 'Juni').replace('July', 'Juli').replace('August', 'Agustus').replace('October', 'Oktober').replace('December', 'Desember');
      monthsSet[monthYear] = true;
    }
  }
  
  var availableMonths = Object.keys(monthsSet);
  var selectedFilter = filterMonth ? filterMonth.trim().replace(/\s+/g, ' ') : "Semua Bulan";
  
  if (selectedFilter !== "Semua Bulan") {
      var parts = selectedFilter.split(" ");
      if (parts.length === 1) {
          for (var i = 0; i < availableMonths.length; i++) {
              if (availableMonths[i].toLowerCase().indexOf(selectedFilter.toLowerCase()) !== -1) {
                  selectedFilter = availableMonths[i]; 
                  break;
              }
          }
      }
  }
  
  var filteredData = [];
  for (var i = 0; i < allValidData.length; i++) {
      var item = allValidData[i];
      var monthYear = Utilities.formatDate(item.date, Session.getScriptTimeZone(), "MMMM yyyy");
      monthYear = monthYear.replace('January', 'Januari').replace('February', 'Februari').replace('March', 'Maret').replace('May', 'Mei').replace('June', 'Juni').replace('July', 'Juli').replace('August', 'Agustus').replace('October', 'Oktober').replace('December', 'Desember');
      
      if (selectedFilter === "Semua Bulan") {
          filteredData.push(item.row);
      } else if (monthYear.toLowerCase() === selectedFilter.toLowerCase() || monthYear.toLowerCase().indexOf(selectedFilter.toLowerCase()) !== -1) {
          filteredData.push(item.row);
      }
  }
  
  if (filteredData.length === 0) {
      return { data: [["Tidak ada data untuk " + selectedFilter]], availableMonths: availableMonths }; 
  }
  
  var headers = jagoData.length > 0 ? jagoData[0] : [];
  var hMap = {};
  for (var c = 0; c < headers.length; c++) {
    hMap[String(headers[c]).toLowerCase().trim()] = c;
  }
  
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
  
  var isTwoColumnMode = (hMap['kas masuk'] !== undefined && hMap['kas keluar'] !== undefined);
  var balColIdx = hMap['saldo'] !== undefined ? hMap['saldo'] : (isTwoColumnMode ? 6 : 5);
  var unitColIdx = hMap['unit'] !== undefined ? hMap['unit'] : (isTwoColumnMode ? 7 : 6);
  var masukColIdx = hMap['kas masuk'] !== undefined ? hMap['kas masuk'] : 4;
  var keluarColIdx = hMap['kas keluar'] !== undefined ? hMap['kas keluar'] : 5;
  var detailColIdx = hMap['transaction details'] !== undefined ? hMap['transaction details'] : 2;
  var notesColIdx = hMap['notes'] !== undefined ? hMap['notes'] : 3;

  function getRawAmount(row) {
    if (isTwoColumnMode) {
      var m = parseSafeFloat(row[masukColIdx]);
      var k = parseSafeFloat(row[keluarColIdx]);
      return (m > 0) ? m : ((k > 0) ? -k : 0);
    }
    return parseSafeFloat(row[4]);
  }
  
  var isNewestToOldest = false;
  if (filteredData.length >= 2) {
      for (var i = 0; i < filteredData.length - 1; i++) {
          var b0 = parseSafeFloat(filteredData[i][balColIdx]);
          var b1 = parseSafeFloat(filteredData[i+1][balColIdx]);
          var a0 = Math.abs(getRawAmount(filteredData[i]));
          var a1 = Math.abs(getRawAmount(filteredData[i+1]));
          
          if (Math.abs(a0 - a1) > 1) {
              if (a0 > 0 && Math.abs(Math.abs(b0 - b1) - a0) < 1) {
                  isNewestToOldest = true; 
                  break;
              }
              if (a1 > 0 && Math.abs(Math.abs(b1 - b0) - a1) < 1) {
                  isNewestToOldest = false;
                  break;
              }
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
              if (Math.abs(Math.abs(diff) - amountMag) < 1) {
                  trueAmt = diff;
              } else if (diff < 0 && rawAmount > 0) {
                  trueAmt = -amountMag; 
              } else if (diff > 0 && rawAmount < 0) {
                  trueAmt = amountMag;
              }
          }
      } else if (!isNewestToOldest && i > 0) {
          var olderBalance = parseSafeFloat(filteredData[i-1][balColIdx]);
          if (currentBalance > 0 && olderBalance > 0) {
              var diff = currentBalance - olderBalance;
              if (Math.abs(Math.abs(diff) - amountMag) < 1) {
                  trueAmt = diff;
              } else if (diff < 0 && rawAmount > 0) {
                  trueAmt = -amountMag;
              } else if (diff > 0 && rawAmount < 0) {
                  trueAmt = amountMag;
              }
          }
      }
      
      trueAmounts[i] = trueAmt;
  }
  
  function getMonthIndex(monthName) {
    var m = monthName.toLowerCase();
    if (m.indexOf('jan') !== -1) return 0;
    if (m.indexOf('feb') !== -1) return 1;
    if (m.indexOf('mar') !== -1) return 2;
    if (m.indexOf('apr') !== -1) return 3;
    if (m.indexOf('mei') !== -1 || m.indexOf('may') !== -1) return 4;
    if (m.indexOf('jun') !== -1) return 5;
    if (m.indexOf('jul') !== -1) return 6;
    if (m.indexOf('agu') !== -1 || m.indexOf('aug') !== -1) return 7;
    if (m.indexOf('sep') !== -1) return 8;
    if (m.indexOf('okt') !== -1 || m.indexOf('oct') !== -1) return 9;
    if (m.indexOf('nov') !== -1) return 10;
    if (m.indexOf('des') !== -1 || m.indexOf('dec') !== -1) return 11;
    return 0;
  }

  var saldoAwal = 0;
  if (selectedFilter === "Semua Bulan") {
      if (isNewestToOldest) {
          for (var i = filteredData.length - 1; i >= 0; i--) {
              if (filteredData[i][balColIdx] !== "" && filteredData[i][balColIdx] !== undefined) {
                  saldoAwal = parseSafeFloat(filteredData[i][balColIdx]) - trueAmounts[i];
                  break;
              }
          }
      } else {
          for (var i = 0; i < filteredData.length; i++) {
              if (filteredData[i][balColIdx] !== "" && filteredData[i][balColIdx] !== undefined) {
                  saldoAwal = parseSafeFloat(filteredData[i][balColIdx]) - trueAmounts[i];
                  break;
              }
          }
      }
  } else {
      var parts = selectedFilter.split(" ");
      var targetMonthStart = new Date(parseInt(parts[1]), getMonthIndex(parts[0]), 1);
      var foundPrevBalance = false;
      
      if (isNewestToOldest) {
          for (var i = 0; i < allValidData.length; i++) {
              if (allValidData[i].date < targetMonthStart) {
                  var bal = allValidData[i].row[balColIdx];
                  if (bal !== "" && bal !== undefined) {
                      saldoAwal = parseSafeFloat(bal);
                      foundPrevBalance = true;
                      break;
                  }
              }
          }
      } else {
          for (var i = allValidData.length - 1; i >= 0; i--) {
              if (allValidData[i].date < targetMonthStart) {
                  var bal = allValidData[i].row[balColIdx];
                  if (bal !== "" && bal !== undefined) {
                      saldoAwal = parseSafeFloat(bal);
                      foundPrevBalance = true;
                      break;
                  }
              }
          }
      }
      
      if (!foundPrevBalance) {
          if (isNewestToOldest) {
              for (var i = filteredData.length - 1; i >= 0; i--) {
                  if (filteredData[i][balColIdx] !== "" && filteredData[i][balColIdx] !== undefined) {
                      saldoAwal = parseSafeFloat(filteredData[i][balColIdx]) - trueAmounts[i];
                      break;
                  }
              }
          } else {
              for (var i = 0; i < filteredData.length; i++) {
                  if (filteredData[i][balColIdx] !== "" && filteredData[i][balColIdx] !== undefined) {
                      saldoAwal = parseSafeFloat(filteredData[i][balColIdx]) - trueAmounts[i];
                      break;
                  }
              }
          }
      }
  }
  
  var units = {};
  var totalDebit = 0;
  var totalKredit = 0;
  
  for (var i = 0; i < filteredData.length; i++) {
    var amount = trueAmounts[i];
    var unit = filteredData[i][unitColIdx];
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
      totalKredit += amount; 
    } else if (amount < 0) {
      var absAmount = Math.abs(amount);
      units[unit].kredit += absAmount; 
      totalDebit += absAmount;
    }
  }
  
  var neracaData = [];
  neracaData.push(["Nama Akun", "Debit (Pemasukan)", "Kredit (Pengeluaran)"]);
  
  neracaData.push(["PEMASUKAN", "", ""]);
  neracaData.push(["Saldo Awal", saldoAwal, ""]);
  
  var totalPemasukan = saldoAwal;
  for (var u in units) {
    if (units[u].debit > 0) {
      var prefix = (u === 'Tanpa Unit') ? "" : "Pemasukan / BODP ";
      neracaData.push([prefix + u, units[u].debit, ""]);
      totalPemasukan += units[u].debit;
    }
  }
  
  neracaData.push(["Total Pemasukan", totalPemasukan, ""]);
  neracaData.push(["", "", ""]); 
  
  neracaData.push(["PENGELUARAN", "", ""]);
  var totalPengeluaran = 0;
  for (var u in units) {
    if (units[u].kredit > 0) {
      var prefix = (u === 'Tanpa Unit') ? "" : "Operasional / BODP ";
      neracaData.push([prefix + u, "", units[u].kredit]);
      totalPengeluaran += units[u].kredit;
    }
  }
  
  neracaData.push(["Total Pengeluaran", "", totalPengeluaran]);
  neracaData.push(["", "", ""]); 
  
  var saldoAkhir = totalPemasukan - totalPengeluaran;
  neracaData.push(["SALDO AKHIR", saldoAkhir, ""]);
  
  return { data: neracaData, availableMonths: availableMonths };
}

function updateNeracaKeuangan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var jagoSheet = ss.getSheetByName('REKAPAN JAGO WEB');
  var neracaSheet = ss.getSheetByName('NERACA KEUANGAN WEB');
  
  if (!jagoSheet) return;
  
  if (!neracaSheet) {
    neracaSheet = ss.insertSheet('NERACA KEUANGAN WEB');
  }
  
  var jagoData = jagoSheet.getDataRange().getValues();
  if (jagoData.length < 2) return;
  
  var filterValue = neracaSheet.getRange("E2").getValue();
  var calcResult = calculateNeraca(jagoData, filterValue);
  
  var availableMonths = calcResult.availableMonths;
  var selectedFilter = filterValue;
  if (!selectedFilter || (selectedFilter !== "Semua Bulan" && availableMonths.indexOf(selectedFilter) === -1)) {
      selectedFilter = "Semua Bulan";
      neracaSheet.getRange("E2").setValue(selectedFilter);
  }
  
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["Semua Bulan"].concat(availableMonths)).build();
  neracaSheet.getRange("E1").setValue("Filter Bulan:").setFontWeight("bold");
  neracaSheet.getRange("E2").setDataValidation(rule).setBackground("#fef08a");
  
  neracaSheet.getRange("A:D").clear(); 
  
  var neracaData = calcResult.data;
  if (neracaData.length > 0 && neracaData[0][0].toString().indexOf("Tidak ada data") !== -1) {
    neracaSheet.getRange("A1").setValue(neracaData[0][0]);
    return;
  }
  
  neracaSheet.getRange(1, 1, neracaData.length, 3).setValues(neracaData);
  
  // Styling
  neracaSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f4f6");
  
  for (var r = 0; r < neracaData.length; r++) {
    var rowName = neracaData[r][0];
    if (rowName === "PEMASUKAN" || rowName === "PENGELUARAN") {
      neracaSheet.getRange(r + 1, 1, 1, 3).setFontWeight("bold").setBackground("#e5e7eb");
    }
    if (rowName === "Total Pemasukan" || rowName === "Total Pengeluaran" || rowName === "SALDO AKHIR") {
      neracaSheet.getRange(r + 1, 1, 1, 3).setFontWeight("bold").setBackground("#d1d5db");
    }
  }
  
  neracaSheet.getRange(2, 2, neracaData.length - 1, 2).setNumberFormat('"Rp" #,##0');
  neracaSheet.autoResizeColumns(1, 3);
}

// --- PLANNING SALDO FUNCTIONS ---
function getOrCreatePlanningSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PLANNING");
  if (!sheet) {
    sheet = ss.insertSheet("PLANNING");
    sheet.appendRow(["ID", "Timestamp", "Planning", "Kategori", "Budget"]);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f4f6");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getPlanningData() {
  var sheet = getOrCreatePlanningSheet();
  var data = sheet.getDataRange().getValues();
  var plans = [];
  if (data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // Check if ID exists
        plans.push({
          id: data[i][0],
          timestamp: data[i][1],
          planning: data[i][2],
          kategori: data[i][3],
          budget: data[i][4],
          row: i + 1
        });
      }
    }
  }
  return plans;
}
