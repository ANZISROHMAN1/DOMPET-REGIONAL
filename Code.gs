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
            return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Gagal: Nota ini terdeteksi sama persis dengan nota yang sudah pernah diklaim sebelumnya!'})).setMimeType(ContentService.MimeType.JSON);
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
        
        var jagoDataToInsert = [now, sourceDest, transDetails, requestData.kegiatan, amount];
        
        // Mencari baris kosong pertama di kolom A (Menghindari bug appendRow jika ada ArrayFormula)
        var jagoColA = sheetJago.getRange("A:A").getValues();
        var jagoTargetRow = jagoColA.length + 1;
        for (var i = 0; i < jagoColA.length; i++) {
          if (jagoColA[i][0] === "" && i > 0) { // i > 0 untuk melewati header
            jagoTargetRow = i + 1;
            break;
          }
        }
        
        sheetJago.getRange(jagoTargetRow, 1, 1, 5).setValues([jagoDataToInsert]);
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
