// FUNGSI INI WAJIB DI-RUN SATU KALI UNTUK MEMAKSA IZIN PENUH
function paksaIzin() {
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getRootFolder().getFiles();
  DriveApp.createFile("test", "test", MimeType.PLAIN_TEXT).setTrashed(true);
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // ID FOLDER SUDAH OTOMATIS SAYA MASUKKAN DI SINI! (TIDAK PERLU DIUBAH LAGI)
    var FOLDER_ID = '1KZUfE8WprVIYasopAPFZ71SXHnC1MuAs'; 
    var folder = DriveApp.getFolderById(FOLDER_ID);
    
    if (action === 'submit') {
      var fileUrl = '';
      if (requestData.fileData && requestData.fileName && requestData.mimeType) {
        var base64Data = requestData.fileData.split(',')[1] || requestData.fileData;
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), requestData.mimeType, requestData.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      }
      
      var newId = new Date().getTime().toString().slice(-6);
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      
      sheet.appendRow([newId, now, requestData.nama, requestData.kegiatan, requestData.nominal, requestData.bank, requestData.rekening, 'Pending', fileUrl, '']);
      
      return ContentService.createTextOutput(JSON.stringify({success: true, id: newId})).setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === 'update_status') {
      var tfUrl = '';
      if (requestData.fileData && requestData.fileName && requestData.mimeType) {
        var base64Data = requestData.fileData.split(',')[1] || requestData.fileData;
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), requestData.mimeType, "TF_" + requestData.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        tfUrl = file.getUrl();
      }
      
      var rowIndex = parseInt(requestData.row_index);
      sheet.getRange(rowIndex, 8).setValue(requestData.status);
      if (tfUrl) sheet.getRange(rowIndex, 10).setValue(tfUrl);
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
