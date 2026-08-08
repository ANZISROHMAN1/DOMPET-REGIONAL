function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var action = e.parameter.action;
  
  // Convert sheet data to JSON objects
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    row['row_index'] = i + 1; // 1-based index in sheets
    rows.push(row);
  }
  
  if (action === 'status') {
    var reqId = e.parameter.id;
    var result = rows.find(function(r) { return r['ID'].toString() === reqId.toString(); });
    
    if (result) {
      return ContentService.createTextOutput(JSON.stringify({success: true, data: result}))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: 'ID tidak ditemukan'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } else if (action === 'list') {
    // For admin
    // Sort descending by row_index (newest first)
    rows.reverse();
    return ContentService.createTextOutput(JSON.stringify({success: true, data: rows}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Action not found'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // CORS Headers
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  // Handle CORS preflight options
  if (typeof e !== 'undefined' && e.postData === undefined) {
    var output = ContentService.createTextOutput('');
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
  
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // REPLACE WITH YOUR FOLDER ID
    var FOLDER_ID = 'GANTI_DENGAN_ID_FOLDER_DRIVE_ANDA'; 
    var folder = DriveApp.getFolderById(FOLDER_ID);
    
    if (action === 'submit') {
      // 1. Save File to Drive
      var fileUrl = '';
      if (requestData.fileData && requestData.fileName && requestData.mimeType) {
        var base64Data = requestData.fileData.split(',')[1] || requestData.fileData; // Remove data:image/png;base64,
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), requestData.mimeType, requestData.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      }
      
      // 2. Save Data to Sheet
      var newId = new Date().getTime().toString().slice(-6); // Simple random 6 digit ID
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      
      // Columns: ID, Tanggal, Nama, Kegiatan, Nominal, Bank, Rekening, Status, Bukti_Path, Bukti_TF_Path
      sheet.appendRow([
        newId,
        now,
        requestData.nama,
        requestData.kegiatan,
        requestData.nominal,
        requestData.bank,
        requestData.rekening,
        'Pending',
        fileUrl,
        ''
      ]);
      
      return ContentService.createTextOutput(JSON.stringify({success: true, id: newId}))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === 'update_status') {
      var tfUrl = '';
      if (requestData.fileData && requestData.fileName && requestData.mimeType) {
        var base64Data = requestData.fileData.split(',')[1] || requestData.fileData;
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), requestData.mimeType, "TF_" + requestData.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        tfUrl = file.getUrl();
      }
      
      // Find row and update
      var rowIndex = parseInt(requestData.row_index);
      
      // Column H (8) is Status, Column J (10) is Bukti_TF_Path
      sheet.getRange(rowIndex, 8).setValue(requestData.status);
      if (tfUrl) {
        sheet.getRange(rowIndex, 10).setValue(tfUrl);
      }
      
      return ContentService.createTextOutput(JSON.stringify({success: true}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// SETUP FUNCTION (Run this once to add headers to a new sheet)
function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Tanggal', 'Nama', 'Kegiatan', 'Nominal', 'Bank', 'Rekening', 'Status', 'Bukti_Path', 'Bukti_TF_Path']);
  }
}
