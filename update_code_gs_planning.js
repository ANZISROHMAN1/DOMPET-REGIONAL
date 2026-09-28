const fs = require('fs');
let code = fs.readFileSync('Code.gs', 'utf8');

const planningFunctions = `
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
`;

if (!code.includes('getOrCreatePlanningSheet')) {
    code += '\n' + planningFunctions + '\n';
}

// Add to doGet
const getTagihanListRegex = /(} else if \(action === 'tagihan_rutin_list'\) \{[\s\S]*?)(  \} else \{)/;
if (code.match(getTagihanListRegex) && !code.includes("action === 'get_planning'")) {
    const replacement = `$1  } else if (action === 'get_planning') {
    var plans = getPlanningData();
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: plans })).setMimeType(ContentService.MimeType.JSON);
$2`;
    code = code.replace(getTagihanListRegex, replacement);
}

// Add to doPost
const uploadBuktiRegex = /(} else if \(action === 'upload_bukti_tagihan'\) \{[\s\S]*?)(    \} else \{)/;
if (code.match(uploadBuktiRegex) && !code.includes("action === 'save_planning'")) {
    const replacement = `$1    } else if (action === 'save_planning') {
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
$2`;
    code = code.replace(uploadBuktiRegex, replacement);
}

fs.writeFileSync('Code.gs', code);
console.log('Code.gs updated successfully');
