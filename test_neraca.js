const fs = require('fs');
const code = fs.readFileSync('Code.gs', 'utf8');

const ContentService = {
  createTextOutput: (str) => ({
    setMimeType: () => str
  }),
  MimeType: { JSON: 'JSON' }
};

const SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: (name) => {
      if (name === 'REKAPAN JAGO WEB') {
        return {
          getDataRange: () => ({
            getValues: () => [
              ['Date & Time', 'Source/Destination', 'Transaction Details', 'Notes', 'Kas Masuk', 'Kas Keluar', 'Saldo', 'Unit', 'Kategori'],
              ['01 Sep 2026', 'A', 'B', 'C', '1000', '', '1000', 'Unit 1', 'Kategori 1'],
              ['02 Sep 2026', 'D', 'E', 'F', '', '500', '500', 'Unit 2', 'Kategori 2']
            ]
          })
        };
      }
      return null;
    },
    getActiveSheet: () => ({
      getDataRange: () => ({
        getValues: () => []
      })
    })
  })
};

const Session = {
  getScriptTimeZone: () => 'Asia/Jakarta'
};

const Utilities = {
  formatDate: (date, tz, format) => {
    if (format === 'MMMM yyyy') return 'September 2026';
    return date.toString();
  }
};

eval(code);

const e = { parameter: { action: 'neraca_data', month: 'Semua Bulan' } };
try {
  const result = doGet(e);
  console.log("SUCCESS:", result);
} catch (err) {
  console.log("ERROR:", err);
}
