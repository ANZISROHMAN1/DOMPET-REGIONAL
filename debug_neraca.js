const code = require('fs').readFileSync('Code.gs', 'utf8');

const ContentService = {
  createTextOutput: (str) => ({
    setMimeType: () => str
  }),
  MimeType: { JSON: 'JSON' }
};

let outputLog = "";

const SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: (name) => {
      if (name === 'REKAPAN JAGO WEB') {
        return {
          getDataRange: () => ({
            getValues: () => {
              // I can't read the actual google sheet from local Node.js.
              // We need to run this on Apps Script.
              return [];
            }
          })
        };
      }
      return null;
    }
  })
};

