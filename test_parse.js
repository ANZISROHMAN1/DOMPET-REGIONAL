const fs = require('fs');
const code = fs.readFileSync('Code.gs', 'utf8');
console.log(code.match(/var dbSheet = ss\.getSheetByName\([^;]+;/g));
