var notesAndMoney = "Transfer to Joko - 100.000 500.000";
var tokens = notesAndMoney.split(/\s+/);
var balanceStr = tokens.pop();
var amountStr = tokens.pop();
var notes = tokens.join(" ");
console.log("balanceStr:", balanceStr);
console.log("amountStr:", amountStr);
console.log("notes:", notes);
