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

console.log(parseSafeFloat("Rp 453.297.450")); // 453297450
console.log(parseSafeFloat("Rp453,297,450")); // 453297450
console.log(parseSafeFloat("123.456,78")); // 123456.78
console.log(parseSafeFloat("Rp 1000")); // 1000
console.log(parseSafeFloat(453297450)); // 453297450
console.log(parseSafeFloat("")); // 0
