var xl = require('./source/dev.js');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('My First Sheet');
console.log(ws);
console.log(ws.toXML());

wb.write('Excel.xlsx');
