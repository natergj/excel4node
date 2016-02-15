var xl = require('./source/dev.js');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('First Sheet');
ws.Cell(1,1).String('Cell 1A');

wb.write('Excel.xlsx');
