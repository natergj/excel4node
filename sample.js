var xl = require('./source/dev.js');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('First Sheet');
ws.Cell(1, 1).String('Cell 1A');
ws.Cell(1, 2).Number(200);
ws.Cell(1, 3).Bool(true);
ws.Cell(2, 1).Number(1);
ws.Cell(2, 2).Number(2);
ws.Cell(2, 3).Formula('SUM(A2:B2)');

wb.write('Excel.xlsx');
