var xl = require('./lib');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('Test Worksheet');

var myCell = ws.Cell(1, 1).String('One: !!');
var myCell = ws.Cell(2, 1).String('Two: ??');
var myCell = ws.Cell(3, 1).String('Three: ?');

wb.write('./dev.xlsx');

console.log('ok');
