var xl = require('./lib');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('Test Worksheet');

var myCell = ws.Cell(1, 1).String('One: !!');
var myCell = ws.Cell(2, 1).String('Two: ??');
var myCell = ws.Cell(3, 1).String('Three: ?');
var myCell = ws.Cell(4, 1).String('Four: !!');
var myCell = ws.Cell(5, 1).String('Five: ?');
var myCell = ws.Cell(6, 1).String('Six: !');

ws.addConditionalFormattingRule('A1:A10', {
    type: 'containsText',
    priority: 1,
    operator: 'containsText',
    text: '!!',
    formula: 'NOT(ISERROR(SEARCH("!!", A1)))'
});

wb.write('./dev.xlsx');

console.log('ok');
