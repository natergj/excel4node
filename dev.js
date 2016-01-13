var xl = require('./lib');

var wb = new xl.WorkBook();
var ws = wb.WorkSheet('Test Worksheet');

var myCell = ws.Cell(1, 1).String('One: !!');
var myCell = ws.Cell(2, 1).String('Two: ??');
var myCell = ws.Cell(3, 1).String('Three: ?');
var myCell = ws.Cell(4, 1).String('Four: !!');
var myCell = ws.Cell(5, 1).String('Five: ?');
var myCell = ws.Cell(6, 1).String('Six: !');

var style = wb.Style();
style.Font.Bold();
style.Font.Italics();
style.Font.Underline();
style.Font.Color('000055');
style.Fill.Color('DDEEFF');
style.Fill.Pattern('solid');
style.Border({
    top: {
        style: 'thin',
        color: 'CCFFCC'
    },
    bottom: {
        style: 'thick'
    },
    left: {
        style: 'thin'
    },
    right: {
        style: 'thin'
    }
});

ws.addConditionalFormattingRule('A1:A10', {
    type: 'expression',
    priority: 1,
    formula: 'NOT(ISERROR(SEARCH("!!", A1)))',
    style: style
});

wb.write('./dev.xlsx');

console.log('ok');
