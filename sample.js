var xl = require('./distribution/');

var wb = new xl.WorkBook({logLevel:3});
var ws = wb.WorkSheet('First Sheet', {
    'sheetProtection': {
        'autoFilter': false,
        'deleteColumns': null,
        'deleteRow': null,
        'formatCells': true,
        'formatColumns': true,
        'formatRows': true,
        'hashValue': null,
        'insertColumns': null,
        'insertHyperlinks': null,
        'insertRows': null,
        'objects': null,
        'password': 'PassWD',
        'pivotTables': null,
        'scenarios': null,
        'selectLockedCells': null,
        'selectUnlockedCell': null,
        'sheet': null,
        'sort': null
    }
});
ws.Cell(1, 1).String('Cell 1A');
ws.Cell(1, 2).Number(200);
ws.Cell(1, 3).Bool(true);
ws.Cell(2, 1).String('ok');
ws.Cell(2, 2).String('notOK');
ws.Cell(3, 1).String('2');
ws.Cell(3, 2).String('2');
ws.Cell(4, 1).String('3');
ws.Cell(4, 2).String('notOK');
ws.Cell(5, 1).String('5');
ws.Cell(5, 2).String('ok');
ws.Cell(7, 1, 7, 5, true).String('string');

ws.Row(1).Filter({ startColumn: 1, endColumn: 3 });

var style = wb.Style({
    font: {
        bold: true,
        color: '00FF00'
    }
});

var style2 = wb.Style({
    font: {
        italics: true,
        color: 'Green'
    }
});

/*
ws.addConditionalFormattingRule('A1:A10', {      // apply ws formatting ref 'A1:A10' 
    type: 'expression',                          // the conditional formatting type 
    priority: 1,                                 // rule priority order (required) 
    formula: 'NOT(ISERROR(SEARCH("ok", A1)))',   // formula that returns nonzero or 0 
    style: style                                 // a style object containing styles to apply 
});
ws.addConditionalFormattingRule('B1:B10', {      // apply ws formatting ref 'A1:A10' 
    type: 'expression',                          // the conditional formatting type 
    priority: 1,                                 // rule priority order (required) 
    formula: 'NOT(ISERROR(SEARCH("notOK", B1)))',   // formula that returns nonzero or 0 
    style: style                                 // a style object containing styles to apply 
});
*/

wb.write('Excel.xlsx');
