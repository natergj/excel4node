//require('babel-register');
//var xl = require('./source');

require('source-map-support').install();
var xl = require('./distribution');


var wb = new xl.WorkBook({
    defaultFont: {
        name: 'Arial'
    },
    logLevel: 5,
    dateFormat: 'mm-dd-yyyy hh:MM'
});
var ws = wb.addWorksheet('Sheet 1');
/*
ws.cell(1, 1).string('My simple string').style({ alignment: { wrapText: true } });
ws.cell(1, 2).number(5);
ws.cell(1, 3).formula('B1 * 10');
ws.cell(1, 4).date(new Date());
ws.column(4).setWidth(15);
ws.cell(1, 5).link('http://iamnater.com');
ws.column(5).width = 15;
ws.cell(1, 6).link('http://iamnater.com', 'Website Link', 'Link to my website');
ws.column(6).width = 15;
ws.cell(1, 7).bool(true);

ws.cell(2, 1, 2, 6, true).string('One big merged cell');
ws.cell(3, 1, 3, 6).number(1); // All 6 cells set to number 1


var complexString = [
    'Workbook default font String\n',
    {
        bold: true,
        underline: true,
        italic: true,
        color: 'FF0000',
        size: 18,
        name: 'Courier',
        value: 'Hello'
    },
    ' World!',
    {
        color: '000000',
        underline: false,
        name: 'Arial',
        vertAlign: 'subscript'
    },
    ' All',
    ' these',
    ' strings',
    ' are',
    ' black subsript,',
    {
        color: '0000FF',
        value: '\nbut',
        vertAlign: 'baseline'
    },
    ' now are blue'
];

ws.cell(4, 1).style({ alignment: { wrapText: true } }).string(complexString);
ws.row(4).setHeight(100);
ws.column(1).setWidth(75);
*/

var myStyle = wb.createStyle({
    font: {
        bold: true,
        underline: true
    }, 
    alignment: {
        wrapText: true,
        horizontal: 'center'
    }
});

ws.cell(5, 1).string('my \n multiline\n string').style(myStyle);
ws.cell(6, 1).string('row 6 string');
ws.cell(7, 1).string('row 7 string');
ws.cell(6, 1, 7, 1).style(myStyle);
ws.cell(7, 1).style({ font: { underline: false } });

var ws2 = wb.addWorksheet('Sheet 2');
var myStyle2 = wb.createStyle({
    font: {
        bold: true,
        color: '00FF00'
    }
});
 
ws2.addConditionalFormattingRule('A1:A10', {      // apply ws formatting ref 'A1:A10' 
    type: 'expression',                          // the conditional formatting type 
    priority: 1,                                 // rule priority order (required) 
    formula: 'NOT(ISERROR(SEARCH("ok", A1)))',   // formula that returns nonzero or 0 
    style: myStyle2                               // a style object containing styles to apply 
});


ws2.addImage({
    path: './screenshot.png',
    position: {
        type: 'oneCellAnchor',
        from: {
            col: 1,
            colOff: '0.5in',
            row: 1,
            rowOff: 0 
        }
    }
});


ws2.addImage({
    path: './screenshot2.png',
    position: {
        type: 'absoluteAnchor',
        x: '1in',
        y: '2in'
    }
});

wb.write('Excel.xlsx');