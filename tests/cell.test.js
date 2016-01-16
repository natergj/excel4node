var test = require('tape');

var xl = require('../lib/index');

function makeCell() {
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test');
    return ws.Cell(1, 1);
}

test('Cell coverage', function (t) {
    t.plan(1);
    var cell = makeCell();
    t.ok(cell);
});

test('Cell takes a Style object', function (t) {
    t.plan(1);
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test');
    var cell = ws.Cell(1, 1);
    var style = wb.Style();
    t.ok(cell.Style(style));
});

test('Cell test Complex String', function (t) {
    t.plan(1);
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test');
    var cell = ws.Cell(1, 1);
    var content = [{
            bold:true,
            underline: true,
            italic: true,
            color: 'FF0000',
            size: 18,
            value: 'Hello'
        }, ' World!', {color: '000000'},  ' All', ' these', ' strings', ' are', ' black', 
                   {
                      color: '0000FF',
                      value: ', but'
                   },  ' now', ' are', ' blue' ];

    cell.Complex(content);

    t.equal(JSON.stringify(wb.workbook.sharedStrings[0]), JSON.stringify(content), 'validating complex text');
});