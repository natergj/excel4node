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
