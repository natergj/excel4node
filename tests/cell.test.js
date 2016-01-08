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
