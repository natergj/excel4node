var test = require('tape');

var xl = require('../lib/index');
var Column = require('../lib/Column').Column;

function makeColumn() {
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test');
    return ws.Column(1);
}

test('Column coverage', function (t) {
    t.plan(4);
    var col = makeColumn();
    t.ok(col.Width(100), 'Width()');
    t.ok(col.Freeze(), 'Freeze()');
    t.ok(col.Hide(), 'Hide()');
    t.ok(col.Group('foo', false), 'Group()');
});
