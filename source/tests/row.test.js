var test = require('tape');

var xl = require('../lib/index');

function makeRow() {
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test');
    return ws.Row(1);
}

test('Row coverage', function (t) {
    t.plan(5);
    var row = makeRow();
    t.ok(row.Height(100), 'Width()');
    t.ok(row.Freeze(), 'Freeze()');
    t.ok(row.Hide(), 'Hide()');
    t.ok(row.Group(1, false), 'Group()');
    t.ok(row.Filter(1, 2, [
        {
            column: 2,
            rules: [
                {
                    val: 'food'
                },
                {
                    val: 'coffee'
                }
            ]
        }
    ]), 'Filter()');
});

