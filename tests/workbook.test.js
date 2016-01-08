var test = require('tape');

// load prototype extensions
// TODO fix prototype extensions and remove this
require('../lib/index');

var WorkBook = require('../lib/WorkBook');

test('WorkBook init', function (t) {
    t.plan(1);
    var wb = new WorkBook();
    t.ok(wb);
});

// Initial test to cover lib at a high level
test('WorkBook coverage', function (t) {
    t.plan(1);

    var wb = new WorkBook();

    var ws = wb.WorkSheet('Test Worksheet');

    var myCell = ws.Cell(1, 1);
    myCell.String('Test Value');

    t.ok(
        Buffer.isBuffer(wb.writeToBuffer()),
        'WorkBook#writeToBuffer() returns a Buffer'
    );
});

