var test = require('tape');

var WorkBook = require('../lib/WorkBook');

test('WorkBook init', function (t) {
    t.plan(1);
    var wb = new WorkBook();
    t.ok(wb);
});

