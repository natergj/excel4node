var test = require('tape');

var CfRule = require('../lib/worksheet/cf/cf_rule');

test('CfRule init', function (t) {
    t.plan(3);

    t.ok(new CfRule({ type: 'expression' }), 'init with valid and support type');

    try {
        var cfr = new CfRule({ type: 'bogusType' });
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with invalid type should throw an error'
        );
    }

    try {
        var cfr = new CfRule({ type: 'dataBar' });
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with an unsupported type should throw an error'
        );
    }

});
