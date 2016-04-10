var lodash = require('lodash');
var test = require('tape');

var CfRule = require('../distribution/lib/worksheet/cf/cf_rule');

test('CfRule init', function (t) {
    t.plan(4);

    var baseConfig = {
        type: 'expression',
        formula: 'NOT(ISERROR(SEARCH("??", A1)))',
        priority: 1,
        dxfId: 0
    };

    t.ok(new CfRule(baseConfig), 'init with valid and support type');

    try {
        var cfr = new CfRule(lodash.extend(baseConfig, { type: 'bogusType' }));
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with invalid type should throw an error'
        );
    }

    try {
        var cfr = new CfRule(lodash.extend(baseConfig, { type: 'dataBar' }));
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with an unsupported type should throw an error'
        );
    }

    try {
        var cfr = new CfRule(lodash.extend(baseConfig, { forumla: null }));
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with missing properties should throw an error'
        );
    }

});
