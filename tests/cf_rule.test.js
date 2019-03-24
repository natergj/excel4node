var deepmerge = require('deepmerge');
var test = require('tape');

var CfRule = require('../source/lib/worksheet/cf/cf_rule');

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
        var cfr = new CfRule(deepmerge(baseConfig, {
            type: 'bogusType'
        }));
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with invalid type should throw an error'
        );
    }

    try {
        var cfr = new CfRule(deepmerge(baseConfig, {
            type: 'dataBar'
        }));
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with an unsupported type should throw an error'
        );
    }

    try {
        var cfr = new CfRule(deepmerge(baseConfig, {
            formula: null
        }));
    } catch (err) {
        t.ok(
            err instanceof TypeError,
            'init of CfRule with missing properties should throw an error'
        );
    }

});