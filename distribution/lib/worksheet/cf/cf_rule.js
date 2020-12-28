'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _reduce = require('lodash.reduce');
var _get = require('lodash.get');
var CF_RULE_TYPES = require('./cf_rule_types');

var CfRule = function () {
    // §18.3.1.10 cfRule (Conditional Formatting Rule)
    function CfRule(ruleConfig) {
        var _this = this;

        _classCallCheck(this, CfRule);

        this.type = ruleConfig.type;
        this.priority = ruleConfig.priority;
        this.formula = ruleConfig.formula;
        this.dxfId = ruleConfig.dxfId;

        var foundType = CF_RULE_TYPES[this.type];

        if (!foundType) {
            throw new TypeError('"' + this.type + '" is not a valid conditional formatting rule type');
        }

        if (!foundType.supported) {
            throw new TypeError('Conditional formatting type "' + this.type + '" is not yet supported');
        }

        var missingProps = _reduce(foundType.requiredProps, function (list, prop) {
            if (_get(_this, prop, null) === null) {
                list.push(prop);
            }
            return list;
        }, []);

        if (missingProps.length) {
            throw new TypeError('Conditional formatting rule is missing required properties: ' + missingProps.join(', '));
        }
    }

    _createClass(CfRule, [{
        key: 'addToXMLele',
        value: function addToXMLele(ele) {
            var thisRule = ele.ele('cfRule');
            if (this.type !== undefined) {
                thisRule.att('type', this.type);
            }
            if (this.dxfId !== undefined) {
                thisRule.att('dxfId', this.dxfId);
            }
            if (this.priority !== undefined) {
                thisRule.att('priority', this.priority);
            }

            if (this.formula !== undefined) {
                thisRule.ele('formula').text(this.formula);
                thisRule.up();
            }
            thisRule.up();
        }
    }]);

    return CfRule;
}();

module.exports = CfRule;
//# sourceMappingURL=cf_rule.js.map