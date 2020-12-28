'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var CfRule = require('./cf_rule');

// -----------------------------------------------------------------------------

var CfRulesCollection = function () {
    // §18.3.1.18 conditionalFormatting (Conditional Formatting)
    function CfRulesCollection() {
        _classCallCheck(this, CfRulesCollection);

        // rules are indexed by cell refs
        this.rulesBySqref = {};
    }

    _createClass(CfRulesCollection, [{
        key: 'add',
        value: function add(sqref, ruleConfig) {
            var rules = this.rulesBySqref[sqref] || [];
            var newRule = new CfRule(ruleConfig);
            rules.push(newRule);
            this.rulesBySqref[sqref] = rules;
            return this;
        }
    }, {
        key: 'addToXMLele',
        value: function addToXMLele(ele) {
            var _this = this;

            Object.keys(this.rulesBySqref).forEach(function (sqref) {
                var thisEle = ele.ele('conditionalFormatting').att('sqref', sqref);
                _this.rulesBySqref[sqref].forEach(function (rule) {
                    rule.addToXMLele(thisEle);
                });
                thisEle.up();
            });
        }
    }, {
        key: 'count',
        get: function get() {
            return Object.keys(this.rulesBySqref).length;
        }
    }]);

    return CfRulesCollection;
}();

module.exports = CfRulesCollection;
//# sourceMappingURL=cf_rules_collection.js.map