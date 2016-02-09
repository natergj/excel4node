var CfRule = require('./cf_rule');

module.exports = CfRulesCollection;

// -----------------------------------------------------------------------------

function CfRulesCollection() {
    // rules are indexed by cell refs
    this.rulesBySqref = {};
    return this;
}

CfRulesCollection.prototype.add = function (sqref, ruleConfig) {
    var rules = this.rulesBySqref[sqref] || [];
    var newRule = new CfRule(ruleConfig);
    rules.push(newRule);
    this.rulesBySqref[sqref] = rules;
    return this;
};


CfRulesCollection.prototype.getBuilderElements = function () {
    var list = [];
    var self = this;
    if (!Object.keys(this.rulesBySqref).length) {
        return list;
    }
    Object.keys(this.rulesBySqref).forEach(function (sqref) {
        var rules = self.rulesBySqref[sqref];
        list.push({
            conditionalFormatting: {
                '@sqref': sqref,
                '#list': rules.map(function (rule, idx) {
                    return rule.getBuilderData();
                })
            }
        });
    });
    return list;
};
