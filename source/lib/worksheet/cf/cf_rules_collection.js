const CfRule = require('./cf_rule');

// -----------------------------------------------------------------------------

class CfRulesCollection {
    constructor() {
        // rules are indexed by cell refs
        this.rulesBySqref = {};
    }

    get count() {
        return Object.keys(this.rulesBySqref).length;
    }

    add(sqref, ruleConfig) {
        let rules = this.rulesBySqref[sqref] || [];
        let newRule = new CfRule(ruleConfig);
        rules.push(newRule);
        this.rulesBySqref[sqref] = rules;
        return this;
    }



    getBuilderElements() {
        let list = [];
        let self = this;
        if (!Object.keys(this.rulesBySqref).length) {
            return list;
        }
        Object.keys(this.rulesBySqref).forEach(function (sqref) {
            let rules = self.rulesBySqref[sqref];
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
    }
}


module.exports = CfRulesCollection;