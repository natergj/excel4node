const CfRule = require('./cf_rule');

// -----------------------------------------------------------------------------

class CfRulesCollection { // §18.3.1.18 conditionalFormatting (Conditional Formatting)
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

    addToXMLele(ele) {
        Object.keys(this.rulesBySqref).forEach((sqref) => {
            let thisEle = ele.ele('conditionalFormatting').att('sqref', sqref);
            this.rulesBySqref[sqref].forEach((rule) => {
                rule.addToXMLele(thisEle);
            });
            thisEle.up();
        });
    }
}


module.exports = CfRulesCollection;