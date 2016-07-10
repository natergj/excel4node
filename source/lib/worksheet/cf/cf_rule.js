const _ = require('lodash');
const CF_RULE_TYPES = require('./cf_rule_types');

class CfRule { // §18.3.1.10 cfRule (Conditional Formatting Rule)
    constructor(ruleConfig) {
        this.type = ruleConfig.type;
        this.priority = ruleConfig.priority;
        this.formula = ruleConfig.formula;
        this.dxfId = ruleConfig.dxfId;

        let foundType = CF_RULE_TYPES[this.type];

        if (!foundType) {
            throw new TypeError('"' + this.type + '" is not a valid conditional formatting rule type');
        }

        if (!foundType.supported) {
            throw new TypeError('Conditional formatting type "' + this.type + '" is not yet supported');
        }

        let missingProps = _.reduce(foundType.requiredProps, (list, prop) => {
            if (_.get(this, prop, null) === null) {
                list.push(prop);
            }
            return list;
        }, []);

        if (missingProps.length) {
            throw new TypeError('Conditional formatting rule is missing required properties: ' + missingProps.join(', '));
        }
    }

    addToXMLele(ele) {
        let thisRule = ele.ele('cfRule');
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
        }
        thisRule.up();
    }
}


module.exports = CfRule;
