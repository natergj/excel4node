module.exports = CfRule;

var CF_RULE_TYPES = require('./cf_rule_types');

function CfRule(ruleConfig) {
    this.config = ruleConfig;
    var foundType = CF_RULE_TYPES[this.config.type];
    if (!foundType) {
        throw new TypeError('"' + this.config.type + '" is not a valid conditional formatting rule type');
    }
    if (!foundType.supported) {
        throw new TypeError('Conditional formatting type "' + this.config.type + '" is not yet supported');
    }
    return this;
}

CfRule.prototype.getBuilderData = function () {
    return {
        cfRule: {
            '@type': this.config.type,
            '@dxfId': this.config.dxfId,
            '@priority': this.config.priority,
            // '@operator': this.config.operator,
            // '@text': this.config.text,
            formula: this.config.formula
        }
    };
};
