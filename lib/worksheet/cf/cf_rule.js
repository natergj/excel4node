module.exports = CfRule;

function CfRule(ruleConfig) {
    this.config = ruleConfig;
    return this;
}

CfRule.prototype.getBuilderData = function () {
    return {
        cfRule: {
            '@type': this.config.type,
            '@dxfId': 1, // TODO what is dxfId?
            '@priority': this.config.priority,
            '@operator': this.config.operator,
            '@text': this.config.text,
            formula: this.config.formula
        }
    };
};
