var lodash = require('lodash');

module.exports = CfRule;

var CF_RULE_TYPES = require('./cf_rule_types');

function CfRule(ruleConfig) {
    var self = this;
    this.config = ruleConfig;

    var foundType = CF_RULE_TYPES[this.config.type];

    if (!foundType) {
        throw new TypeError('"' + this.config.type + '" is not a valid conditional formatting rule type');
    }

    if (!foundType.supported) {
        throw new TypeError('Conditional formatting type "' + this.config.type + '" is not yet supported');
    }

    var missingProps = lodash.reduce(foundType.requiredProps, function (list, prop) {
        if (lodash.get(self.config, prop, null) === null) {
            list.push(prop);
        }
        return list;
    }, []);

    if (missingProps.length) {
        throw new TypeError('Conditional formatting rule is missing required properties: ' + missingProps.join(', '));
    }

    return this;
}

CfRule.prototype.getBuilderData = function () {
    return {
        cfRule: {
            '@type': this.config.type,
            '@dxfId': this.config.dxfId,
            '@priority': this.config.priority,
            formula: this.config.formula
        }
    };
};
