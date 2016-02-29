const _ = require('lodash');
const CF_RULE_TYPES = require('./cf_rule_types');

class CfRule {
    constructor(ruleConfig) {
        this.config = ruleConfig;

        let foundType = CF_RULE_TYPES[this.config.type];

        if (!foundType) {
            throw new TypeError('"' + this.config.type + '" is not a valid conditional formatting rule type');
        }

        if (!foundType.supported) {
            throw new TypeError('Conditional formatting type "' + this.config.type + '" is not yet supported');
        }

        let missingProps = _.reduce(foundType.requiredProps, function (list, prop) {
            if (_.get(this.config, prop, null) === null) {
                list.push(prop);
            }
            return list;
        }, []);

        if (missingProps.length) {
            throw new TypeError('Conditional formatting rule is missing required properties: ' + missingProps.join(', '));
        }
    }

    getBuilderData() {
        return {
            'cfRule': {
                '@type': this.config.type,
                '@dxfId': this.config.dxfId,
                '@priority': this.config.priority,
                'formula': this.config.formula
            }
        };
    }
}


module.exports = CfRule;