'use strict';

// Types from xlsx spec:
//     http://download.microsoft.com/download/D/3/3/D334A189-E51B-47FF-B0E8-C0479AFB0E3C/[MS-XLSX].pdf

module.exports = {
    cellIs: {
        supported: false
    },
    expression: {
        supported: true,
        requiredProps: ['dxfId', 'priority', 'formula']
    },
    colorScale: {
        supported: false
    },
    dataBar: {
        supported: false
    },
    iconSet: {
        supported: false
    },
    containsText: {
        supported: false
    },
    notContainsText: {
        supported: false
    },
    beginsWith: {
        supported: false
    },
    endsWith: {
        supported: false
    },
    containsBlanks: {
        supported: false
    },
    notContainsBlanks: {
        supported: false
    },
    containsErrors: {
        supported: false
    },
    notContainsErrors: {
        supported: false
    }
};
//# sourceMappingURL=cf_rule_types.js.map