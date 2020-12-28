'use strict';

//§22.9.2.12 ST_PositiveUniversalMeasure (Positive Universal Measurement)

function measure() {}

measure.prototype.validate = function (val) {
    var re = new RegExp('[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)');
    if (re.test(val) !== true) {
        throw new TypeError('Invalid value for universal positive measure. Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. 10.5cm');
    } else {
        return true;
    }
};

module.exports = new measure();
//# sourceMappingURL=positiveUniversalMeasure.js.map