class EMU {
    constructor(val) {
        this._value;
        this.value = val;
    }

    get value() {
        return this._value;
    }

    set value(val) {
        if (val === undefined) {
            this._value = 0;
        } else if (typeof val === 'number') {
            this._value = val ? parseInt(val) : 0; 
        } else if (typeof val === 'string') {
            let re = new RegExp('[0-9]+(\.[0-9]+)?(mm|cm|in)');
            if (re.test(val) === true) {
                let measure = parseFloat(/[0-9]+(\.[0-9]+)?/.exec(val)[0]);
                let unit = /(mm|cm|in)/.exec(val)[0];

                switch (unit) {
                case 'mm':
                    this._value = parseInt(measure * 36000);
                    break;

                case 'cm':
                    this._value = parseInt(measure * 360000);
                    break;

                case 'in':
                    this._value = parseInt(measure * 914400);
                    break;
                }
            } else {
                throw new TypeError('EMUs must be specified as whole integer EMUs or Floats immediately followed by unit of measure in cm, mm, or in. i.e. "1.5in"');
            }
        }        
    }

    toInt() {
        return this._value;
    }

    toInch() {
        return this._value / 914400;
    }

    toCM() {
        return this._value / 360000;
    }
}

module.exports = EMU;

/*
M.4.1.1 EMU Unit of Measurement

1 emu  = 1/914400 in = 1/360000 cm

Throughout ECMA-376, the EMU is used as a unit of measurement for length. An EMU is defined as follows:
The EMU was created in order to be able to evenly divide in both English and Metric units, in order to 
avoid rounding errors during the calculation. The usage of EMUs also facilitates a more seamless system 
switch and interoperability between different locales utilizing different units of measurement. 
EMUs define an integer based, high precision coordinate system.
*/