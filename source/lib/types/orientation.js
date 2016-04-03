//§18.18.50 ST_Orientation (Orientation)

function items() {
    let opts = ['default', 'portrait', 'landscape'];
    opts.forEach((o, i) => {
        this[o] = i + 1;
    });
}


items.prototype.validate = function (val) {
    if (this[val.toLowerCase()] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for pageSetup.orientation; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();