//§18.18.60 ST_PrintError (Print Errors)
function items() {
    let opts = ['displayed', 'blank', 'dash', 'NA'];
    opts.forEach((o, i) => {
        this[o] = i + 1;
    });
}


items.prototype.validate = function (val) {
    if (this[val] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for pageSetup.errors; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();