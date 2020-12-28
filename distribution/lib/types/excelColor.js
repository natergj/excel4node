'use strict';

function items() {
    var _this = this;

    // subset of §20.1.10.48 ST_PresetColorVal (Preset Color Value)
    this['aqua'] = 'FF33CCCC';
    this['black'] = 'FF000000';
    this['blue'] = 'FF0000FF';
    this['blue-gray'] = 'FF666699';
    this['bright green'] = 'FF00FF00';
    this['brown'] = 'FF993300';
    this['dark blue'] = 'FF000080';
    this['dark green'] = 'FF003300';
    this['dark red'] = 'FF800000';
    this['dark teal'] = 'FF003366';
    this['dark yellow'] = 'FF808000';
    this['gold'] = 'FFFFCC00';
    this['gray-25'] = 'FFC0C0C0';
    this['gray-40'] = 'FF969696';
    this['gray-50'] = 'FF808080';
    this['gray-80'] = 'FF333333';
    this['green'] = 'FF008000';
    this['indigo'] = 'FF333399';
    this['lavender'] = 'FFCC99FF';
    this['light blue'] = 'FF3366FF';
    this['light green'] = 'FFCCFFCC';
    this['light orange'] = 'FFFF9900';
    this['light turquoise'] = 'FFCCFFFF';
    this['light yellow'] = 'FFFFFF99';
    this['lime'] = 'FF99CC00';
    this['olive green'] = 'FF333300';
    this['orange'] = 'FFFF6600';
    this['pale blue'] = 'FF99CCFF';
    this['pink'] = 'FFFF00FF';
    this['plum'] = 'FF993366';
    this['red'] = 'FFFF0000';
    this['rose'] = 'FFFF99CC';
    this['sea green'] = 'FF339966';
    this['sky blue'] = 'FF00CCFF';
    this['tan'] = 'FFFFCC99';
    this['teal'] = 'FF008080';
    this['turquoise'] = 'FF00FFFF';
    this['violet'] = 'FF800080';
    this['white'] = 'FFFFFFFF';
    this['yellow'] = 'FFFFFF00';

    this.opts = [];
    Object.keys(this).forEach(function (k) {
        if (typeof _this[k] === 'string') {
            _this.opts.push(k);
        }
    });
}

items.prototype.validate = function (val) {
    if (this[val.toLowerCase()] === undefined) {
        var opts = [];
        for (var name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for ST_PresetColorVal; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

items.prototype.getColor = function (val) {
    // check for RGB, RGBA or Excel Color Names and return RGBA

    if (typeof this[val.toLowerCase()] === 'string') {
        // val was a named color that matches predefined list. return corresponding color
        return this[val.toLowerCase()];
    } else if (val.length === 8 && /^[a-fA-F0-9()]+$/.test(val)) {
        // val is already a properly formatted color string, return upper case version of itself
        return val.toUpperCase();
    } else if (val.length === 6 && /^[a-fA-F0-9()]+$/.test(val)) {
        // val is color code without Alpha, add it and return
        return 'FF' + val.toUpperCase();
    } else if (val.length === 7 && val.substr(0, 1) === '#' && /^[a-fA-F0-9()]+$/.test(val.substr(1))) {
        // val was sent as html style hex code, remove # and add alpha
        return 'FF' + val.substr(1).toUpperCase();
    } else {
        // I don't know what this is, return valid color and console.log error
        throw new TypeError('valid color options are html style hex codes, ARGB strings or these colors by name: %s', this.opts.join(', '));
    }
};

module.exports = new items();
//# sourceMappingURL=excelColor.js.map