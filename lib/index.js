var style = require('./style');

module.exports = {
    WorkBook: require('./workbook'),
    Style: style.Style
};

// TODO do not modify core prototypes 

Number.prototype.toExcelAlpha = function (isCaps) {
    isCaps = isCaps ? isCaps : true;
    var remaining = this;
    var aCharCode = isCaps ? 65 : 97;
    var columnName = '';
    while (remaining > 0) {
        var mod = (remaining - 1) % 26;
        columnName = String.fromCharCode(aCharCode + mod) + columnName;
        remaining = (remaining - 1 - mod) / 26;
    } 
    return columnName;
};

String.prototype.toExcelRowCol = function () {
    var numeric = this.split(/\D/).filter(function (el) {
        return el !== '';
    })[0];
    var alpha = this.split(/\d/).filter(function (el) {
        return el !== '';
    })[0];
    var row = parseInt(numeric, 10);
    var col = alpha.toUpperCase().split('').reduce(function (a, b, index, arr) {
        return a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1);
    }, 0);
    return { row: row, col: col };
};

Date.prototype.getExcelTS = function () {
    var epoch = new Date(1899, 11, 31);
    var dt = this.setDate(this.getDate() + 1);
    var ts = (dt-epoch) / (1000 * 60 * 60 * 24);
    return ts;
};
