const utils = require('../utils.js');
const _ = require('lodash');

class Column {
    constructor(col) {
        this.bestFit = null;
        this.collapsed = null;
        this.customWidth = null;
        this.hidden = null;
        this.max = null;
        this.min = null;
        this.outlineLevel = null;
        this.style = null;
        this.width = null;
    }
}

module.exports = Column;