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
        this.colWidth = null;
    }

    width(w) {
        if (w === 'auto') {
            this.bestFit = true;
            this.customWidth = true;
        } else if (parseInt(w) === w) {
            this.colWidth = w;
            this.customWidth = true;
        } else {
            throw new TypeError('Column width must be a positive integer or \'auto\'');
        }
        return this;
    }

    hide() {
        this.hidden = true;
        return this;
    }

    group(level, collapsed) {
        if (parseInt(level) === level) {
            this.outlineLevel = level;
        } else {
            throw new TypeError('Column group level must be a positive integer');
        }

        if (collapsed !== undefined && typeof collapsed === 'boolean') {
            this.collapsed = collapsed;
        } else {
            throw new TypeError('Column group collapse flag must be a boolean');
        }
        
    }
}

module.exports = Column;