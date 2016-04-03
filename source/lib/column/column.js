const utils = require('../utils.js');
const _ = require('lodash');

class Column {
    constructor(col, ws) {
        this.ws = ws;
        this.collapsed = null;
        this.customWidth = null;
        this.hidden = null;
        this.max = col;
        this.min = col;
        this.outlineLevel = null;
        this.style = null;
        this.colWidth = null;
    }

    get width() {
        return this.colWidth;
    }

    set width(w) {
        if (parseInt(w) === w) {
            this.colWidth = w;
            this.customWidth = true;
        } else {
            throw new TypeError('Column width must be a positive integer');
        }
        return this.colWidth;
    }

    setWidth(w) {
        if (parseInt(w) === w) {
            this.colWidth = w;
            this.customWidth = true;
        } else {
            throw new TypeError('Column width must be a positive integer');
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

        if (collapsed === undefined) {
            return this;
        }

        if (typeof collapsed === 'boolean') {
            this.collapsed = collapsed;
            this.hidden = collapsed;
        } else {
            throw new TypeError('Column group collapse flag must be a boolean');
        }

        return this; 
    }

    freeze(jumpTo) {
        let o = this.ws.opts.sheetView.pane;
        jumpTo = typeof jumpTo === 'number' && jumpTo > this.min ? jumpTo : this.min + 1;
        o.state = 'frozen';
        o.xSplit = this.min;
        o.activePane = 'bottomRight';
        o.ySplit === null ? 
            o.topLeftCell = utils.getExcelCellRef(1, jumpTo) : 
            o.topLeftCell = utils.getExcelCellRef(utils.getExcelRowCol(o.topLeftCell).row, jumpTo);

        this.ws.wb.logger.debug(o);
        return this;
    }
}

module.exports = Column;