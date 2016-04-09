const CTMarker = require('../classes/ctMarker.js');
const Point = require('../classes/point.js');
const EMU = require('../classes/emu.js');

class Drawing {
    constructor() {
        this._anchorType = null;
        this._anchorFrom = null;
        this._anchorTo = null;
        this._editAs = null;
        this._position = null;
    }

    get anchorType() {
        return this._anchorType;
    }
    set anchorType(type) {
        let types = ['absoluteAnchor', 'oneCellAnchor', 'twoCellAnchor'];
        if (types.indexOf(type) < 0) {
            throw new TypeError('Invalid option for anchor type. anchorType must be one of ' + types.join(', '));
        }
        this._anchorType = type;
    }

    get editAs() {
        return this._editAs;
    }
    set editAs(val) {
        let types = ['absolute', 'oneCell', 'twoCell'];
        if (types.indexOf(val) < 0) {
            throw new TypeError('Invalid option for editAs. editAs must be one of ' + types.join(', '));
        }
        this._editAs = val;
    }

    get anchorFrom() {
        return this._anchorFrom;
    }
    set anchorFrom(obj) {
        if (obj !== undefined && obj instanceof Object) {
            this._anchorFrom = new CTMarker(obj.col - 1, obj.colOff, obj.row - 1, obj.rowOff);
        }
    }

    get anchorTo() {
        return this._anchorTo;
    }
    set anchorTo(obj) {
        if (obj !== undefined && obj instanceof Object) {
            this._anchorTo = new CTMarker(obj.col - 1, obj.colOff, obj.row - 1, obj.rowOff);
        }
    }

    anchor(type, from, to) {
        if (type === 'twoCellAnchor') {
            if (from === undefined || to === undefined) {
                throw new TypeError('twoCellAnchor requires both from and two markers');
            }
            this.editAs = 'oneCell';
        }
        this.anchorType = type;
        this.anchorFrom = from;
        this.anchorTo = to;
        return this;
    }

    position(cx, cy) {
        this.anchorType = 'absoluteAnchor';
        let thisCx = new EMU(cx);
        let thisCy = new EMU(cy);
        this._position = new Point(thisCx.value, thisCy.value);
    }
}

module.exports = Drawing;