const CTMarker = require('../classes/ctMarker.js');
const Point = require('../classes/point.js');
const EMU = require('../classes/emu.js');

class Drawing {
    /**
     * Element representing an Excel Drawing superclass
     * @property {String} anchorType Proprty for type of anchor. One of 'absoluteAnchor', 'oneCellAnchor', 'twoCellAnchor'
     * @property {CTMarker} anchorFrom Property for the top left corner position of drawing
     * @property {CTMarker} anchorTo Property for the bottom left corner position of drawing
     * @property {String} editAs Property that states how to interact with the Drawing in Excel. One of 'absolute', 'oneCell', 'twoCell'
     * @property {Point} _position Internal property for position on Excel Worksheet when drawing type is absoluteAnchor
     * @returns {Drawing} Excel Drawing 
     */
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

    /**
     * @alias Drawing.achor
     * @desc Sets the postion and anchor properties of the Drawing
     * @func Drawing.achor
     * @param {String} type Anchor type of drawing
     * @param {Object} from Properties for achorFrom property
     * @param {Number} from.col Left edge of drawing will align with left edge of this column
     * @param {String} from.colOff Offset. Drawing will be shifted to the right the specified amount. Float followed by measure [0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi). i.e '10.5mm'
     * @param {Number} from.row Top edge of drawing will align with top edge of this row
     * @param {String} from.rowOff Offset. Drawing will be shifted down the specified amount. Float followed by measure [0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi). i.e '10.5mm'
     * @param {Object} to Properties for anchorTo property
     * @param {Number} to.col Left edge of drawing will align with left edge of this column
     * @param {String} to.colOff Offset. Drawing will be shifted to the right the specified amount. Float followed by measure [0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi). i.e '10.5mm'
     * @param {Number} to.row Top edge of drawing will align with top edge of this row
     * @param {String} to.rowOff Offset. Drawing will be shifted down the specified amount. Float followed by measure [0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi). i.e '10.5mm'
     * @returns {Drawing} Excel Drawing with attached methods
     */
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

    /**
     * @alias Drawing.position
     * @desc The position of the top left corner of the image on the Worksheet
     * @func Drawing.position
     * @param {ST_PositiveUniversalMeasure} cx Postion from left of Worksheet edge
     * @param {ST_PositiveUniversalMeasure} cy Postion from top of Worksheet edge
     */
    position(cx, cy) {
        this.anchorType = 'absoluteAnchor';
        let thisCx = new EMU(cx);
        let thisCy = new EMU(cy);
        this._position = new Point(thisCx.value, thisCy.value);
    }
}

module.exports = Drawing;