let EMU = require('./emu.js');

class CTMarker {
    constructor(colId, colOffset, rowId, rowOffset) {
        this._col = colId;
        this._colOff = new EMU(colOffset);
        this._row = rowId;
        this._rowOff = new EMU(rowOffset);
    }

    get col() {
        return this._col;
    }
    set col(val) {
        if (parseInt(val, 10) !== val || val < 0) {
            throw new TypeError('CTMarker column must be a positive integer');
        }
        this._col = val;
    }

    get row() {
        return this._row;
    }
    set row(val) {
        if (parseInt(val, 10) !== val || val < 0) {
            throw new TypeError('CTMarker row must be a positive integer');
        }
        this._row = val;
    }

    get colOff() {
        return this._colOff.value;
    }
    set colOff(val) {
        this._colOff = new EMU(val);
    }

    get rowOff() {
        return this._rowOff.value;
    }
    set rowOff(val) {
        this._rowOff = new EMU(val);
    }
}

module.exports = CTMarker;