class NumberFormat {
    /**
    * @class NumberFormat
    * @param {String} fmt Format of the Number
    * @returns {NumberFormat}
    */
    constructor(fmt) {
        this.formatCode = fmt;
        this.id;
    }

    get numFmtId() {
        return this.id;
    }
    set numFmtId(id) {
        this.id = id;
    }

    /**
     * @alias NumberFormat.addToXMLele
     * @desc When generating Workbook output, attaches style to the styles xml file
     * @func NumberFormat.addToXMLele
     * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
     */
    addToXMLele(ele) {
        if (this.formatCode !== undefined) {
            ele.ele('numFmt')
            .att('formatCode', this.formatCode)
            .att('numFmtId', this.numFmtId);
        }
    }
}

module.exports = NumberFormat;