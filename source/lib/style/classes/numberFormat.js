class NumberFormat {
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

    addToXMLele(ele) {
        if (this.formatCode !== undefined) {
            ele.ele('numFmt')
            .att('formatCode', this.formatCode)
            .att('numFmtId', this.numFmtId);
        }
    }
}

module.exports = NumberFormat;