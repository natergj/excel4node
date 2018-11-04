const builtInFormatIdLookup: Map<string, number> = new Map([
  ['General 10', 0],
  ['0.00', 2],
  ['#, ##0', 3],
  ['#, ##0.00', 4],
  ['0 %', 9],
  ['0.00 %', 10],
  ['0.00E+00', 11],
  ['# ?/?', 12],
  ['# ??/??', 13],
  ['mm - dd - yy', 14],
  ['d - mmm - yy', 15],
  ['d - mmm', 16],
  ['mmm - yy', 17],
  ['h: mm AM / PM', 18],
  ['h: mm: ss AM / PM', 19],
  ['h: mm', 20],
  ['h: mm: ss', 21],
  ['m / d / yy h: mm', 22],
  ['#, ##0; (#, ##0)', 37],
  ['#, ##0;[Red](#, ##0)', 38],
  ['#, ##0.00; (#, ##0.00)', 39],
  ['#, ##0.00;[Red](#, ##0.00)', 40],
  ['mm: ss', 45],
  ['[h]: mm: ss', 46],
  ['mmss.0', 47],
  ['##0.0E+0', 48],
  ['@', 49],
]);

export class NumberFormat {
  _formatCode: string;
  _numFmtId: number;

  constructor(fmt: string) {
    this._formatCode = fmt;
    if (builtInFormatIdLookup.get(fmt)) {
      this._numFmtId = builtInFormatIdLookup.get(fmt);
    }
  }

  get hash(): string {
    // Since the number format is often shorter than an MD5 hash, we will use that as the lookup hash
    return this._formatCode;
  }

  get numFmtId(): number {
    return this._numFmtId;
  }

  set numFmtId(id) {
    this._numFmtId = id;
  }

  get formatCode() {
    return this._formatCode;
  }

  /**
   * @alias NumberFormat.addToXMLele
   * @desc When generating Workbook output, attaches style to the styles xml file
   * @func NumberFormat.addToXMLele
   * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
   */
  addToXMLele(ele) {
    if (this.formatCode !== undefined) {
      ele
        .ele('numFmt')
        .att('formatCode', this.formatCode)
        .att('numFmtId', this.numFmtId);
    }
  }
}
