import * as utils from '../utils';

// §18.3.1.4 c (Cell)
enum CellDataType {
  boolean = 'b',
  date = 'd',
  error = 'e',
  inlineString = 'inlineStr',
  number = 'n',
  sharedString = 's',
  string = 'str',
}

// §18.3.1.4 c (Cell)
export class Cell {
  reference: string;
  styleIndex: number;
  type: CellDataType;
  formulaStr: string;
  row: number;
  column: number;
  value: string | number;
  inlineRichText: string;

  constructor(row: number, column: number) {
    this.reference = `${utils.getExcelAlpha(column)}${row}`;
    this.styleIndex = 0;
    this.type = null;
    this.formulaStr = null;
    this.value = null;
    this.row = row;
    this.column = column;
  }

  string(index) {
    this.type = CellDataType.sharedString;
    this.value = index;
    this.formulaStr = null;
  }

  number(val) {
    this.type = CellDataType.number;
    this.value = val;
    this.formulaStr = null;
  }

  formula(formula) {
    this.type = null;
    this.value = null;
    this.formulaStr = formula;
  }

  bool(val) {
    this.type = CellDataType.boolean;
    this.value = val;
    this.formulaStr = null;
  }

  date(dt) {
    this.type = null;
    this.value = utils.getExcelTS(dt);
    this.formulaStr = null;
  }

  style(sId) {
    this.styleIndex = sId;
  }

  addToXMLele(ele) {
    if (this.value === null && this.inlineRichText === null) {
      return;
    }

    const cEle = ele
      .ele('c')
      .att('r', this.reference)
      .att('s', this.styleIndex);
    if (this.type !== null) {
      cEle.att('t', this.type);
    }
    if (this.formulaStr !== null) {
      cEle
        .ele('f')
        .txt(this.formulaStr)
        .up();
    }
    if (this.value !== null) {
      cEle
        .ele('v')
        .txt(this.value)
        .up();
    }
    cEle.up();
  }
}
