import { Worksheet } from '../worksheet';
import { Cell } from './cell';
import { getExcelAlpha, sortCellRefs } from '../utils';
export * from './cell';

export class CellAccessor {
  ws: Worksheet;
  startRow: number;
  startCol: number;
  endRow?: number;
  endCol?: number;
  isMerged?: boolean;
  cells: Cell[];
  excelRefs: string[];

  constructor(ws: Worksheet, startRow: number, startCol: number, endRow?: number, endCol?: number, isMerged?: boolean) {
    this.ws = ws;
    this.startRow = startRow;
    this.startCol = startCol;
    this.endRow = endRow || startRow;
    this.endCol = endCol || startCol;
    this.isMerged = !!isMerged || false;
    this.cells = [];
    this.excelRefs = [];

    for (let r = this.startRow; r <= this.endRow; r++) {
      const row = this.ws.row(r);
      for (let c = this.startCol; c <= this.endCol; c++) {
        const ref = `${getExcelAlpha(c)}${r}`;
        row.cellRefs.add(ref);
        if (!this.ws.cells.get(ref)) {
          const newCell = new Cell(r, c);
          this.ws.cells.set(ref, newCell);
        }
        this.excelRefs.push(ref);
      }
    }
    this.excelRefs = this.excelRefs.sort(sortCellRefs);
    this.cells = this.excelRefs.map(r => this.ws.cells.get(r));
  }

  get matrix() {
    const matrix = [];
    const tmpObj = {};
    this.cells.forEach(c => {
      if (!tmpObj[c.row]) {
        tmpObj[c.row] = [];
      }
      tmpObj[c.row].push(c);
    });
    const rows = Object.keys(tmpObj);
    rows.forEach(r => {
      tmpObj[r].sort((a, b) => {
        return a.col - b.col;
      });
      matrix.push(tmpObj[r]);
    });
    return matrix;
  }

  get firstRow() {
    let firstRow;
    this.cells.forEach(c => {
      if (c.row < firstRow || firstRow === undefined) {
        firstRow = c.row;
      }
    });
    return firstRow;
  }

  get lastRow() {
    let lastRow;
    this.cells.forEach(c => {
      if (c.row > lastRow || lastRow === undefined) {
        lastRow = c.row;
      }
    });
    return lastRow;
  }

  get firstCol() {
    let firstCol;
    this.cells.forEach(c => {
      if (c.column < firstCol || firstCol === undefined) {
        firstCol = c.column;
      }
    });
    return firstCol;
  }

  get lastCol() {
    let lastCol;
    this.cells.forEach(c => {
      if (c.column > lastCol || lastCol === undefined) {
        lastCol = c.column;
      }
    });
    return lastCol;
  }

  string(str: string) {
    const { sharedStrings } = this.ws.wb;
    let index = sharedStrings.get(str);
    if (index === undefined) {
      index = sharedStrings.size;
      sharedStrings.set(str, index);
    }
    if (this.isMerged) {
      this.ws.cells.get(this.excelRefs[0]).string(index);
    } else {
      this.excelRefs.forEach(ref => this.ws.cells.get(ref).string(index));
    }
    return this;
  }

  number(num: number) {
    if (this.isMerged) {
      this.ws.cells.get(this.excelRefs[0]).number(num);
    } else {
      this.excelRefs.forEach(ref => this.ws.cells.get(ref).number(num));
    }
    return this;
  }

  formula(f: string) {
    if (this.isMerged) {
      this.ws.cells.get(this.excelRefs[0]).formula(f);
    } else {
      this.excelRefs.forEach(ref => this.ws.cells.get(ref).formula(f));
    }
    return this;
  }

  bool(b: boolean) {
    if (this.isMerged) {
      this.ws.cells.get(this.excelRefs[0]).bool(b);
    } else {
      this.excelRefs.forEach(ref => this.ws.cells.get(ref).bool(b));
    }
    return this;
  }

  date(dt: Date | string) {
    if (this.isMerged) {
      this.ws.cells.get(this.excelRefs[0]).date(dt);
    } else {
      this.excelRefs.forEach(ref => this.ws.cells.get(ref).date(dt));
    }
    return this;
  }

  style(s: any) {
    if (this.isMerged) {
      this.ws.cells.get(this.excelRefs[0]).style(s);
    } else {
      this.excelRefs.forEach(ref => this.ws.cells.get(ref).style(s));
    }
    return this;
  }
}
