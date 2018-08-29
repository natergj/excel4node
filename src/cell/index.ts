import { Worksheet } from '../worksheet';
import { Cell } from './cell';
import { getExcelAlpha } from '../utils/excel4node';

export function cellAccessor(ws: Worksheet, startRow: number, startCol: number, endRow?: number, endCol?: number, isMerged?: boolean) {
  this.ws = ws;
  this.startRow = startRow;
  this.startCol = startCol;
  this.endRow = endRow || startRow;
  this.endCol = endCol || startCol;
  this.isMerged = !!isMerged || false;
  this.excelRefs = [];

  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      let ref = `${getExcelAlpha(c)}${r}`;
      if (!this.ws.cells.get(ref)) {
        this.ws.cells.set(ref, new Cell(r, c));
      }

      this.excelRefs.push(ref);
    }
  }
}
