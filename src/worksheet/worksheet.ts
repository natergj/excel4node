import { Workbook } from '../workbook';
import { defaultWorksheetOpts } from './defaultWorksheetOptions';
import { IWorksheetOpts, IWorksheetConstructorOpts } from './types';
import { addWorksheetFile, addWorksheetRelsFile } from './builder';
import { CellAccessor, Cell } from '../cell';
import { Row } from '../row';
import { IWorkbookBuilder } from '../workbook/builder';

export default class Worksheet {
  name: string;
  opts: Partial<IWorksheetOpts>;
  wb: Workbook;
  cells: Map<string, Cell>;
  mergedCells: string[];
  columns: any[];
  rows: Map<number, Row>;
  drawingCollection: any;
  sheetId: number;
  relationships: any[];

  constructor(init: IWorksheetConstructorOpts) {
    const { name, wb, opts } = init;
    this.name = name;
    this.wb = wb;
    this.opts = {
      ...defaultWorksheetOpts,
      ...opts,
    };
    this.cells = new Map();
    this.mergedCells = [];
    this.columns = [];
    this.rows = new Map();
    this.sheetId = this.wb.sheets.size + 1;
    this.drawingCollection = [];
    this.relationships = [];
  }

  get columnCount() {
    return this.columns.length;
  }

  get lastUsedCol() {
    // TODO implement
    return 1;
  }
  get lastUsedRow() {
    // TODO implement
    return 1;
  }

  cell(startRow: number, startCol: number, endRow?: number, endCol?: number, isMerged?: boolean) {
    return new CellAccessor(this, startRow, startCol, endRow, endCol, isMerged);
  }

  row(row: number) {
    if (!this.rows.has(row)) {
      this.rows.set(row, new Row(row));
    }
    return this.rows.get(row);
  }

  addSheetToXlsxPackage(builder: IWorkbookBuilder) {
    addWorksheetFile(builder, this);
    addWorksheetRelsFile(builder, this);
  }
}
