import { Workbook } from '../workbook';
import { defaultWorksheetOpts } from './defaultWorksheetOptions';
import { IWorksheetOpts, IWorksheetConstructorOpts } from './types';
import IWorkbookBuilder from '../workbook/types/IWorkbookBuilder';
import { addWorksheetFile, addWorksheetRelsFile } from './builder';
import { cellAccessor } from '../cell';

export default class Worksheet {
  name: string;
  opts: Partial<IWorksheetOpts>;
  wb: Workbook;
  cells: Map<string, any>;
  mergedCells: string[];
  columns: any[];
  rows: any[];
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
    this.rows = [];
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
    return cellAccessor(this, startRow, startCol, endRow, endCol, isMerged);
  }

  addSheetToXlsxPackage(builder: IWorkbookBuilder) {
    addWorksheetFile(builder, this);
    addWorksheetRelsFile(builder, this);
  }
}
