import { Worksheet } from '../worksheet';

// §18.3.1.73 row (Row)
export class Row {
  /**
   * true if the rows 1 level of outlining deeper than the current row are in the collapsed outline state.
   * It means that the rows which are 1 outline level deeper (numerically higher value) than the current
   * row are currently hidden due to a collapsed outline state.
   */
  collapsed?: boolean;
  /**
   * true if the row style should be applied.
   */
  customFormat?: boolean;
  /**
   * true if the row height has been manually set.
   */
  customHeight?: boolean;
  /**
   * true if the row is hidden, e.g., due to a collapsed outline or by manually selecting and hiding a row.
   */
  hidden?: boolean;
  /**
   * Row height measured in point size. There is no margin padding on row height.
   */
  ht?: number;
  /**
   * Outlining level of the row, when outlining is on. See description of outlinePr element's
   * summaryBelow and summaryRight attributes for detailed information.
   */
  outlineLevel?: number;
  /**
   * true if the row should show phonetic.
   */
  ph?: boolean;
  /**
   * Row index. Indicates to which row in the sheet this <row> definition corresponds.
   */
  r: number;
  /**
   * Index to style record for the row (only applied if customFormat attribute is true)
   */
  s?: number;
  /**
   * Optimization only, and not required. Specifies the range of non-empty columns (in the format X:Y)
   * for the block of rows to which the current row belongs. To achieve the optimization,
   * span attribute values in a single block should be the same.
   */
  spans?: string;
  /**
   * if any cell in the row has a medium or thick bottom border, or if any cell in the row directly
   * below the current row has a thick top border.
   */
  thickBot?: boolean;
  /**
   * True if the row has a medium or thick top border, or if any cell in the row directly above
   * the current row has a thick bottom border.
   */
  thickTop?: boolean;
  /**
   * Array of cell references included in row
   */
  cellRefs: Set<string>;

  constructor(row: number, options: Partial<Row> = {}) {
    Object.keys(options).forEach(o => this[o] = options[o]);
    this.r = row;
    this.cellRefs = options.cellRefs instanceof Array ? new Set(options.cellRefs) : new Set();
  }
}
