import * as _ from 'lodash';
import { Cell } from './cell';
import { Row } from '../row/row';
import { Column } from '../column/column';
import { Style } from '../style/style';
import * as utils from '../utils';
import * as util from 'util';

function stringSetter(val) {
  const logger = this.ws.wb.logger;
  const chars = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;
  const chr = val.match(chars);
  let cleanedVal = val;
  if (chr) {
    logger.warn(
      'Invalid Character for XML "' + chr + '" in string "' + val + '"',
    );
    cleanedVal = val.replace(chr, '');
  }

  if (typeof val !== 'string') {
    logger.warn(
      'Value sent to String function of cells %s was not a string, it has type of %s',
      JSON.stringify(this.excelRefs),
      typeof val,
    );
    cleanedVal = '';
  }

  // Remove Control characters, they aren't understood by xmlbuilder
  cleanedVal = cleanedVal.replace(
    /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/,
    '',
  );

  if (!this.merged) {
    this.cells.forEach((c) => {
      c.string(this.ws.wb.getStringIndex(val));
    });
  } else {
    const c = this.cells[0];
    c.string(this.ws.wb.getStringIndex(val));
  }
  return this;
}

function complexStringSetter(val) {
  if (!this.merged) {
    this.cells.forEach((c) => {
      c.string(this.ws.wb.getStringIndex(val));
    });
  } else {
    const c = this.cells[0];
    c.string(this.ws.wb.getStringIndex(val));
  }
  return this;
}

function numberSetter(val) {
  if (val === undefined || parseFloat(val) !== val) {
    throw new TypeError(
      util.format(
        'Value sent to Number function of cells %s was not a number, it has type of %s and value of %s',
        JSON.stringify(this.excelRefs),
        typeof val,
        val,
      ),
    );
  }
  const cleanedVal = parseFloat(val);

  if (!this.merged) {
    this.cells.forEach((c, i) => {
      c.number(cleanedVal);
    });
  } else {
    const c = this.cells[0];
    c.number(cleanedVal);
  }
  return this;
}

function booleanSetter(val) {
  if (
    val === undefined ||
    typeof (
      val.toString().toLowerCase() === 'true' ||
      (val.toString().toLowerCase() === 'false' ? false : val)
    ) !== 'boolean'
  ) {
    throw new TypeError(
      util.format(
        'Value sent to Bool function of cells %s was not a bool, it has type of %s and value of %s',
        JSON.stringify(this.excelRefs),
        typeof val,
        val,
      ),
    );
  }
  const cleanedVal = val.toString().toLowerCase() === 'true';

  if (!this.merged) {
    this.cells.forEach((c, i) => {
      c.bool(cleanedVal.toString());
    });
  } else {
    const c = this.cells[0];
    c.bool(cleanedVal.toString());
  }
  return this;
}

function formulaSetter(val) {
  if (typeof val !== 'string') {
    throw new TypeError(
      util.format(
        'Value sent to Formula function of cells %s was not a string, it has type of %s',
        JSON.stringify(this.excelRefs),
        typeof val,
      ),
    );
  }
  if (this.merged !== true) {
    this.cells.forEach((c, i) => {
      c.formula(val);
    });
  } else {
    const c = this.cells[0];
    c.formula(val);
  }

  return this;
}

function dateSetter(val) {
  const thisDate = new Date(val);
  if (isNaN(thisDate.getTime())) {
    throw new TypeError(
      util.format(
        'Invalid date sent to date function of cells. %s could not be converted to a date.',
        val,
      ),
    );
  }
  if (this.merged !== true) {
    this.cells.forEach((c, i) => {
      c.date(thisDate);
    });
  } else {
    const c = this.cells[0];
    c.date(thisDate);
  }
  return styleSetter.bind(this)({
    numberFormat: '[$-409]' + this.ws.wb.opts.dateFormat,
  });
}

function styleSetter(val) {
  let thisStyle;
  if (val instanceof Style) {
    thisStyle = val.toObject();
  } else if (val instanceof Object) {
    thisStyle = val;
  } else {
    throw new TypeError(
      util.format(
        'Parameter sent to Style function must be an instance of a Style or a style configuration object',
      ),
    );
  }

  const borderEdges = {} as any;
  if (thisStyle.border && thisStyle.border.outline) {
    borderEdges.left = this.firstCol;
    borderEdges.right = this.lastCol;
    borderEdges.top = this.firstRow;
    borderEdges.bottom = this.lastRow;
  }

  this.cells.forEach((c) => {
    if (thisStyle.border && thisStyle.border.outline) {
      const thisCellsBorder = {} as any;
      if (c.row === borderEdges.top && thisStyle.border.top) {
        thisCellsBorder.top = thisStyle.border.top;
      }
      if (c.row === borderEdges.bottom && thisStyle.border.bottom) {
        thisCellsBorder.bottom = thisStyle.border.bottom;
      }
      if (c.col === borderEdges.left && thisStyle.border.left) {
        thisCellsBorder.left = thisStyle.border.left;
      }
      if (c.col === borderEdges.right && thisStyle.border.right) {
        thisCellsBorder.right = thisStyle.border.right;
      }
      thisStyle.border = thisCellsBorder;
    }

    if (c.s === 0) {
      const thisCellStyle = this.ws.wb.createStyle(thisStyle);
      c.style(thisCellStyle.ids.cellXfs);
    } else {
      const curStyle = this.ws.wb.styles[c.s];
      const newStyleOpts = _.merge({}, curStyle.toObject(), thisStyle);
      const mergedStyle = this.ws.wb.createStyle(newStyleOpts);
      c.style(mergedStyle.ids.cellXfs);
    }
  });
  return this;
}

function hyperlinkSetter(url: string, displayStr: string = url, tooltip) {
  this.excelRefs.forEach((ref) => {
    this.ws.hyperlinkCollection.add({
      tooltip,
      ref,
      location: url,
      display: displayStr,
    });
  });
  stringSetter.bind(this)(displayStr);
  return styleSetter.bind(this)({
    font: {
      color: 'Blue',
      underline: true,
    },
  });
}

function mergeCells(cellBlock) {
  const excelRefs = cellBlock.excelRefs;
  if (excelRefs instanceof Array && excelRefs.length > 0) {
    excelRefs.sort(utils.sortCellRefs);

    const cellRange = excelRefs[0] + ':' + excelRefs[excelRefs.length - 1];
    const rangeCells = excelRefs;

    let okToMerge = true;
    cellBlock.ws.mergedCells.forEach((cr) => {
      // Check to see if currently merged cells contain cells in new merge request
      const curCells = utils.getAllCellsInExcelRange(cr);
      const intersection = utils.arrayIntersectSafe(rangeCells, curCells);
      if (intersection.length > 0) {
        okToMerge = false;
        cellBlock.ws.wb.logger.error(
          `Invalid Range for: ${cellRange}. ` +
            `Some cells in this range are already included in another merged cell range: ${cr}.`,
        );
      }
    });
    if (okToMerge) {
      cellBlock.ws.mergedCells.push(cellRange);
    }
  } else {
    throw new TypeError(
      util.format(
        'excelRefs variable sent to mergeCells function must be an array with length > 0',
      ),
    );
  }
}

/**
 * @class cellBlock
 */
class cellBlock {
  public ws;
  public cells;
  public excelRefs;
  public merged;

  constructor() {
    this.ws;
    this.cells = [];
    this.excelRefs = [];
    this.merged = false;
  }

  get matrix() {
    const matrix = [];
    const tmpObj = {};
    this.cells.forEach((c) => {
      if (!tmpObj[c.row]) {
        tmpObj[c.row] = [];
      }
      tmpObj[c.row].push(c);
    });
    const rows = Object.keys(tmpObj);
    rows.forEach((r) => {
      tmpObj[r].sort((a, b) => {
        return a.col - b.col;
      });
      matrix.push(tmpObj[r]);
    });
    return matrix;
  }

  get firstRow() {
    let firstRow;
    this.cells.forEach((c) => {
      if (c.row < firstRow || firstRow === undefined) {
        firstRow = c.row;
      }
    });
    return firstRow;
  }

  get lastRow() {
    let lastRow;
    this.cells.forEach((c) => {
      if (c.row > lastRow || lastRow === undefined) {
        lastRow = c.row;
      }
    });
    return lastRow;
  }

  get firstCol() {
    let firstCol;
    this.cells.forEach((c) => {
      if (c.col < firstCol || firstCol === undefined) {
        firstCol = c.col;
      }
    });
    return firstCol;
  }

  get lastCol() {
    let lastCol;
    this.cells.forEach((c) => {
      if (c.col > lastCol || lastCol === undefined) {
        lastCol = c.col;
      }
    });
    return lastCol;
  }

  /**
   * @alias cellBlock.string
   * @func cellBlock.string
   * @param {String} val Value of String
   * @returns {cellBlock} Block of cells with attached methods
   */
  string(val) {
    if (val instanceof Array) {
      return complexStringSetter.bind(this)(val);
    }
    return stringSetter.bind(this)(val);
  }

  /**
   * @alias cellBlock.style
   * @func cellBlock.style
   * @param {Object} style One of a Style instance or an object with Style parameters
   * @returns {cellBlock} Block of cells with attached methods
   */
  style = styleSetter;

  /**
   * @alias cellBlock.number
   * @func cellBlock.number
   * @param {Number} val Value of Number
   * @returns {cellBlock} Block of cells with attached methods
   */
  number = numberSetter;

  /**
   * @alias cellBlock.bool
   * @func cellBlock.bool
   * @param {Boolean} val Value of Boolean
   * @returns {cellBlock} Block of cells with attached methods
   */
  bool = booleanSetter;

  /**
   * @alias cellBlock.formula
   * @func cellBlock.formula
   * @param {String} val Excel style formula as string
   * @returns {cellBlock} Block of cells with attached methods
   */
  formula = formulaSetter;

  /**
   * @alias cellBlock.date
   * @func cellBlock.date
   * @param {Date} val Value of Date
   * @returns {cellBlock} Block of cells with attached methods
   */
  date = dateSetter;

  /**
   * @alias cellBlock.link
   * @func cellBlock.link
   * @param {String} url Value of Hyperlink URL
   * @param {String} displayStr Value of String representation of URL
   * @param {String} tooltip Value of text to display as hover
   * @returns {cellBlock} Block of cells with attached methods
   */
  link = hyperlinkSetter;
}

/**
 * Module repesenting a Cell Accessor
 * @alias Worksheet.cell
 * @namespace
 * @func Worksheet.cell
 * @desc Access a range of cells in order to manipulate values
 * @param {Number} row1 Row of top left cell
 * @param {Number} col1 Column of top left cell
 * @param {Number} row2 Row of bottom right cell (optional)
 * @param {Number} col2 Column of bottom right cell (optional)
 * @param {Boolean} isMerged Merged the cell range into a single cell
 * @returns {cellBlock}
 */
export function cellAccessor(
  row1: number,
  col1: number,
  row2: number = row1,
  col2: number = col1,
  isMerged: boolean = false,
) {
  const theseCells = new cellBlock();
  theseCells.ws = this;

  if (row2 > this.lastUsedRow) {
    this.lastUsedRow = row2;
  }

  if (col2 > this.lastUsedCol) {
    this.lastUsedCol = col2;
  }

  for (let r = row1; r <= row2; r += 1) {
    for (let c = col1; c <= col2; c += 1) {
      const ref = `${utils.getExcelAlpha(c)}${r}`;
      if (!this.cells[ref]) {
        this.cells[ref] = new Cell(r, c);
      }
      if (!this.rows[r]) {
        this.rows[r] = new Row(r, this);
      }
      if (this.rows[r].cellRefs.indexOf(ref) < 0) {
        this.rows[r].cellRefs.push(ref);
      }

      theseCells.cells.push(this.cells[ref]);
      theseCells.excelRefs.push(ref);
    }
  }
  if (isMerged) {
    theseCells.merged = true;
    mergeCells(theseCells);
  }

  return theseCells;
}
