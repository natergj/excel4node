/* REFERENCES
    http://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf
    http://www.ecma-international.org/publications/standards/Ecma-376.htm
    http://www.openoffice.org/sc/excelfileformat.pdf 
    http://officeopenxml.com/anatomyofOOXML-xlsx.php
*/

/* 
    Code references specifications sections from ECMA-376 2nd edition doc
    ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference.pdf
    found in ECMA-376 2nd edition Part 1 download at http://www.ecma-international.org/publications/standards/Ecma-376.htm
    Sections are referenced in code comments with § 
*/

import {
  getExcelAlpha,
  getExcelCellRef,
  getExcelRowCol,
  getExcelTS,
} from "./lib/utils";
import * as types from "./lib/types/index";
import Workbook from "./lib/workbook/workbook";

export default {
  Workbook,
  getExcelRowCol: getExcelRowCol,
  getExcelAlpha: getExcelAlpha,
  getExcelTS: getExcelTS,
  getExcelCellRef: getExcelCellRef,
  PaperSize: types.paperSize,
  CellComment: types.cellComments,
  PrintError: types.printError,
  PageOrder: types.pageOrder,
  Orientation: types.orientation,
  Pane: types.pane,
  PaneState: types.paneState,
  HorizontalAlignment: types.alignment.horizontal,
  VerticalAlignment: types.alignment.vertical,
  BorderStyle: types.borderStyle,
  PresetColorVal: types.excelColor,
  PatternType: types.fillPattern,
  PositiveUniversalMeasure: types.positiveUniversalMeasure,
};
