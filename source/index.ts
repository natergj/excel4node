/* REFERENCES
    http://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf
    http://www.ecma-international.org/publications/standards/Ecma-376.htm
    http://www.openoffice.org/sc/excelfileformat.pdf 
    http://officeopenxml.com/anatomyofOOXML-xlsx.php
*/

/* 
    Code references specifications sections from ECMA-376 2nd edition doc
    ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference.pdf
    found in ECMA-376 2nd edition Part 1 download at
    http://www.ecma-international.org/publications/standards/Ecma-376.htm
    Sections are referenced in code comments with § 
*/

export {
  getExcelAlpha,
  getExcelCellRef,
  getExcelRowCol,
  getExcelTS,
} from './lib/utils';
export { Workbook } from './lib/workbook/workbook';
export {
  paperSize as PaperSize,
  cellComments as CellComment,
  printError as PrintError,
  pageOrder as PageOrder,
  orientation as Orientation,
  pane as Pane,
  paneState as PaneState,
  borderStyle as BorderStyle,
  fillPattern as PatternType,
  positiveUniversalMeasure as PositiveUniversalMeasure,
} from './lib/types/index';
export {
  vertical as VerticalAlignment,
  horizontal as HorizontalAlignment,
} from './lib/types/alignment';
