import { IWorksheetOptsMargins } from './IWorksheetOptsMargins';
import { IWorksheetOptsPrintOptions } from './IWorksheetOptsPrintOptions';
import { IWorksheetOptsHeaderFooter } from './IWorksheetOptsHeaderFooter';
import { IWorksheetOptsPageSetup } from './IWorksheetOptsPageSetup';
import { IWorksheetOptsSheetView } from './IWorksheetOptsSheetView';
import { IWorksheetOptsSheetFormat } from './IWorksheetOptsSheetFormat';
import { IWorksheetOptsSheetProtection } from './IWorksheetOptsSheetProtection';
import { IWorksheetOptsOutline } from './IWorksheetOptsOutline';
import { IWorksheetAutoFilter } from './IWorksheetAutoFilter';

export interface IWorksheetOpts {
  autoFilter: Partial<IWorksheetAutoFilter>;
  margins: Partial<IWorksheetOptsMargins>;
  printOptions: Partial<IWorksheetOptsPrintOptions>;
  headerFooter: Partial<IWorksheetOptsHeaderFooter>;
  pageSetup: Partial<IWorksheetOptsPageSetup>;
  sheetView: Partial<IWorksheetOptsSheetView>;
  sheetFormat: Partial<IWorksheetOptsSheetFormat>;
  sheetProtection: Partial<IWorksheetOptsSheetProtection>;
  outline: Partial<IWorksheetOptsOutline>;
  /**
   * Flag indicated whether to not include a spans attribute to the row definition in the XML. helps with very large documents.
   * @default false
   * @TJS-type boolean
   */
  disableRowSpansOptimization: boolean;
  /**
   * Flag indicating whether to not hide the worksheet within the workbook.
   *
   * @default false
   * @TJS-type boolean
   */
  hidden: boolean;
}
