import { PAPER_SIZE } from '../../types';

export interface IWorksheetOptsPageSetup {
  blackAndWhite: boolean; //
  cellComments: 'none' | 'asDisplayed' | 'atEnd';
  copies: number; //  How many copies to print
  draft: boolean; //  Should quality be draft
  errors: 'displayed' | 'blank' | 'dash' | 'NA';
  firstPageNumber: number; //  Should the page number of the first page be printed
  fitToHeight: number; //  Number of vertical pages to fit to
  fitToWidth: number; //  Number of horizontal pages to fit to
  horizontalDpi: number; //
  orientation: 'default' | 'portrait' | 'landscape';
  pageOrder: 'downThenOver' | 'overThenDown';
  paperHeight: string; //  Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. '10.5cm'
  paperSize: PAPER_SIZE; //  see lib/types/paperSize.js for all types and descriptions of types. setting paperSize overrides paperHeight and paperWidth settings
  paperWidth: string; //  Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. '10.5cm'
  scale: number; //  zoom of worksheet
  useFirstPageNumber: boolean;
  usePrinterDefaults: boolean;
  verticalDpi: number;
}
