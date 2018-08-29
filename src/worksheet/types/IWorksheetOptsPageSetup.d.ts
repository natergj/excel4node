import { PAPER_SIZE } from '../../types';

// §18.3.1.64
export interface IWorksheetOptsPageSetup {
  /**
   * Print black and white.
   */
  blackAndWhite: boolean;

  /**
   * This attribute specifies how to print cell comments.
   */
  cellComments: 'asDisplayed' | 'atEnd' | 'none';

  /**
   * Number of copies to print.
   */
  copies: number;

  /**
   * Print draft quality.
   */
  draft: boolean;

  /**
   * This enumeration specifies how to display cells with errors when printing the worksheet
   */
  errors: 'blank' | 'dash' | 'displayed' | 'NA';

  /**
   * Page number for first printed page. If no value is specified, then 'automatic' is assumed.
   */
  firstPageNumber: number;

  /**
   * Number of vertical pages to fit on.
   */
  fitToHeight: number;

  /**
   * Number of horizontal pages to fit on.
   */
  fitToWidth: number;

  /**
   * Horizontal print resolution of the device.
   */
  horizontalDpi: number;

  /**
   * Relationship Id of the devMode printer settings part.
   */
  id: string;

  /**
   * Orientation of the page.
   */
  orientation: 'default' | 'landscape' | 'portrait';

  /**
   * Order of printed pages.
   */
  pageOrder: number;

  /**
   * Height of custom paper as a number followed by a unit identifier.
   * @example: 297mm
   * @example 11in
   */
  paperHeight: string;

  paperSize: PAPER_SIZE;

  /**
   * Width of custom paper as a number followed by a unit identifier.
   * @example 21cm
   * @example 8.5in
   */
  paperWidth: string;

  /**
   * Print scaling. This attribute is restricted to values ranging from 10 to 400.
   * Values represent percentages. 10 = 10%, 120 = 120%;
   */
  scale: number;

  /**
   * Use firstPageNumber value for first page number, and do not auto number the pages.
   */
  useFirstPageNumber: boolean;

  /**
   * Use the printer’s defaults settings for page setup values and don't use the default values specified in the schema.
   * [Example: If dpi is not present or specified in the XML, the application must not assume 600dpi as specified in the schema as a default and instead must let the printer specify the default dpi. end example]
   */
  usePrinterDefaults: boolean;

  /**
   * Vertical print resolution of the device.
   */
  verticalDpi: number;
}
