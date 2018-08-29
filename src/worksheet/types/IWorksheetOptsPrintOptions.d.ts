// §18.3.1.70

export interface IWorksheetOptsPrintOptions {
  /**
   * Used in conjunction with gridLinesSet. If both gridLines and gridlinesSet are true, then grid lines shall print.
   */
  gridLines: boolean;

  /**
   * Used in conjunction with gridLines. If both gridLines and gridLinesSet are true, then grid lines shall print
   */
  gridLinesSet: boolean;

  /**
   * Should Heading be printed
   */
  headings: boolean;

  /**
   * Should data be centered horizontally when printed
   */
  horizontalCentered: boolean;

  /**
   * Should data be centered vertically when printed
   */
  verticalCentered: boolean;
}
