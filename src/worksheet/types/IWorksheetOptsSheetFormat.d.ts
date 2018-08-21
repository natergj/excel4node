export interface IWorksheetOptsSheetFormat {
  baseColWidth: number; // Defaults to 10. Specifies the number of characters of the maximum digit width of the normal style's font. This value does not include margin padding or extra padding for gridlines. It is only the number of characters.,
  defaultColWidth: number;
  defaultRowHeight: number;
  thickBottom: boolean; // 'True' if rows have a thick bottom border by default.
  thickTop: boolean; // 'True' if rows have a thick top border by default.
}
