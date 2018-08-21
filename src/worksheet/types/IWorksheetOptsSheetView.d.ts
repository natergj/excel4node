import { IWorksheetOptsSheetViewPane } from './IWorksheetOptsSheetViewPane';

export interface IWorksheetOptsSheetView {
  pane: Partial<IWorksheetOptsSheetViewPane>;
  rightToLeft: boolean; // Flag indicating whether the sheet is in 'right to left' display mode. When in this mode, Column A is on the far right, Column B ;is one column left of Column A, and so on. Also, information in cells is displayed in the Right to Left format.
  showGridLines: boolean; // Flag indicating whether the sheet should have grid lines enabled or disabled during view.
  zoomScale: number; // Defaults to 100
  zoomScaleNormal: number; // Defaults to 100
  zoomScalePageLayoutView: number; // Defaults to 100
}
