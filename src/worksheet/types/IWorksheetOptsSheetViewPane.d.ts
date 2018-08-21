export interface IWorksheetOptsSheetViewPane {
  activePane: 'bottomLeft' | 'bottomRight' | 'topLeft' | 'topRight';
  state: 'split' | 'frozen' | 'frozenSplit';
  topLeftCell: string; // Cell Reference i.e. 'A1'
  xSplit: string; // Horizontal position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of columns visible in the top pane.
  ySplit: string; // Vertical position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of rows visible in the left pane.
}
