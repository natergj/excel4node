export interface IWorksheetOptsSheetProtection {
  autoFilter: boolean; // True means that that user will be unable to modify this setting
  deleteColumns: boolean; // True means that that user will be unable to modify this setting
  deleteRows: boolean; // True means that that user will be unable to modify this setting
  formatCells: boolean; // True means that that user will be unable to modify this setting
  formatColumns: boolean; // True means that that user will be unable to modify this setting
  formatRows: boolean; // True means that that user will be unable to modify this setting
  insertColumns: boolean; // True means that that user will be unable to modify this setting
  insertHyperlinks: boolean; // True means that that user will be unable to modify this setting
  insertRows: boolean; // True means that that user will be unable to modify this setting
  objects: boolean; // True means that that user will be unable to modify this setting
  password: string; // Password used to protect sheet
  pivotTables: boolean; // True means that that user will be unable to modify this setting
  scenarios: boolean; // True means that that user will be unable to modify this setting
  selectLockedCells: boolean; // True means that that user will be unable to modify this setting
  selectUnlockedCells: boolean; // True means that that user will be unable to modify this setting
  sheet: boolean; // True means that that user will be unable to modify this setting
  sort: boolean; // True means that that user will be unable to modify this setting
}
