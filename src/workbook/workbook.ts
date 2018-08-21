import { Worksheet, IWorksheetOpts } from '../worksheet';
import { IWorkbookOptions } from '../workbook';

// Default Options for Workbook
const workbookDefaultOpts = {
  jszip: {
    compression: 'DEFLATE',
  },
  defaultFont: {
    color: 'FF000000',
    name: 'Calibri',
    size: 12,
    family: 'roman',
  },
  dateFormat: 'm/d/yy',
  workbookView: {
    activeTab: 1,
    autoFilterDateGrouping: true,
    firstSheet: 1,
    minimized: false,
    showHorizontalScroll: true,
    showSheetTabs: true,
    showVerticalScroll: true,
    tabRatio: 1,
    visibility: 'visible',
    windowHeight: 1000,
    windowWidth: 1000,
    xWindow: 100,
    yWindow: 100,
  },
};

export default class Workbook {
  opts: IWorkbookOptions;
  sheets: Worksheet[];

  constructor(opts: Partial<IWorkbookOptions> = {}) {
    this.opts = {
      ...workbookDefaultOpts,
      ...opts,
    };
    this.sheets = [];
  }

  addWorksheet(name: string, opts: Partial<IWorksheetOpts>) {
    this.sheets.push(
      new Worksheet({
        name,
        opts,
        wb: this,
      })
    );
  }

  write() {
    console.log(this);
  }
}
