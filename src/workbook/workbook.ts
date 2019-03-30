import { Worksheet, IWorksheetOpts } from '../worksheet';
import buildWorkbook from './builder';
import { SimpleLogger, ILogger, LogLevel } from '../utils/logger';
import { StyleSheet } from '../style';
import { JSZipFileOptions } from 'jszip';

interface IWorkbookView {
  activeTab: number;
  autoFilterDateGrouping: boolean;
  firstSheet: number;
  minimized: boolean;
  showHorizontalScroll: boolean;
  showSheetTabs: boolean;
  showVerticalScroll: boolean;
  tabRatio: number;
  visibility: string;
  windowHeight: number;
  windowWidth: number;
  xWindow: number;
  yWindow: number;
}

interface IWorkbookOptions {
  jszip?: JSZipFileOptions;
  logger?: ILogger;
  logLevel?: LogLevel;
  defaultFont?: any;
  dateFormat?: string;
  workbookView?: IWorkbookView;
}

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
    // Defaults defined in §18.2.30
    activeTab: 0,
    autoFilterDateGrouping: true,
    firstSheet: 0,
    minimized: false,
    showHorizontalScroll: true,
    showSheetTabs: true,
    showVerticalScroll: true,
    tabRatio: 600,
    visibility: 'visible',
    windowHeight: 17440, // default of Excel 2016 for Mac
    windowWidth: 28040, // default of Excel 2016 for Mac
    xWindow: 5180, // default of Excel 2016 for Mac
    yWindow: 3060, // default of Excel 2016 for Mac
    showComments: true,
  },
};

export default class Workbook {
  opts: IWorkbookOptions;
  sheets: Map<string, Worksheet>;
  sharedStrings: Map<string | any[], number>;
  definedNameCollection: any;
  dxfCollection: any[];
  styles: any[];
  styleData: any;

  constructor(opts: Partial<IWorkbookOptions> = {}) {
    this.opts = {
      ...workbookDefaultOpts,
      ...opts,
      logger: new SimpleLogger(opts.logLevel || 0),
    };
    this.sheets = new Map();
    this.sharedStrings = new Map();
    // TODO implement
    this.definedNameCollection = {
      isEmpty: true,
    };
    this.styles = [];
    this.styleData = new StyleSheet();
  }

  addWorksheet(name: string, opts: Partial<IWorksheetOpts> = {}) {
    this.sheets.set(
      name,
      new Worksheet({
        name,
        opts,
        wb: this,
      }),
    );
    return this.sheets.get(name);
  }

  write(name: string) {
    try {
      buildWorkbook(name, this);
    } catch (err) {
      this.opts.logger.error('Error building workbook package.', err);
    }
  }
}
