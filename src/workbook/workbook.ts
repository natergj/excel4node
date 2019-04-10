import { JSZipGeneratorOptions } from 'jszip';
import { Worksheet, IWorksheetOpts } from '../worksheet';
import buildWorkbook from './builder';
import { SimpleLogger, ILogger, LogLevel } from '../utils/logger';
import { StyleSheet } from '../style';
import { Font, WorkbookView, WorkbookProperties } from '../types';
import { getJsZipProperties, getDefaultFont, dateFormat, createWorkbookView } from './defaults';

interface IWorkbookOptions {
  jszip?: Partial<JSZipGeneratorOptions>;
  logger?: ILogger;
  logLevel?: LogLevel;
  defaultFont?: Partial<Font>;
  dateFormat?: string;
  defaultWorkbookView?: Partial<WorkbookView>;
  workbookProperties?: Partial<WorkbookProperties>;
}

export default class Workbook {
  jszip: Partial<JSZipGeneratorOptions>;
  logger: ILogger;
  logLevel: LogLevel;
  defaultFont: Partial<Font>;
  dateFormat: string;
  workbookViews: Set<WorkbookView>;
  workbookProperties: Partial<WorkbookProperties>;
  sheets: Map<string, Worksheet>;
  sharedStrings: Map<string, number>;
  definedNameCollection: any;
  dxfCollection: any[];
  styles: any[];
  styleData: any;

  constructor(opts: Partial<IWorkbookOptions> = {}) {
    this.jszip = getJsZipProperties(opts.jszip);
    this.logger = opts.logger || new SimpleLogger(opts.logLevel || 'silent');
    this.defaultFont = getDefaultFont(opts.defaultFont);
    this.dateFormat = opts.dateFormat || dateFormat;
    this.workbookProperties = opts.workbookProperties || {};
    this.workbookViews = new Set();
    this.sheets = new Map();
    this.sharedStrings = new Map();
    this.workbookViews.add(createWorkbookView(opts.defaultWorkbookView));

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
      this.logger.error('Error building workbook package.', err);
    }
  }
}
