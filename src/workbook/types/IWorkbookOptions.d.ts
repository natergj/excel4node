import { JSZipFileOptions } from 'jszip';
import { IWorkbookView } from './IWorkbookView';
import { ILogger, LogLevel } from '../../utils';

export interface IWorkbookOptions {
  jszip?: JSZipFileOptions;
  logger?: ILogger;
  logLevel?: LogLevel;
  defaultFont?: any;
  dateFormat?: string;
  workbookView?: IWorkbookView;
}
