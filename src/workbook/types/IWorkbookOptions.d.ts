import { JSZipFileOptions } from 'jszip';
import { IWorkbookView } from './IWorkbookView';
import { Logger, LogLevel } from '../../utils';

export interface IWorkbookOptions {
  jszip?: JSZipFileOptions;
  logger?: Logger;
  logLevel?: LogLevel;
  defaultFont?: any;
  dateFormat?: string;
  workbookView?: IWorkbookView;
}
