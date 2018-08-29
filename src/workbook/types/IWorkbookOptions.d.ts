import { JSZipFileOptions } from 'jszip';
import { IWorkbookView } from './IWorkbookView';
import Logger, { LogLevel } from '../../utils/logger';

export interface IWorkbookOptions {
  jszip?: JSZipFileOptions;
  logger?: Logger;
  logLevel?: LogLevel;
  defaultFont?: any;
  dateFormat?: string;
  workbookView?: IWorkbookView;
}
