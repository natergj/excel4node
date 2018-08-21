import { JSZipFileOptions } from 'jszip';
import { IWorkbookView } from './IWorkbookView';

export interface IWorkbookOptions {
  jszip?: JSZipFileOptions;
  defaultFont: any;
  dateFormat: string;
  workbookView: IWorkbookView;
}
