import { Workbook } from '../../workbook';
import { IWorksheetOpts } from './IWorksheetOpts';

export interface IWorksheetConstructorOpts {
  name: string;
  opts: Partial<IWorksheetOpts>;
  wb: Workbook;
}
