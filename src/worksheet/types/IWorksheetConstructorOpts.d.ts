import { Workbook } from '../../workbook';
import { IWorksheetOpts } from './IWorksheetOpts';

interface IWorksheetConstructorOpts {
  name: string;
  opts: Partial<IWorksheetOpts>;
  wb: Workbook;
}
