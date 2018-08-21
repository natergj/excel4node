import { Workbook } from '../workbook';
import { defaultWorksheetOpts } from './defaultWorksheetOptions';
import { IWorksheetOpts, IWorksheetConstructorOpts } from './types';

export default class Worksheet {
  name: string;
  opts: Partial<IWorksheetOpts>;
  wb: Workbook;

  constructor(init: IWorksheetConstructorOpts) {
    const { name, wb, opts } = init;
    this.name = name;
    this.wb = wb;
    this.opts = {
      ...defaultWorksheetOpts,
      ...opts,
    };
  }
}
