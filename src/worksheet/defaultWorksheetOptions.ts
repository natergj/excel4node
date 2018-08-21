import { PAPER_SIZE } from '../types';
import { IWorksheetOpts } from './types/IWorksheetOpts';

export const defaultWorksheetOpts: Partial<IWorksheetOpts> = {
  pageSetup: {
    paperSize: PAPER_SIZE.LETTER_PAPER,
  },
};
