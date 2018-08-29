import Workbook from '../workbook';
import JSZip from 'jszip';

export default interface IWorkbookBuilder {
  wb: Workbook;
  xlsx: JSZip;
}
