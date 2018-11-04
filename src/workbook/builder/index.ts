import fs from 'fs';
import JSZip from 'jszip';
import Workbook from '../workbook';
import addContentTypes from './contentTypes';
import addRootRels from './rootRels';
import addWorksheets from './worksheets';
import addWorkbookXml from './workbook';
import addWorkbookRels from './wokbookRels';
import addSharedStrings from './sharedStrings';
import addStyles from './styles';

export default async function buildWorkbook(name: string, wb: Workbook) {
  if (wb.sheets.size === 0) {
    wb.addWorksheet('Sheet 1');
  }
  const builder = {
    wb,
    xlsx: new JSZip(),
  };

  addContentTypes(builder);
  addWorksheets(builder);
  addRootRels(builder);
  addWorkbookXml(builder);
  addWorkbookRels(builder);
  addSharedStrings(builder);
  addStyles(builder);

  const xlsxContent = await builder.xlsx.generateAsync({
    type: 'uint8array',
    streamFiles: true,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    compression: 'DEFLATE',
  });

  const fd = fs.createWriteStream(name, { encoding: 'utf-8' });

  fd.on('finish', () => {
    wb.opts.logger.info(`${name} file written`);
  });
  fd.write(xlsxContent);
  fd.end();
}
