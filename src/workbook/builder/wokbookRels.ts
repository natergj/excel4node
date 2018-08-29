import xmlbuilder from 'xmlbuilder';
import IWorkbookBuilder from '../types/IWorkbookBuilder';
import { getDataStream } from '../../utils/dataStream';
import { Worksheet } from '../../worksheet';

export default function addWorkbookXml(builder: IWorkbookBuilder) {
  const { wb, xlsx } = builder;

  const dataStream = getDataStream();
  const writer = xmlbuilder.streamWriter(dataStream);

  const xml = xmlbuilder
    .create('Relationships', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    })
    .att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

  xml
    .ele('Relationship')
    .att('Id', `rId${wb.sheets.size + 1}`)
    .att('Target', 'sharedStrings.xml')
    .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings');

  // xml
  //   .ele('Relationship')
  //   .att('Id', `rId${wb.sheets.size + 2}`)
  //   .att('Target', 'styles.xml')
  //   .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');

  wb.sheets.forEach((ws: Worksheet) => {
    xml
      .ele('Relationship')
      .att('Id', `rId${ws.sheetId}`)
      .att('Target', `worksheets/sheet${ws.sheetId}.xml`)
      .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
  });

  xlsx
    .folder('xl')
    .folder('_rels')
    .file('workbook.xml.rels', dataStream);

  xml.end(writer);
  dataStream.end();
}
