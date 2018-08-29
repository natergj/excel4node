import xmlbuilder from 'xmlbuilder';
import IWorkbookBuilder from '../types/IWorkbookBuilder';
import { getDataStream } from '../../utils/dataStream';

// Required as stated in §12.2
export default function addContentTypes(builder: IWorkbookBuilder) {
  const { xlsx } = builder;

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
    .att('Id', 'rId1')
    .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
    .att('Target', 'xl/workbook.xml');

  xlsx.folder('_rels').file('.rels', dataStream);

  xml.end(writer);
  dataStream.end();
}
