import xmlbuilder from 'xmlbuilder';
import { Worksheet } from '../../worksheet';
import { getDataStream } from '../../utils/dataStream';
import { IWorkbookBuilder } from '.';

// Required as stated in §12.2
export default function addContentTypes(builder: IWorkbookBuilder) {
  const { wb, xlsx } = builder;

  const contentTypesAdded = [];
  const extensionsAdded = [];
  const dataStream = getDataStream();
  const writer = xmlbuilder.streamWriter(dataStream);

  xlsx.file('[Content_Types].xml', dataStream);
  dataStream.on('error', e => {
    throw e;
  });

  const xml = xmlbuilder
    .create('Types', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
      allowSurrogateChars: true,
    })
    .att('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

  wb.sheets.forEach((ws: Worksheet) => {
    if (ws.drawingCollection.length > 0) {
      ws.drawingCollection.drawings.forEach(d => {
        if (extensionsAdded.indexOf(d.extension) < 0) {
          const typeRef = d.contentType + '.' + d.extension;
          if (contentTypesAdded.indexOf(typeRef) < 0) {
            this.xml
              .ele('Default')
              .att('ContentType', d.contentType)
              .att('Extension', d.extension);
          }
          extensionsAdded.push(d.extension);
        }
      });
    }
  });

  xml
    .ele('Default')
    .att('ContentType', 'application/xml')
    .att('Extension', 'xml');

  xml
    .ele('Default')
    .att('Extension', 'rels')
    .att('ContentType', 'application/vnd.openxmlformats-package.relationships+xml');

  xml
    .ele('Override')
    .att('PartName', '/xl/workbook.xml')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');

  wb.sheets.forEach((ws: Worksheet) => {
    xml
      .ele('Override')
      .att('PartName', `/xl/worksheets/sheet${ws.sheetId}.xml`)
      .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');

    if (ws.drawingCollection.length > 0) {
      this.xml
        .ele('Override')
        .att('PartName', '/xl/drawings/drawing' + ws.sheetId + '.xml')
        .att('ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml');
    }
  });

  xml
    .ele('Override')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')
    .att('PartName', '/xl/styles.xml');

  xml
    .ele('Override')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
    .att('PartName', '/xl/sharedStrings.xml');

  xml.end(writer);
  dataStream.end();
}
