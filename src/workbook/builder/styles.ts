import xmlbuilder from 'xmlbuilder';
import IWorkbookBuilder from '../types/IWorkbookBuilder';
import { getDataStream } from '../../utils/dataStream';

// Required as stated in §12.2
export default function addContentTypes(builder: IWorkbookBuilder) {
  const { wb, xlsx } = builder;

  const dataStream = getDataStream();
  const writer = xmlbuilder.streamWriter(dataStream);

  const xml = xmlbuilder
    .create('styleSheet', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
      allowSurrogateChars: true,
    })
    .att('mc:Ignorable', 'x14ac')
    .att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    .att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    .att('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

  if (wb.styleData.numFmts.length > 0) {
    let nfXML = xml.ele('numFmts').att('count', wb.styleData.numFmts.length);
    wb.styleData.numFmts.forEach(nf => {
      nf.addToXMLele(nfXML);
    });
  }

  if (wb.styleData.fonts.length > 0) {
    let fontXML = xml.ele('fonts').att('count', wb.styleData.fonts.length);
    wb.styleData.fonts.forEach(f => {
      f.addToXMLele(fontXML);
    });
  }

  if (wb.styleData.fills.length > 0) {
    let fillXML = xml.ele('fills').att('count', wb.styleData.fills.length);
    wb.styleData.fills.forEach(f => {
      let fXML = fillXML.ele('fill');
      f.addToXMLele(fXML);
    });
  }

  if (wb.styleData.borders.length > 0) {
    let borderXML = xml.ele('borders').att('count', wb.styleData.borders.length);
    wb.styleData.borders.forEach(b => {
      b.addToXMLele(borderXML);
    });
  }

  let cellXfsXML = xml.ele('cellXfs').att('count', wb.styles.length);
  wb.styles.forEach(s => {
    s.addXFtoXMLele(cellXfsXML);
  });

  // TODO implement
  // if (wb.dxfCollection.length > 0) {
  //   wb.dxfCollection.addToXMLele(xml);
  // }

  xlsx.folder('xl').file('styles.xml', dataStream);

  xml.end(writer);
  dataStream.end();
}
