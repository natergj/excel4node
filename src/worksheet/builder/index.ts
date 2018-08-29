import xmlbuilder from 'xmlbuilder';
import Worksheet from '../worksheet';
import IWorkbookBuilder from '../../workbook/types/IWorkbookBuilder';
import { getDataStream } from '../../utils/dataStream';
import addSheetPr from './sheetPr';
import addDimensions from './dimensions';
import addSheetViews from './sheetViews';
import addSheetFormatPr from './sheetFormatPr';
import addColumns from './columns';
import addSheetData from './sheetData';
import addSheetProtection from './sheetProtection';
import addAutoFilter from './autoFilter';
import addMergedCells from './mergedCells';
import addConditionalFormatting from './conditionalFormatting';
import addDataValidation from './dataValidation';
import addHyperlinks from './hyperlinks';
import addPrintOptions from './printOptions';
import addPageMargins from './pageMargins';
import addPageSetup from './pageSetup';
import addHeaderFooter from './headerFooter';
import addDrawing from './drawing';

export function addWorksheetFile(builder: IWorkbookBuilder, ws: Worksheet) {
  ws.wb.opts.logger.debug(`adding worksheet file: ${ws.name}`);
  const dataStream = getDataStream();
  const writer = xmlbuilder.streamWriter(dataStream);

  builder.xlsx
    .folder('xl')
    .folder('worksheets')
    .file(`sheet${ws.sheetId}.xml`, dataStream);

  dataStream.on('error', err => {
    throw err;
  });

  const xml = xmlbuilder
    .create('Types', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
      allowSurrogateChars: true,
    })
    .ele('worksheet')
    .att('mc:Ignorable', 'x14ac')
    .att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    .att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    .att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    .att('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

  // Order is important!
  addSheetPr(xml, ws);
  addDimensions(xml, ws);
  addSheetViews(xml, ws);
  addSheetFormatPr(xml, ws);
  addColumns(xml, ws);
  addSheetData(xml, ws);
  addSheetProtection(xml, ws);
  addAutoFilter(xml, ws);
  addMergedCells(xml, ws);
  addConditionalFormatting(xml, ws);
  addDataValidation(xml, ws);
  addHyperlinks(xml, ws);
  addPrintOptions(xml, ws);
  addPageMargins(xml, ws);
  addPageSetup(xml, ws);
  addHeaderFooter(xml, ws);
  addDrawing(xml, ws);

  xml.end(writer);
  dataStream.end();
}

export function addWorksheetRelsFile(builder: IWorkbookBuilder, ws: Worksheet) {
  ws.wb.opts.logger.debug(`adding worksheet rels file: ${ws.name}`);
  const dataStream = getDataStream();
  const writer = xmlbuilder.streamWriter(dataStream);

  builder.xlsx
    .folder('xl')
    .folder('worksheets')
    .folder('_rels')
    .file(`sheet${ws.sheetId}.xml.rels`, dataStream);

  dataStream.on('error', err => {
    throw err;
  });

  const xml = xmlbuilder.create('Relationships', {
    version: '1.0',
    encoding: 'UTF-8',
    standalone: true,
    allowSurrogateChars: true,
  });
  xml.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

  ws.relationships.forEach((r, i) => {
    let rId = 'rId' + (i + 1);
    // TODO implement
    if (r === 'hyperlink') {
      xml
        .ele('Relationship')
        .att('Id', rId)
        .att('Target', r.location)
        .att('TargetMode', 'External')
        .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
    } else if (r === 'drawing') {
      xml
        .ele('Relationship')
        .att('Id', rId)
        .att('Target', '../drawings/drawing' + ws.sheetId + '.xml')
        .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing');
    }
  });

  xml.end(writer);
  dataStream.end();
}
