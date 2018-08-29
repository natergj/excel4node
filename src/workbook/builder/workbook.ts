import xmlbuilder from 'xmlbuilder';
import IWorkbookBuilder from '../types/IWorkbookBuilder';
import { getDataStream } from '../../utils/dataStream';
import { boolToInt } from '../../utils/excel4node';
import { Worksheet } from '../../worksheet';

export default function addWorkbookXml(builder: IWorkbookBuilder) {
  const { wb, xlsx } = builder;

  const dataStream = getDataStream();
  const writer = xmlbuilder.streamWriter(dataStream);

  const xml = xmlbuilder.create('workbook', {
    version: '1.0',
    encoding: 'UTF-8',
    standalone: true,
  });
  xml.att('mc:Ignorable', 'x15');
  xml.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  xml.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
  xml.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  xml.att('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');

  // bookViews (§18.2.1)
  let booksViewEle = xml.ele('bookViews');
  let workbookViewEle = booksViewEle.ele('workbookView');
  if (wb.opts.workbookView) {
    const viewOpts = wb.opts.workbookView;
    if (viewOpts.activeTab) {
      workbookViewEle.att('activeTab', viewOpts.activeTab);
    }
    if (viewOpts.autoFilterDateGrouping) {
      workbookViewEle.att('autoFilterDateGrouping', boolToInt(viewOpts.autoFilterDateGrouping));
    }
    if (viewOpts.firstSheet) {
      workbookViewEle.att('firstSheet', viewOpts.firstSheet);
    }
    if (viewOpts.minimized) {
      workbookViewEle.att('minimized', boolToInt(viewOpts.minimized));
    }
    if (viewOpts.showHorizontalScroll) {
      workbookViewEle.att('showHorizontalScroll', boolToInt(viewOpts.showHorizontalScroll));
    }
    if (viewOpts.showSheetTabs) {
      workbookViewEle.att('showSheetTabs', boolToInt(viewOpts.showSheetTabs));
    }
    if (viewOpts.showVerticalScroll) {
      workbookViewEle.att('showVerticalScroll', boolToInt(viewOpts.showVerticalScroll));
    }
    if (viewOpts.tabRatio) {
      workbookViewEle.att('tabRatio', viewOpts.tabRatio);
    }
    if (viewOpts.visibility) {
      workbookViewEle.att('visibility', viewOpts.visibility);
    }
    if (viewOpts.windowWidth) {
      workbookViewEle.att('windowWidth', viewOpts.windowWidth);
    }
    if (viewOpts.windowHeight) {
      workbookViewEle.att('windowHeight', viewOpts.windowHeight);
    }
    if (viewOpts.xWindow) {
      workbookViewEle.att('xWindow', viewOpts.xWindow);
    }
    if (viewOpts.yWindow) {
      workbookViewEle.att('yWindow', viewOpts.yWindow);
    }
  }

  let sheetsEle = xml.ele('sheets');
  wb.sheets.forEach((ws: Worksheet) => {
    const sheet = sheetsEle
      .ele('sheet')
      .att('name', ws.name)
      .att('sheetId', ws.sheetId)
      .att('r:id', `rId${ws.sheetId}`);

    if (ws.opts.hidden) {
      sheet.att('state', 'hidden');
    }
  });

  if (!wb.definedNameCollection.isEmpty) {
    wb.definedNameCollection.addToXMLele(xml);
  }

  xlsx.folder('xl').file('workbook.xml', dataStream);
  xml.end(writer);
  dataStream.end();
}
