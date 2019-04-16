import xmlbuilder from 'xmlbuilder';
import { getDataStream } from '../../utils/dataStream';
import { boolToInt } from '../../utils';
import { Worksheet } from '../../worksheet';
import { IWorkbookBuilder } from '.';

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

  // workbookPr (§18.2.28)
  const workbookPrEle = xml.ele('workbookPr');
  if (wb.workbookProperties.allowRefreshQuery) {
    workbookPrEle.att('allowRefreshQuery', boolToInt(wb.workbookProperties.allowRefreshQuery));
  }
  if (wb.workbookProperties.autoCompressPictures) {
    workbookPrEle.att('autoCompressPictures', boolToInt(wb.workbookProperties.autoCompressPictures));
  }
  if (wb.workbookProperties.backupFile) {
    workbookPrEle.att('backupFile', boolToInt(wb.workbookProperties.backupFile));
  }
  if (wb.workbookProperties.checkCompatibility) {
    workbookPrEle.att('checkCompatibility', boolToInt(wb.workbookProperties.checkCompatibility));
  }
  if (wb.workbookProperties.codeName) {
    workbookPrEle.att('codeName', wb.workbookProperties.codeName);
  }
  if (wb.workbookProperties.date1904) {
    workbookPrEle.att('date1904', boolToInt( wb.workbookProperties.date1904));
  }
  if (wb.workbookProperties.dateCompatibility) {
    workbookPrEle.att('dateCompatibility', boolToInt(wb.workbookProperties.dateCompatibility));
  }
  if (wb.workbookProperties.filterPrivacy) {
    workbookPrEle.att('filterPrivacy', boolToInt(wb.workbookProperties.filterPrivacy));
  }
  if (wb.workbookProperties.hidePivotFieldList) {
    workbookPrEle.att('hidePivotFieldList', boolToInt(wb.workbookProperties.hidePivotFieldList));
  }
  if (wb.workbookProperties.promptedSolutions) {
    workbookPrEle.att('promptedSolutions', boolToInt(wb.workbookProperties.promptedSolutions));
  }
  if (wb.workbookProperties.publishItems) {
    workbookPrEle.att('publishItems', boolToInt(wb.workbookProperties.publishItems));
  }
  if (wb.workbookProperties.showBorderUnselectedTables) {
    workbookPrEle.att('showBorderUnselectedTables', boolToInt(wb.workbookProperties.showBorderUnselectedTables));
  }
  if (wb.workbookProperties.showInkAnnotation) {
    workbookPrEle.att('showInkAnnotation', boolToInt(wb.workbookProperties.showInkAnnotation));
  }
  if (wb.workbookProperties.showObjects) {
    workbookPrEle.att('showObjects', wb.workbookProperties.showObjects);
  }
  if (wb.workbookProperties.showPivotChartFilter) {
    workbookPrEle.att('showPivotChartFilter', boolToInt(wb.workbookProperties.showPivotChartFilter));
  }
  if (wb.workbookProperties.updateLinks) {
    workbookPrEle.att('updateLinks', wb.workbookProperties.updateLinks);
  }

  // bookViews (§18.2.1)
  const booksViewEle = xml.ele('bookViews');
  const workbookViewEle = booksViewEle.ele('workbookView');
  for (const workbookView of wb.workbookViews) {
    if (workbookView.activeTab) {
      workbookViewEle.att('activeTab', workbookView.activeTab);
    }
    if (workbookView.autoFilterDateGrouping) {
      workbookViewEle.att('autoFilterDateGrouping', boolToInt(workbookView.autoFilterDateGrouping));
    }
    if (workbookView.firstSheet) {
      workbookViewEle.att('firstSheet', workbookView.firstSheet);
    }
    if (workbookView.minimized) {
      workbookViewEle.att('minimized', boolToInt(workbookView.minimized));
    }
    if (workbookView.showHorizontalScroll) {
      workbookViewEle.att('showHorizontalScroll', boolToInt(workbookView.showHorizontalScroll));
    }
    if (workbookView.showSheetTabs) {
      workbookViewEle.att('showSheetTabs', boolToInt(workbookView.showSheetTabs));
    }
    if (workbookView.showVerticalScroll) {
      workbookViewEle.att('showVerticalScroll', boolToInt(workbookView.showVerticalScroll));
    }
    if (workbookView.tabRatio) {
      workbookViewEle.att('tabRatio', workbookView.tabRatio);
    }
    if (workbookView.visibility) {
      workbookViewEle.att('visibility', workbookView.visibility);
    }
    if (workbookView.windowWidth) {
      workbookViewEle.att('windowWidth', workbookView.windowWidth);
    }
    if (workbookView.windowHeight) {
      workbookViewEle.att('windowHeight', workbookView.windowHeight);
    }
    if (workbookView.xWindow) {
      workbookViewEle.att('xWindow', workbookView.xWindow);
    }
    if (workbookView.yWindow) {
      workbookViewEle.att('yWindow', workbookView.yWindow);
    }
  }

  const sheetsEle = xml.ele('sheets');
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
