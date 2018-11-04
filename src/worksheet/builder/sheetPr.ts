import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';

export default function addSheetPr(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.82 sheetPr (Sheet Properties)
  const options = ws.opts;

  // Check if any option that would require the sheetPr element to be added exists
  if (
    options.pageSetup.fitToHeight ||
    options.pageSetup.fitToWidth ||
    options.outline.summaryBelow ||
    options.outline.summaryRight ||
    options.autoFilter
  ) {
    const ele = xml.ele('sheetPr');
    if (options.autoFilter.ref) {
      ele.att('enableFormatConditionsCalculation', 1);
      ele.att('filterMode', 1);
    }

    if (options.outline.summaryBelow || options.outline.summaryRight) {
      const outlineEle = ele.ele('outlinePr');
      outlineEle.att('applyStyles', 1);
      outlineEle.att('summaryBelow', options.outline.summaryBelow === true ? 1 : 0);
      outlineEle.att('summaryRight', options.outline.summaryRight === true ? 1 : 0);
      outlineEle.up();
    }

    // §18.3.1.65 pageSetUpPr (Page Setup Properties)
    if (options.pageSetup.fitToHeight || options.pageSetup.fitToWidth) {
      ele
        .ele('pageSetUpPr')
        .att('fitToPage', 1)
        .up();
    }
    ele.up();
  }
}
