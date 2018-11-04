import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { boolToInt } from '../../utils/excel4node';

export default function addSheetFormatPr(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.81 sheetFormatPr (Sheet Format Properties)
  const o = ws.opts.sheetFormat;
  const ele = xml.ele('sheetFormatPr');

  if (o.baseColWidth !== null) {
    ele.att('baseColWidth', o.baseColWidth);
  }
  if (o.defaultColWidth !== null) {
    ele.att('defaultColWidth', o.defaultColWidth);
  }
  if (o.defaultRowHeight !== null) {
    ele.att('defaultRowHeight', o.defaultRowHeight);
  } else {
    ele.att('defaultRowHeight', 16);
  }
  if (o.thickBottom !== null) {
    ele.att('thickBottom', boolToInt(o.thickBottom));
  }
  if (o.thickTop !== null) {
    ele.att('thickTop', boolToInt(o.thickTop));
  }

  if (typeof o.defaultRowHeight === 'number') {
    ele.att('customHeight', '1');
  }
  ele.up();
}
