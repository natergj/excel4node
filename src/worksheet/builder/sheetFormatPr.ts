import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { boolToInt } from '../../utils/excel4node';

export default function addSheetFormatPr(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.81 sheetFormatPr (Sheet Format Properties)
  let o = ws.opts.sheetFormat;
  let ele = xml.ele('sheetFormatPr');

  o.baseColWidth !== null ? ele.att('baseColWidth', o.baseColWidth) : null;
  o.defaultColWidth !== null ? ele.att('defaultColWidth', o.defaultColWidth) : null;
  o.defaultRowHeight !== null ? ele.att('defaultRowHeight', o.defaultRowHeight) : ele.att('defaultRowHeight', 16);
  o.thickBottom !== null ? ele.att('thickBottom', boolToInt(o.thickBottom)) : null;
  o.thickTop !== null ? ele.att('thickTop', boolToInt(o.thickTop)) : null;

  if (typeof o.defaultRowHeight === 'number') {
    ele.att('customHeight', '1');
  }
  ele.up();
}
