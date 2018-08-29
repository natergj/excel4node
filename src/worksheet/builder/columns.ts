import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { boolToInt } from '../../utils/excel4node';

export default function addColumns(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.17 cols (Column Information)
  if (ws.columnCount > 0) {
    let colsEle = xml.ele('cols');

    for (let col of ws.columns) {
      let colEle = colsEle.ele('col');

      col.min !== null ? colEle.att('min', col.min) : null;
      col.max !== null ? colEle.att('max', col.max) : null;
      col.width !== null ? colEle.att('width', col.width) : null;
      col.style !== null ? colEle.att('style', col.style) : null;
      col.hidden !== null ? colEle.att('hidden', boolToInt(col.hidden)) : null;
      col.customWidth !== null ? colEle.att('customWidth', boolToInt(col.customWidth)) : null;
      col.outlineLevel !== null ? colEle.att('outlineLevel', col.outlineLevel) : null;
      col.collapsed !== null ? colEle.att('collapsed', boolToInt(col.collapsed)) : null;
      colEle.up();
    }
    colsEle.up();
  }
}
