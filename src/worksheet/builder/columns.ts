import Worksheet from '../worksheet';
import { XMLElement } from 'xmlbuilder';
import { boolToInt } from '../../utils';

export default function addColumns(xml: XMLElement, ws: Worksheet) {
  // §18.3.1.17 cols (Column Information)
  if (ws.columnCount > 0) {
    const colsEle = xml.ele('cols');

    for (const col of ws.columns) {
      const colEle = colsEle.ele('col');

      if (col.min !== null) {
        colEle.att('min', col.min);
      }
      if (col.max !== null) {
        colEle.att('max', col.max);
      }
      if (col.width !== null) {
        colEle.att('width', col.width);
      }
      if (col.style !== null) {
        colEle.att('style', col.style);
      }
      if (col.hidden !== null) {
        colEle.att('hidden', boolToInt(col.hidden));
      }
      if (col.customWidth !== null) {
        colEle.att('customWidth', boolToInt(col.customWidth));
      }
      if (col.outlineLevel !== null) {
        colEle.att('outlineLevel', col.outlineLevel);
      }
      if (col.collapsed !== null) {
        colEle.att('collapsed', boolToInt(col.collapsed));
      }
      colEle.up();
    }
    colsEle.up();
  }
}
