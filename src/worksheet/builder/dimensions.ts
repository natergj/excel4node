import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { getExcelAlpha } from '../../utils/excel4node';

export default function addDimensions(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.35 dimension (Worksheet Dimensions)
  let firstCell = 'A1';
  let lastCell = `${getExcelAlpha(ws.lastUsedCol)}${ws.lastUsedRow}`;
  let ele = xml.ele('dimension');
  ele.att('ref', `${firstCell}:${lastCell}`);
  ele.up();
}
