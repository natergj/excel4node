import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';

export default function addPrintOptions(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.70 printOptions (Print Options)

  let o = ws.opts.printOptions;
  const isPrintOptionsRequired = Object.keys(o)
    .map(k => o[k] !== null)
    .includes(true);

  if (isPrintOptionsRequired === true) {
    let poEle = xml.ele('printOptions');
    o.horizontalCentered === true ? poEle.att('horizontalCentered', 1) : null;
    o.verticalCentered === true ? poEle.att('verticalCentered', 1) : null;
    o.headings === true ? poEle.att('headings', 1) : null;
    if (o.gridLines === true || o.gridLinesSet === true) {
      poEle.att('gridLines', 1);
      poEle.att('gridLinesSet', 1);
    }
    poEle.up();
  }
}
