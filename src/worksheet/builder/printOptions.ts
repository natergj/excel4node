import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';

export default function addPrintOptions(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.70 printOptions (Print Options)

  const o = ws.opts.printOptions;
  const isPrintOptionsRequired = Object.keys(o)
    .map(k => o[k] !== null)
    .includes(true);

  if (isPrintOptionsRequired === true) {
    const poEle = xml.ele('printOptions');
    if (o.horizontalCentered === true) {
      poEle.att('horizontalCentered', 1);
    }
    if (o.verticalCentered === true) {
      poEle.att('verticalCentered', 1);
    }
    if (o.headings === true) {
      poEle.att('headings', 1);
    }
    if (o.gridLines === true || o.gridLinesSet === true) {
      poEle.att('gridLines', 1);
      poEle.att('gridLinesSet', 1);
    }
    poEle.up();
  }
}
