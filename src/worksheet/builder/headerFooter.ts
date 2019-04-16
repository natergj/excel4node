import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { boolToInt } from '../../utils';

export default function addHyperlinks(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.46 headerFooter (Header Footer Settings)
  const o = ws.opts.headerFooter;
  const isHeaderFooterRequired = Object.keys(o)
    .map(k => o[k] !== null)
    .includes(true);

  if (isHeaderFooterRequired === true) {
    const hfEle = xml.ele('headerFooter');

    if (o.alignWithMargins !== null) {
      hfEle.att('alignWithMargins', boolToInt(o.alignWithMargins));
    }
    if (o.differentFirst !== null) {
      hfEle.att('differentFirst', boolToInt(o.differentFirst));
    }
    if (o.differentOddEven !== null) {
      hfEle.att('differentOddEven', boolToInt(o.differentOddEven));
    }
    if (o.scaleWithDoc !== null) {
      hfEle.att('scaleWithDoc', boolToInt(o.scaleWithDoc));
    }

    if (o.oddHeader !== null) {
      hfEle
        .ele('oddHeader')
        .text(o.oddHeader)
        .up();
    }
    if (o.oddFooter !== null) {
      hfEle
        .ele('oddFooter')
        .text(o.oddFooter)
        .up();
    }
    if (o.evenHeader !== null) {
      hfEle
        .ele('evenHeader')
        .text(o.evenHeader)
        .up();
    }
    if (o.evenFooter !== null) {
      hfEle
        .ele('evenFooter')
        .text(o.evenFooter)
        .up();
    }
    if (o.firstHeader !== null) {
      hfEle
        .ele('firstHeader')
        .text(o.firstHeader)
        .up();
    }
    if (o.firstFooter !== null) {
      hfEle
        .ele('firstFooter')
        .text(o.firstFooter)
        .up();
    }
    hfEle.up();
  }
}
