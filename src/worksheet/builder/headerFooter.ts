import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { boolToInt } from '../../utils/excel4node';

export default function addHyperlinks(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.46 headerFooter (Header Footer Settings)
  let o = ws.opts.headerFooter;
  const isHeaderFooterRequired = Object.keys(o)
    .map(k => o[k] !== null)
    .includes(true);

  if (isHeaderFooterRequired === true) {
    let hfEle = xml.ele('headerFooter');

    o.alignWithMargins !== null ? hfEle.att('alignWithMargins', boolToInt(o.alignWithMargins)) : null;
    o.differentFirst !== null ? hfEle.att('differentFirst', boolToInt(o.differentFirst)) : null;
    o.differentOddEven !== null ? hfEle.att('differentOddEven', boolToInt(o.differentOddEven)) : null;
    o.scaleWithDoc !== null ? hfEle.att('scaleWithDoc', boolToInt(o.scaleWithDoc)) : null;

    o.oddHeader !== null
      ? hfEle
          .ele('oddHeader')
          .text(o.oddHeader)
          .up()
      : null;
    o.oddFooter !== null
      ? hfEle
          .ele('oddFooter')
          .text(o.oddFooter)
          .up()
      : null;
    o.evenHeader !== null
      ? hfEle
          .ele('evenHeader')
          .text(o.evenHeader)
          .up()
      : null;
    o.evenFooter !== null
      ? hfEle
          .ele('evenFooter')
          .text(o.evenFooter)
          .up()
      : null;
    o.firstHeader !== null
      ? hfEle
          .ele('firstHeader')
          .text(o.firstHeader)
          .up()
      : null;
    o.firstFooter !== null
      ? hfEle
          .ele('firstFooter')
          .text(o.firstFooter)
          .up()
      : null;
    hfEle.up();
  }
}
