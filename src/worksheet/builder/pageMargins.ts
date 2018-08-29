import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';

// TODO implement
export default function addPageMargins(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.62 pageMargins (Page Margins)
  let o = ws.opts.margins;

  xml
    .ele('pageMargins')
    .att('left', o.left)
    .att('right', o.right)
    .att('top', o.top)
    .att('bottom', o.bottom)
    .att('header', o.header)
    .att('footer', o.footer)
    .up();
}
