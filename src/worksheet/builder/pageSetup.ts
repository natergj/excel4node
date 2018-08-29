import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { PAPER_SIZE } from '../../types/papersize';
import { boolToInt } from '../../utils/excel4node';

export default function addPageSetup(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.63 pageSetup (Page Setup Settings)

  let o = ws.opts.pageSetup;
  const isPageSetupRequired = Object.keys(o)
    .map(k => o[k] !== null)
    .includes(true);

  if (isPageSetupRequired === true) {
    let psEle = xml.ele('pageSetup');
    o.paperSize !== null ? psEle.att('paperSize', PAPER_SIZE[o.paperSize]) : null;
    o.paperHeight !== null ? psEle.att('paperHeight', o.paperHeight) : null;
    o.paperWidth !== null ? psEle.att('paperWidth', o.paperWidth) : null;
    o.scale !== null ? psEle.att('scale', o.scale) : null;
    o.firstPageNumber !== null ? psEle.att('firstPageNumber', o.firstPageNumber) : null;
    o.fitToWidth !== null ? psEle.att('fitToWidth', o.fitToWidth) : null;
    o.fitToHeight !== null ? psEle.att('fitToHeight', o.fitToHeight) : null;
    o.pageOrder !== null ? psEle.att('pageOrder', o.pageOrder) : null;
    o.orientation !== null ? psEle.att('orientation', o.orientation) : null;
    o.usePrinterDefaults !== null ? psEle.att('usePrinterDefaults', boolToInt(o.usePrinterDefaults)) : null;
    o.blackAndWhite !== null ? psEle.att('blackAndWhite', boolToInt(o.blackAndWhite)) : null;
    o.draft !== null ? psEle.att('draft', boolToInt(o.draft)) : null;
    o.cellComments !== null ? psEle.att('cellComments', o.cellComments) : null;
    o.useFirstPageNumber !== null ? psEle.att('useFirstPageNumber', boolToInt(o.useFirstPageNumber)) : null;
    o.errors !== null ? psEle.att('errors', o.errors) : null;
    o.horizontalDpi !== null ? psEle.att('horizontalDpi', o.horizontalDpi) : null;
    o.verticalDpi !== null ? psEle.att('verticalDpi', o.verticalDpi) : null;
    o.copies !== null ? psEle.att('copies', o.copies) : null;
    psEle.up();
  }
}
