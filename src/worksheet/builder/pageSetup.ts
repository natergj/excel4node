import Worksheet from '../worksheet';
import { XMLElement } from 'xmlbuilder';
import { PAPER_SIZE } from '../../types/papersize';
import { boolToInt } from '../../utils';

export default function addPageSetup(xml: XMLElement, ws: Worksheet) {
  // §18.3.1.63 pageSetup (Page Setup Settings)

  const o = ws.opts.pageSetup;
  const isPageSetupRequired = Object.keys(o)
    .map(k => o[k] !== null)
    .includes(true);

  if (isPageSetupRequired === true) {
    const psEle = xml.ele('pageSetup');
    if (o.paperSize !== null) {
      psEle.att('paperSize', PAPER_SIZE[o.paperSize]);
    }
    if (o.paperHeight !== null) {
      psEle.att('paperHeight', o.paperHeight);
    }
    if (o.paperWidth !== null) {
      psEle.att('paperWidth', o.paperWidth);
    }
    if (o.scale !== null) {
      psEle.att('scale', o.scale);
    }
    if (o.firstPageNumber !== null) {
      psEle.att('firstPageNumber', o.firstPageNumber);
    }
    if (o.fitToWidth !== null) {
      psEle.att('fitToWidth', o.fitToWidth);
    }
    if (o.fitToHeight !== null) {
      psEle.att('fitToHeight', o.fitToHeight);
    }
    if (o.pageOrder !== null) {
      psEle.att('pageOrder', o.pageOrder);
    }
    if (o.orientation !== null) {
      psEle.att('orientation', o.orientation);
    }
    if (o.usePrinterDefaults !== null) {
      psEle.att('usePrinterDefaults', boolToInt(o.usePrinterDefaults));
    }
    if (o.blackAndWhite !== null) {
      psEle.att('blackAndWhite', boolToInt(o.blackAndWhite));
    }
    if (o.draft !== null) {
      psEle.att('draft', boolToInt(o.draft));
    }
    if (o.cellComments !== null) {
      psEle.att('cellComments', o.cellComments);
    }
    if (o.useFirstPageNumber !== null) {
      psEle.att('useFirstPageNumber', boolToInt(o.useFirstPageNumber));
    }
    if (o.errors !== null) {
      psEle.att('errors', o.errors);
    }
    if (o.horizontalDpi !== null) {
      psEle.att('horizontalDpi', o.horizontalDpi);
    }
    if (o.verticalDpi !== null) {
      psEle.att('verticalDpi', o.verticalDpi);
    }
    if (o.copies !== null) {
      psEle.att('copies', o.copies);
    }
    psEle.up();
  }
}
