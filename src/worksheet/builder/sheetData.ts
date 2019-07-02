import Worksheet from '../worksheet';
import { XMLElement } from 'xmlbuilder';
import { sortCellRefs } from '../../utils';

export default function addSheetData(xml: XMLElement, ws: Worksheet) {
  // §18.3.1.80 sheetData (Sheet Data)
  const ele = xml.ele('sheetData');

  // TODO asynchronous loop
  for (let r = 1; r <= ws.rows.size; r++) {
    const thisRow = ws.row(r);
    const sortedRefs = Array.from(thisRow.cellRefs).sort(sortCellRefs);

    const rEle = ele.ele('row');

    rEle.att('r', thisRow.r);
    if (ws.opts.disableRowSpansOptimization !== true && thisRow.spans) {
      rEle.att('spans', thisRow.spans);
    }

    if (thisRow.s !== undefined) {
      rEle.att('s', thisRow.s);
    }
    if (thisRow.customFormat !== undefined) {
      rEle.att('customFormat', thisRow.customFormat);
    }
    if (thisRow.ht !== undefined) {
      rEle.att('ht', thisRow.ht);
    }
    if (thisRow.hidden !== undefined) {
      rEle.att('hidden', thisRow.hidden);
    }
    if (thisRow.customHeight === true || typeof ws.opts.sheetFormat.defaultRowHeight === 'number') {
      rEle.att('customHeight', 1);
    }
    if (thisRow.outlineLevel !== undefined) {
      rEle.att('outlineLevel', thisRow.outlineLevel);
    }
    if (thisRow.collapsed !== undefined) {
      rEle.att('collapsed', thisRow.collapsed);
    }
    if (thisRow.thickTop !== undefined) {
      rEle.att('thickTop', thisRow.thickTop);
    }
    if (thisRow.thickBot !== undefined) {
      rEle.att('thickBot', thisRow.thickBot);
    }

    sortedRefs.forEach(ref => {
      ws.cells.get(ref).addToXMLele(rEle);
    });

    rEle.up();
  }
}
