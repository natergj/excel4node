import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { sortCellRefs } from '../../utils/excel4node';

export default function addSheetData(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.80 sheetData (Sheet Data)
  let ele = xml.ele('sheetData');

  // TODO asynchronous loop
  for (var r = 0; r < ws.rows.length; r++) {
    let thisRow = ws.rows[r];
    thisRow.cellRefs.sort(sortCellRefs);

    let rEle = ele.ele('row');

    rEle.att('r', thisRow.r);
    if (ws.opts.disableRowSpansOptimization !== true && thisRow.spans) {
      rEle.att('spans', thisRow.spans);
    }
    thisRow.s !== null ? rEle.att('s', thisRow.s) : null;
    thisRow.customFormat !== null ? rEle.att('customFormat', thisRow.customFormat) : null;
    thisRow.ht !== null ? rEle.att('ht', thisRow.ht) : null;
    thisRow.hidden !== null ? rEle.att('hidden', thisRow.hidden) : null;
    thisRow.customHeight === true || typeof ws.opts.sheetFormat.defaultRowHeight === 'number' ? rEle.att('customHeight', 1) : null;
    thisRow.outlineLevel !== null ? rEle.att('outlineLevel', thisRow.outlineLevel) : null;
    thisRow.collapsed !== null ? rEle.att('collapsed', thisRow.collapsed) : null;
    thisRow.thickTop !== null ? rEle.att('thickTop', thisRow.thickTop) : null;
    thisRow.thickBot !== null ? rEle.att('thickBot', thisRow.thickBot) : null;

    for (var i = 0; i < thisRow.cellRefs.length; i++) {
      ws.cells[thisRow.cellRefs[i]].addToXMLele(rEle);
    }

    rEle.up();
  }
}
