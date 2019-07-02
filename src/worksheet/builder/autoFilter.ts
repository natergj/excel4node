import Worksheet from '../worksheet';
import { XMLElement } from 'xmlbuilder';
import { getExcelRowCol } from '../../utils';

// TODO implement
export default function addAutoFilter(xml: XMLElement, ws: Worksheet) {
  // §18.3.1.2 autoFilter (AutoFilter Settings)
  const o = ws.opts.autoFilter;
  if (typeof o.ref !== 'string') {
    return;
  }

  const [startCell, endCell] = o.ref.split(':');
  const startRow = getExcelRowCol(startCell).row;
  const startCol = getExcelRowCol(startCell).col;
  const endRow = getExcelRowCol(endCell).row;
  const endCol = getExcelRowCol(endCell).col;
}
