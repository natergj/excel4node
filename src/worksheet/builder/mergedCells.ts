import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';

export default function addMergedCells(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.55 mergeCells (Merge Cells)
  if (ws.mergedCells.length > 0) {
    let ele = xml.ele('mergeCells').att('count', ws.mergedCells.length);
    ws.mergedCells.forEach(cr => {
      ele
        .ele('mergeCell')
        .att('ref', cr)
        .up();
    });
    ele.up();
  }
}
