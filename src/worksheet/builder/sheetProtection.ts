import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';
import { boolToInt, getHashOfPassword } from '../../utils/excel4node';

export default function addSheetProtection(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.85 sheetProtection (Sheet Protection Options)
  let o = ws.opts.sheetProtection;
  let includeSheetProtection = false;
  Object.keys(o).forEach(k => {
    if (o[k] !== null) {
      includeSheetProtection = true;
    }
  });

  if (includeSheetProtection) {
    // Set required fields with defaults if not specified
    o.sheet = o.sheet !== null ? o.sheet : true;
    o.objects = o.objects !== null ? o.objects : true;
    o.scenarios = o.scenarios !== null ? o.scenarios : true;

    let ele = xml.ele('sheetProtection');
    Object.keys(o).forEach(k => {
      if (o[k] !== null) {
        if (k === 'password') {
          ele.att('password', getHashOfPassword(o[k]));
        } else {
          ele.att(k, boolToInt(o[k]));
        }
      }
    });
    ele.up();
  }
}
