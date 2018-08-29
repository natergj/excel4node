import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';

export default function addSheetViews(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.87 sheetViews (Sheet Views)
  let o = ws.opts.sheetView;
  let ele = xml.ele('sheetViews');
  let sv = ele
    .ele('sheetView')
    .att('showGridLines', o.showGridLines)
    .att('tabSelected', o.tabSelected)
    .att('workbookViewId', o.workbookViewId)
    .att('rightToLeft', o.rightToLeft)
    .att('zoomScale', o.zoomScale)
    .att('zoomScaleNormal', o.zoomScaleNormal)
    .att('zoomScalePageLayoutView', o.zoomScalePageLayoutView);

  let modifiedPaneParams = [];
  Object.keys(o.pane).forEach(k => {
    if (o.pane[k] !== null) {
      modifiedPaneParams.push(k);
    }
  });
  if (modifiedPaneParams.length > 0) {
    let pEle = sv.ele('pane');
    o.pane.xSplit !== null ? pEle.att('xSplit', o.pane.xSplit) : null;
    o.pane.ySplit !== null ? pEle.att('ySplit', o.pane.ySplit) : null;
    o.pane.topLeftCell !== null ? pEle.att('topLeftCell', o.pane.topLeftCell) : null;
    o.pane.activePane !== null ? pEle.att('activePane', o.pane.activePane) : null;
    o.pane.state !== null ? pEle.att('state', o.pane.state) : null;
    pEle.up();
  }
  sv.up();
  ele.up();
}
