import Worksheet from '../worksheet';
import { XMLElementOrXMLNode } from 'xmlbuilder';

export default function addSheetViews(xml: XMLElementOrXMLNode, ws: Worksheet) {
  // §18.3.1.87 sheetViews (Sheet Views)
  const o = ws.opts.sheetView;
  const ele = xml.ele('sheetViews');
  const sv = ele
    .ele('sheetView')
    .att('showGridLines', o.showGridLines)
    .att('tabSelected', o.tabSelected)
    .att('workbookViewId', o.workbookViewId)
    .att('rightToLeft', o.rightToLeft)
    .att('zoomScale', o.zoomScale)
    .att('zoomScaleNormal', o.zoomScaleNormal)
    .att('zoomScalePageLayoutView', o.zoomScalePageLayoutView);

  const modifiedPaneParams = [];
  Object.keys(o.pane).forEach(k => {
    if (o.pane[k] !== null) {
      modifiedPaneParams.push(k);
    }
  });
  if (modifiedPaneParams.length > 0) {
    const pEle = sv.ele('pane');
    if (o.pane.xSplit !== null) {
      pEle.att('xSplit', o.pane.xSplit);
    }
    if (o.pane.ySplit !== null) {
      pEle.att('ySplit', o.pane.ySplit);
    }
    if (o.pane.topLeftCell !== null) {
      pEle.att('topLeftCell', o.pane.topLeftCell);
    }
    if (o.pane.activePane !== null) {
      pEle.att('activePane', o.pane.activePane);
    }
    if (o.pane.state !== null) {
      pEle.att('state', o.pane.state);
    }
    pEle.up();
  }
  sv.up();
  ele.up();
}
