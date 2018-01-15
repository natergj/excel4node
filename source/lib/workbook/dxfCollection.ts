import Style from "../style/style";
import { isEqual } from "lodash";

class DXFItem {
  // §18.8.14 dxf (Formatting)
  public wb;
  public style;
  public id;

  constructor(style, wb) {
    this.wb = wb;
    this.style = style;
    this.id;
  }
  get dxfId() {
    return this.id;
  }

  addToXMLele(ele) {
    this.style.addDXFtoXMLele(ele);
  }
}

export default class DXFCollection {
  // §18.8.15 dxfs (Formats)
  private wb;
  private items;
  constructor(wb) {
    this.wb = wb;
    this.items = [];
  }

  add(style) {
    if (!(style instanceof Style)) {
      style = this.wb.Style(style);
    }

    let thisItem;
    this.items.forEach(item => {
      if (isEqual(item.style.toObject(), style.toObject())) {
        return (thisItem = item);
      }
    });
    if (!thisItem) {
      thisItem = new DXFItem(style, this.wb);
      this.items.push(thisItem);
      thisItem.id = this.items.length - 1;
    }
    return thisItem;
  }

  get length() {
    return this.items.length;
  }

  addToXMLele(ele) {
    let dxfXML = ele.ele("dxfs").att("count", this.length);

    this.items.forEach(item => {
      item.addToXMLele(dxfXML);
    });
  }
}
