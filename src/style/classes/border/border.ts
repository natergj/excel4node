import { BorderOrdinal } from './borderOrdinal';
import { XMLElement } from 'xmlbuilder';
import { createHash } from 'crypto';

export interface IBorderOptions {
  diagonalDown?: boolean;
  diagonalUp?: boolean;
  outline?: boolean;
  bottom?: BorderOrdinal;
  diagonal?: BorderOrdinal;
  end?: BorderOrdinal;
  horizontal?: BorderOrdinal;
  start?: BorderOrdinal;
  top?: BorderOrdinal;
  vertical?: BorderOrdinal;
}

export class Border {
  diagonalDown?: boolean;
  diagonalUp?: boolean;
  outline?: boolean;
  bottom?: BorderOrdinal;
  diagonal?: BorderOrdinal;
  end?: BorderOrdinal;
  horizontal?: BorderOrdinal;
  start?: BorderOrdinal;
  top?: BorderOrdinal;
  vertical?: BorderOrdinal;
  hash: string;

  constructor(opts: IBorderOptions = {}) {
    this.diagonalDown = opts.diagonalDown;
    this.diagonalUp = opts.diagonalUp;
    this.outline = opts.outline;
    this.bottom = opts.bottom;
    this.diagonal = opts.diagonal;
    this.end = opts.end;
    this.horizontal = opts.horizontal;
    this.start = opts.start;
    this.top = opts.top;
    this.vertical = opts.vertical;

    this.hash = createHash('md5')
      .update(JSON.stringify(opts))
      .digest('hex');
  }

  addToXML(ele: XMLElement) {
    if (this.bottom) {
      const bottomEle = ele.ele('bottom');
      if (this.bottom.style) {
        bottomEle.att('style', this.bottom.style);
      }
      if (this.bottom.color) {
        this.bottom.color.addToXML(bottomEle);
      }
    }
  }
}
