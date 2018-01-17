import { version } from 'punycode';

class horizontalAlignment {
  private opts;
  constructor() {
    this.opts = [
      // §18.18.40 ST_HorizontalAlignment (Horizontal Alignment Type)
      'center',
      'centerContinuous',
      'distributed',
      'fill',
      'general',
      'justify',
      'left',
      'right',
    ];
    this.opts.forEach((o, i) => {
      this[o] = i + 1;
    });
  }

  validate(val) {
    if (this[val] === undefined) {
      const opts = [];
      for (const name in this) {
        if (this.hasOwnProperty(name)) {
          opts.push(name);
        }
      }
      throw new TypeError(
        `Invalid value for alignment.horizontal ${val}; Value must be one of ${this.opts.join(
          ', ',
        )}`,
      );
    } else {
      return true;
    }
  }
}

class verticalAlignment {
  private opts;
  constructor() {
    this.opts = [
      // §18.18.88 ST_VerticalAlignment (Vertical Alignment Types)
      'bottom',
      'center',
      'distributed',
      'justify',
      'top',
    ];
    this.opts.forEach((o, i) => {
      this[o] = i + 1;
    });
  }

  validate(val) {
    if (this[val] === undefined) {
      const opts = [];
      for (const name in this) {
        if (this.hasOwnProperty(name)) {
          opts.push(name);
        }
      }
      throw new TypeError(
        `Invalid value for alignment.vertical ${val}; Value must be one of ${this.opts.join(
          ', ',
        )}`,
      );
    } else {
      return true;
    }
  }
}

class readingOrdersAlignment {
  private opts;
  constructor() {
    this['contextDependent'] = 0;
    this['leftToRight'] = 1;
    this['rightToLeft'] = 2;
    this.opts = ['contextDependent', 'leftToRight', 'rightToLeft'];
  }

  validate(val) {
    if (this[val] === undefined) {
      const opts = [];
      for (const name in this) {
        if (this.hasOwnProperty(name)) {
          opts.push(name);
        }
      }
      throw new TypeError(
        `Invalid value for alignment.readingOrders ${val}; Value must be one of ${this.opts.join(
          ', ',
        )}`,
      );
    } else {
      return true;
    }
  }
}

export const horizontal = new horizontalAlignment();
export const vertical = new verticalAlignment();
export const readingOrders = new readingOrdersAlignment();
