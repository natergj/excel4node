import { Border, IBorderOptions } from './border';

export class BorderCollection {
  borders: Map<string, Border>;

  constructor() {
    this.borders = new Map();
  }

  get count() {
    return this.borders.size;
  }

  addBorder(border: Border | IBorderOptions) {
    if (border instanceof Border) {
      this.borders.set(border.hash, border);
    } else {
      const newBorder = new Border(border);
      this.borders.set(newBorder.hash, newBorder);
    }
  }
}
