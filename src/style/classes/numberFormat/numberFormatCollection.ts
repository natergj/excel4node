import { NumberFormat } from './numberFormat';

export class NumberFormatCollection {
  numberFormats: Map<string, NumberFormat>;

  constructor() {
    this.numberFormats = new Map();
  }

  get count() {
    return this.numberFormats.size;
  }

  addNumberFormat(numberFormat: NumberFormat) {
    if (this.numberFormats.get(numberFormat.hash)) {
      this.numberFormats.set(numberFormat.hash, numberFormat);
    }
  }
}
