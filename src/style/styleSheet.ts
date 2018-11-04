import { FontCollection } from './classes/font';
import { NumberFormatCollection } from '../classes/numberFormat';
import { BorderCollection } from '../classes/border';

// §18.8.39 styleSheet (Style Sheet)
export class StyleSheet {
  fonts: FontCollection;
  numFmts: NumberFormatCollection;
  borders: BorderCollection;

  constructor() {
    this.fonts = new FontCollection();
    this.numFmts = new NumberFormatCollection();
    this.borders = new BorderCollection();
  }
}
