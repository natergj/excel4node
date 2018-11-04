import { Font } from './font';

export class FontCollection {
  fonts: Map<string, Font>;

  constructor() {
    this.fonts = new Map();
  }

  get count() {
    return this.fonts.size;
  }

  addFont(font: Font | Partial<Font>) {
    if (font instanceof Font) {
      this.fonts.set(font.hash, font);
    } else {
      const newFont = new Font(font);
      this.fonts.set(newFont.hash, newFont);
    }
  }
}
