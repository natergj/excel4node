import { XMLElementOrXMLNode } from 'xmlbuilder';

// §18.3.1.15 color (Data Bar Color)

export class DataBarColor {
  /**
   * A boolean value indicating the color is automatic and system color dependent.
   */
  auto: boolean;

  /**
   * Indexed color value. Only used for backwards compatibility. References a color in indexedColors.
   */
  indexed: number;

  /**
   * Standard Alpha Red Green Blue color value (ARGB).
   * The possible values for this attribute are defined by the ST_UnsignedIntHex simple type (§18.18.86).
   */
  rgb: string;

  /**
   * A zero-based index into the <clrScheme> collection (§20.1.6.2), referencing a particular <sysClr> or <srgbClr> value expressed in the Theme part.
   */
  theme: number;

  /**
   * Specifies the tint value applied to the color.
   * If tint is supplied, then it is applied to the RGB value of the color to determine the final color applied.
   * The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
   */
  tint: number;

  constructor(color: Partial<DataBarColor>) {
    this.auto = color.auto;
    this.indexed = color.indexed;
    this.rgb = color.rgb;
    this.theme = color.theme;
    this.tint = color.tint;
  }

  addToXML(ele: XMLElementOrXMLNode) {
    if (this.auto) {
      ele.att('auto', this.auto);
    }
    if (this.indexed) {
      ele.att('indexed', this.indexed);
    }
    if (this.rgb) {
      ele.att('rgb', this.rgb);
    }
    if (this.theme) {
      ele.att('theme', this.theme);
    }
    if (this.tint) {
      ele.att('tint', this.tint);
    }
  }
}
