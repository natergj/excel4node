import { FONT_FAMILY } from '../../types/fontFamily';
import { getColor, EXCEL_COLOR } from '../excelColor';
import { FONT_SCHEME } from '../../types/fontScheme';

export default class Font {
  /**
   * HEX color of font
   */
  color: string;
  /**
   * Name of Font. i.e. Calibri
   */
  name: string;
  /**
   * Font Scheme. defaults to major
   */
  scheme: FONT_SCHEME;
  /**
   * Pt size of Font
   */
  size: number;
  /**
   * Font Family. defaults to roman
   */
  family: FONT_FAMILY;
  /**
   * Specifies font as subscript or superscript
   */
  vertAlign: string;
  /**
   * Character set of font as defined in §18.4.1 charset (Character Set) or standard
   */
  charset: number;
  /**
   * Macintosh compatibility settings to squeeze text together when rendering
   */
  condense: boolean;
  /**
   * Stretches out the text when rendering
   */
  extend: boolean;
  /**
   * States whether font should be bold
   */
  bold: boolean;
  /**
   * States whether font should be in italics
   */
  italics: boolean;
  /**
   * States whether font should be outlined
   */
  outline: boolean;
  /**
   * States whether font should have a shadow
   */
  shadow: boolean;
  /**
   * States whether font should have a strikethrough
   */
  strike: boolean;
  /**
   * States whether font should be underlined
   */
  underline: boolean;

  constructor(opts: Partial<Font> = {}) {
    // Default Font
    this.name = 'Calibri';
    this.size = 12;
    this.color = EXCEL_COLOR.black;
    this.family = FONT_FAMILY.roman;
    this.scheme = FONT_SCHEME.minor;

    typeof opts.color === 'string' ? (this.color = getColor(opts.color)) : null;
    typeof opts.name === 'string' ? (this.name = opts.name) : null;
    typeof opts.scheme === 'string' ? (this.scheme = opts.scheme) : null;
    typeof opts.size === 'number' ? (this.size = opts.size) : null;
    typeof opts.family === 'string' ? (this.family = opts.family) : null;

    typeof opts.vertAlign === 'string' ? (this.vertAlign = opts.vertAlign) : null;
    typeof opts.charset === 'number' ? (this.charset = opts.charset) : null;

    typeof opts.condense === 'boolean' ? (this.condense = opts.condense) : null;
    typeof opts.extend === 'boolean' ? (this.extend = opts.extend) : null;
    typeof opts.bold === 'boolean' ? (this.bold = opts.bold) : null;
    typeof opts.italics === 'boolean' ? (this.italics = opts.italics) : null;
    typeof opts.outline === 'boolean' ? (this.outline = opts.outline) : null;
    typeof opts.shadow === 'boolean' ? (this.shadow = opts.shadow) : null;
    typeof opts.strike === 'boolean' ? (this.strike = opts.strike) : null;
    typeof opts.underline === 'boolean' ? (this.underline = opts.underline) : null;
  }

  /**
   * @alias Font.addToXMLele
   * @desc When generating Workbook output, attaches style to the styles xml file
   * @func Font.addToXMLele
   * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
   */
  addToXMLele(fontXML) {
    let fEle = fontXML.ele('font');
    fEle.ele('sz').att('val', this.size !== undefined ? this.size : 12);
    fEle.ele('color').att('rgb', this.color !== undefined ? this.color : 'FF000000');
    fEle.ele('name').att('val', this.name !== undefined ? this.name : 'Calibri');
    if (this.family !== undefined) {
      fEle.ele('family').att('val', this.family);
    }
    if (this.scheme !== undefined) {
      fEle.ele('scheme').att('val', this.scheme);
    }

    this.condense === true ? fEle.ele('condense') : null;
    this.extend === true ? fEle.ele('extend') : null;
    this.bold === true ? fEle.ele('b') : null;
    this.italics === true ? fEle.ele('i') : null;
    this.outline === true ? fEle.ele('outline') : null;
    this.shadow === true ? fEle.ele('shadow') : null;
    this.strike === true ? fEle.ele('strike') : null;
    this.underline === true ? fEle.ele('u') : null;
    !!this.vertAlign ? fEle.ele('vertAlign').att('val', this.vertAlign) : null;

    return true;
  }
}
