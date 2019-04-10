import { FontFamily, ST_FontScheme } from '../../../types';
import { getColor, EXCEL_COLOR } from '../../excelColor';
import { createHash } from 'crypto';

export class Font {
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
  scheme: ST_FontScheme;
  /**
   * Pt size of Font
   */
  size: number;
  /**
   * Font Family. defaults to roman
   */
  family: FontFamily;
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
  /**
   * Hash of font object used for lookups
   */
  hash: string;

  constructor(opts: Partial<Font> = {}) {
    // Default Font
    this.name = 'Calibri';
    this.size = 12;
    this.color = EXCEL_COLOR.black;
    this.family = 'roman';
    this.scheme = 'minor';

    if (typeof opts.color === 'string') {
      this.color = getColor(opts.color);
    }
    if (typeof opts.name === 'string') {
      this.name = opts.name;
    }
    if (typeof opts.scheme === 'string') {
      this.scheme = opts.scheme;
    }
    if (typeof opts.size === 'number') {
      this.size = opts.size;
    }
    if (typeof opts.family === 'string') {
      this.family = opts.family;
    }

    if (typeof opts.vertAlign === 'string') {
      this.vertAlign = opts.vertAlign;
    }
    if (typeof opts.charset === 'number') {
      this.charset = opts.charset;
    }

    if (typeof opts.condense === 'boolean') {
      this.condense = opts.condense;
    }
    if (typeof opts.extend === 'boolean') {
      this.extend = opts.extend;
    }
    if (typeof opts.bold === 'boolean') {
      this.bold = opts.bold;
    }
    if (typeof opts.italics === 'boolean') {
      this.italics = opts.italics;
    }
    if (typeof opts.outline === 'boolean') {
      this.outline = opts.outline;
    }
    if (typeof opts.shadow === 'boolean') {
      this.shadow = opts.shadow;
    }
    if (typeof opts.strike === 'boolean') {
      this.strike = opts.strike;
    }
    if (typeof opts.underline === 'boolean') {
      this.underline = opts.underline;
    }

    this.hash = createHash('md5')
      .update(JSON.stringify(this))
      .digest('hex');
  }

  /**
   * @alias Font.addToXMLele
   * @desc When generating Workbook output, attaches style to the styles xml file
   * @func Font.addToXMLele
   * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
   */
  addToXMLele(fontXML) {
    const fEle = fontXML.ele('font');
    if (this.condense === true) {
      fEle.ele('condense');
    }
    if (this.extend === true) {
      fEle.ele('extend');
    }
    if (this.bold === true) {
      fEle.ele('b');
    }
    if (this.italics === true) {
      fEle.ele('i');
    }
    if (this.outline === true) {
      fEle.ele('outline');
    }
    if (this.shadow === true) {
      fEle.ele('shadow');
    }
    if (this.strike === true) {
      fEle.ele('strike');
    }
    if (this.underline === true) {
      fEle.ele('u');
    }
    if (!!this.vertAlign) {
      fEle.ele('vertAlign').att('val', this.vertAlign);
    }

    fEle.ele('sz').att('val', this.size !== undefined ? this.size : 12);
    fEle.ele('color').att('rgb', this.color !== undefined ? this.color : 'FF000000');
    fEle.ele('name').att('val', this.name !== undefined ? this.name : 'Calibri');
    if (this.family !== undefined) {
      fEle.ele('family').att('val', this.family);
    }
    if (this.scheme !== undefined) {
      fEle.ele('scheme').att('val', this.scheme);
    }

    return true;
  }
}
