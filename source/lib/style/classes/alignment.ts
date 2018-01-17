import * as types from '../../types/index';

export class Alignment {
  // §18.8.1 alignment (Alignment)
  private horizontal;
  private vertical;
  private readingOrder;
  private indent;
  private justifyLastLine;
  private relativeIndent;
  private shrinkToFit;
  private textRotation;
  private wrapText;

  // tslint:disable:max-line-length
  /**
   * @class Alignment
   * @param {Object} opts Properties of Alignment object
   * @param {String} opts.horizontal Horizontal Alignment property of text.
   * @param {String} opts.vertical Vertical Alignment property of text.
   * @param {String} opts.readingOrder Reading order for language of text.
   * @param {Number} opts.indent How much text should be indented. Setting indent to 1 will indent text 3 spaces
   * @param {Boolean} opts.justifyLastLine Specifies whether to justify last line of text
   * @param {Number} opts.relativeIndent Used in conditional formatting to state how much more text should be indented if rule passes
   * @param {Boolean} opts.shrinkToFit Indicates if text should be shrunk to fit into cell
   * @param {Number} opts.textRotation Number of degrees to rotate text counterclockwise
   * @param {Boolean} opts.wrapText States whether text with newline characters should wrap
   * @returns {Alignment}
   */
  // tslint:enable:max-line-length
  constructor(opts) {
    if (opts.horizontal !== undefined) {
      this.horizontal = types.alignment.horizontal.validate(opts.horizontal)
        ? opts.horizontal
        : null;
    }

    if (opts.vertical !== undefined) {
      this.vertical = types.alignment.vertical.validate(opts.vertical)
        ? opts.vertical
        : null;
    }

    if (opts.readingOrder !== undefined) {
      this.readingOrder = types.alignment.readingOrders.validate(
        opts.readingOrder,
      )
        ? opts.readingOrder
        : null;
    }

    if (opts.indent !== undefined) {
      if (
        typeof opts.indent === 'number' &&
        parseInt(opts.indent, 10) === opts.indent &&
        opts.indent > 0
      ) {
        this.indent = opts.indent;
      } else {
        throw new TypeError('alignment indent must be a positive integer.');
      }
    }

    if (opts.justifyLastLine !== undefined) {
      if (typeof opts.justifyLastLine === 'boolean') {
        this.justifyLastLine = opts.justifyLastLine;
      } else {
        throw new TypeError(
          'justifyLastLine alignment option must be of type boolean',
        );
      }
    }

    if (opts.relativeIndent !== undefined) {
      if (
        typeof opts.relativeIndent === 'number' &&
        parseInt(opts.relativeIndent, 10) === opts.relativeIndent &&
        opts.relativeIndent > 0
      ) {
        this.relativeIndent = opts.relativeIndent;
      } else {
        throw new TypeError('alignment indent must be a positive integer.');
      }
    }

    if (opts.shrinkToFit !== undefined) {
      if (typeof opts.shrinkToFit === 'boolean') {
        this.shrinkToFit = opts.shrinkToFit;
      } else {
        throw new TypeError(
          'justifyLastLine alignment option must be of type boolean',
        );
      }
    }

    if (opts.textRotation !== undefined) {
      if (
        typeof opts.textRotation === 'number' &&
        parseInt(opts.textRotation, 10) === opts.textRotation
      ) {
        this.textRotation = opts.textRotation;
      } else if (opts.textRotation !== undefined) {
        throw new TypeError('alignment indent must be an integer.');
      }
    }

    if (opts.wrapText !== undefined) {
      if (typeof opts.wrapText === 'boolean') {
        this.wrapText = opts.wrapText;
      } else {
        throw new TypeError(
          'justifyLastLine alignment option must be of type boolean',
        );
      }
    }
  }

  /**
   * @func Alignment.toObject
   * @desc Converts the Alignment instance to a javascript object
   * @returns {Object}
   */
  toObject() {
    const obj = {} as any;

    this.horizontal !== undefined ? (obj.horizontal = this.horizontal) : null;
    this.indent !== undefined ? (obj.indent = this.indent) : null;
    this.justifyLastLine !== undefined
      ? (obj.justifyLastLine = this.justifyLastLine)
      : null;
    this.readingOrder !== undefined
      ? (obj.readingOrder = this.readingOrder)
      : null;
    this.relativeIndent !== undefined
      ? (obj.relativeIndent = this.relativeIndent)
      : null;
    this.shrinkToFit !== undefined
      ? (obj.shrinkToFit = this.shrinkToFit)
      : null;
    this.textRotation !== undefined
      ? (obj.textRotation = this.textRotation)
      : null;
    this.vertical !== undefined ? (obj.vertical = this.vertical) : null;
    this.wrapText !== undefined ? (obj.wrapText = this.wrapText) : null;

    return obj;
  }

  /**
   * @alias Alignment.addToXMLele
   * @desc When generating Workbook output, attaches style to the styles xml file
   * @func Alignment.addToXMLele
   * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
   */
  addToXMLele(ele) {
    const thisEle = ele.ele('alignment');
    this.horizontal !== undefined
      ? thisEle.att('horizontal', this.horizontal)
      : null;
    this.indent !== undefined ? thisEle.att('indent', this.indent) : null;
    this.justifyLastLine === true ? thisEle.att('justifyLastLine', 1) : null;
    this.readingOrder !== undefined
      ? thisEle.att('readingOrder', this.readingOrder)
      : null;
    this.relativeIndent !== undefined
      ? thisEle.att('relativeIndent', this.relativeIndent)
      : null;
    this.shrinkToFit === true ? thisEle.att('shrinkToFit', 1) : null;
    this.textRotation !== undefined
      ? thisEle.att('textRotation', this.textRotation)
      : null;
    this.vertical !== undefined ? thisEle.att('vertical', this.vertical) : null;
    this.wrapText === true ? thisEle.att('wrapText', 1) : null;
  }
}
