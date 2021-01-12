'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var Drawing = require('./drawing.js');
var path = require('path');
var imgsz = require('image-size');
var mime = require('mime');
var uniqueId = require('lodash.uniqueid');

var EMU = require('../classes/emu.js');
var xmlbuilder = require('xmlbuilder');

var headerFooterPicture = function (_Drawing) {
    _inherits(headerFooterPicture, _Drawing);

    /**
     * Element representing an Excel Picture subclass of Drawing
     * @property {String} kind Kind of picture (currently only image is supported)
     * @property {String} type ooxml schema
     * @property {String} imagePath Filesystem path to image
     * @property {Buffer} image Buffer with image
     * @property {String} contentType Mime type of image
     * @property {String} description Description of image
     * @property {String} title Title of image
     * @property {String} id ID of image
     * @property {String} noGrp pickLocks property
     * @property {String} noSelect pickLocks property
     * @property {String} noRot pickLocks property
     * @property {String} noChangeAspect pickLocks property
     * @property {String} noMove pickLocks property
     * @property {String} noResize pickLocks property
     * @property {String} noEditPoints pickLocks property
     * @property {String} noAdjustHandles pickLocks property
     * @property {String} noChangeArrowheads pickLocks property
     * @property {String} noChangeShapeType pickLocks property
     * @returns {Picture} Excel Picture  pickLocks property
     */
    function headerFooterPicture(opts) {
        _classCallCheck(this, headerFooterPicture);

        var _this = _possibleConstructorReturn(this, (headerFooterPicture.__proto__ || Object.getPrototypeOf(headerFooterPicture)).call(this));

        _this.kind = 'image';
        _this.type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
        _this.imagePath = opts.path;
        _this.image = opts.image;

        _this._name = _this.image ? opts.name || uniqueId('image-') : opts.name || path.basename(_this.imagePath);

        var size = imgsz(_this.imagePath || _this.image);

        _this._pxWidth = size.width;
        _this._pxHeight = size.height;

        _this._extension = _this.image ? size.type : path.extname(_this.imagePath).substr(1);

        _this.contentType = mime.getType(_this._extension);
        _this.scale = opts.scale;

        _this._descr = null;
        _this._title = null;
        _this._id;
        // picLocks §20.1.2.2.31 picLocks (Picture Locks)
        _this.noGrp;
        _this.noSelect;
        _this.noRot;
        _this.noChangeAspect = true;
        _this.noMove;
        _this.noResize;
        _this.noEditPoints;
        _this.noAdjustHandles;
        _this.noChangeArrowheads;
        _this.noChangeShapeType;
        return _this;
    }

    _createClass(headerFooterPicture, [{
        key: 'addToXMLele',


        /**
         * @alias Picture.addToXMLele
         * @desc When generating Workbook output, attaches pictures to the drawings xml file
         * @func Picture.addToXMLele
         * @param {xmlbuilder.Element} xmlele Element object of the xmlbuilder module
         */
        value: function addToXMLele(xmlele) {
            var headerFooterPositions = ['LF', 'CF', 'RF', 'LH', 'CH', 'RH'];
            var ind = headerFooterPositions.indexOf(this.position);

            //Size calculations see: https://stackoverflow.com/a/14369133/14456373, http://lcorneliussen.de/raw/dashboards/ooxml/
            var _scale = 1.0;if (this.scale) _scale = this.scale;
            var x_width = Math.round(this.width * _scale / 12700);
            var x_height = Math.round(this.height * _scale / 12700);

            var sh = xmlele.ele('v:shape');
            sh.att('id', this.position).att('o:spid', '_x0000_s' + 1025 + ind).att('type', '#_x0000_t75').att('style', 'position:absolute;margin-left:0;margin-top:0;width:' + x_width + ';height:' + x_height + ';z-index:' + (ind + 1));
            sh.ele('v:imagedata').att('o:relid', this.rId).att('o:title', 'image' + this.id);
            sh.ele('o:lock').att('v:ext', 'edit').att('rotation', 't');
        }
    }, {
        key: 'name',
        get: function get() {
            return this._name;
        },
        set: function set(newName) {
            this._name = newName;
        }
    }, {
        key: 'id',
        get: function get() {
            return this._id;
        },
        set: function set(id) {
            this._id = id;
        }
    }, {
        key: 'rId',
        get: function get() {
            return 'rId' + this._id;
        }
    }, {
        key: 'description',
        get: function get() {
            return this._descr !== null ? this._descr : this._name;
        },
        set: function set(desc) {
            this._descr = desc;
        }
    }, {
        key: 'title',
        get: function get() {
            return this._title !== null ? this._title : this._name;
        },
        set: function set(title) {
            this._title = title;
        }
    }, {
        key: 'extension',
        get: function get() {
            return this._extension;
        }
    }, {
        key: 'width',
        get: function get() {
            var inWidth = this._pxWidth / 96;
            var emu = new EMU(inWidth + 'in');
            return emu.value;
        }
    }, {
        key: 'height',
        get: function get() {
            var inHeight = this._pxHeight / 96;
            var emu = new EMU(inHeight + 'in');
            return emu.value;
        }
    }]);

    return headerFooterPicture;
}(Drawing);

module.exports = headerFooterPicture;
//# sourceMappingURL=headerFooterPicture.js.map