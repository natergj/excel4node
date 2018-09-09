const Drawing = require('./drawing.js');
const path = require('path');
const imgsz = require('image-size');
const mime = require('mime');
const uniqueId = require('lodash.uniqueid');

const EMU = require('../classes/emu.js');
const xmlbuilder = require('xmlbuilder');

class Picture extends Drawing {
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
    constructor(opts) {
        super();
        this.kind = 'image';
        this.type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
        this.imagePath = opts.path;
        this.image = opts.image;

        this._name = this.image ?
            opts.name || uniqueId('image-') :
            opts.name || path.basename(this.imagePath);

        const size = imgsz(this.imagePath || this.image);

        this._pxWidth = size.width;
        this._pxHeight = size.height;

        this._extension = this.image ?
            size.type :
            path.extname(this.imagePath).substr(1);

        this.contentType = mime.getType(this._extension);

        this._descr = null;
        this._title = null;
        this._id;
        // picLocks §20.1.2.2.31 picLocks (Picture Locks)
        this.noGrp;
        this.noSelect;
        this.noRot;
        this.noChangeAspect = true;
        this.noMove;
        this.noResize;
        this.noEditPoints;
        this.noAdjustHandles;
        this.noChangeArrowheads;
        this.noChangeShapeType;
        if (['oneCellAnchor', 'twoCellAnchor'].indexOf(opts.position.type) >= 0) {
            this.anchor(opts.position.type, opts.position.from, opts.position.to);
        } else if (opts.position.type === 'absoluteAnchor') {
            this.position(opts.position.x, opts.position.y);
        } else {
            throw new TypeError('Invalid option for anchor type. anchorType must be one of oneCellAnchor, twoCellAnchor, or absoluteAnchor');
        }
    }

    get name() {
        return this._name;
    }
    set name(newName) {
        this._name = newName;
    }
    get id() {
        return this._id;
    }
    set id(id) {
        this._id = id;
    }

    get rId() {
        return 'rId' + this._id;
    }

    get description() {
        return this._descr !== null ? this._descr : this._name;
    }
    set description(desc) {
        this._descr = desc;
    }

    get title() {
        return this._title !== null ? this._title : this._name;
    }
    set title(title) {
        this._title = title;
    }

    get extension() {
        return this._extension;
    }

    get width() {
        let inWidth = this._pxWidth / 96;
        let emu = new EMU(inWidth + 'in');
        return emu.value;
    }

    get height() {
        let inHeight = this._pxHeight / 96;
        let emu = new EMU(inHeight + 'in');
        return emu.value;
    }

    /**
     * @alias Picture.addToXMLele
     * @desc When generating Workbook output, attaches pictures to the drawings xml file
     * @func Picture.addToXMLele
     * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
     */
    addToXMLele(ele) {

        let anchorEle = ele.ele('xdr:' + this.anchorType);

        if (this.editAs !== null) {
            anchorEle.att('editAs', this.editAs);
        }

        if (this.anchorType === 'absoluteAnchor') {
            anchorEle.ele('xdr:pos').att('x', this._position.x).att('y', this._position.y);
        }

        if (this.anchorType !== 'absoluteAnchor') {
            let af = this.anchorFrom;
            let afEle = anchorEle.ele('xdr:from');
            afEle.ele('xdr:col').text(af.col);
            afEle.ele('xdr:colOff').text(af.colOff);
            afEle.ele('xdr:row').text(af.row);
            afEle.ele('xdr:rowOff').text(af.rowOff);
        }

        if (this.anchorTo && this.anchorType === 'twoCellAnchor') {
            let at = this.anchorTo;
            let atEle = anchorEle.ele('xdr:to');
            atEle.ele('xdr:col').text(at.col);
            atEle.ele('xdr:colOff').text(at.colOff);
            atEle.ele('xdr:row').text(at.row);
            atEle.ele('xdr:rowOff').text(at.rowOff);
        }

        if (this.anchorType === 'oneCellAnchor' || this.anchorType === 'absoluteAnchor') {
            anchorEle.ele('xdr:ext').att('cx', this.width).att('cy', this.height);
        }

        let picEle = anchorEle.ele('xdr:pic');
        let nvPicPrEle = picEle.ele('xdr:nvPicPr');
        let cNvPrEle = nvPicPrEle.ele('xdr:cNvPr');
        cNvPrEle.att('descr', this.description);
        cNvPrEle.att('id', this.id + 1);
        cNvPrEle.att('name', this.name);
        cNvPrEle.att('title', this.title);
        let cNvPicPrEle = nvPicPrEle.ele('xdr:cNvPicPr');

        this.noGrp === true ? cNvPicPrEle.ele('a:picLocks').att('noGrp', 1) : null;
        this.noSelect === true ? cNvPicPrEle.ele('a:picLocks').att('noSelect', 1) : null;
        this.noRot === true ? cNvPicPrEle.ele('a:picLocks').att('noRot', 1) : null;
        this.noChangeAspect === true ? cNvPicPrEle.ele('a:picLocks').att('noChangeAspect', 1) : null;
        this.noMove === true ? cNvPicPrEle.ele('a:picLocks').att('noMove', 1) : null;
        this.noResize === true ? cNvPicPrEle.ele('a:picLocks').att('noResize', 1) : null;
        this.noEditPoints === true ? cNvPicPrEle.ele('a:picLocks').att('noEditPoints', 1) : null;
        this.noAdjustHandles === true ? cNvPicPrEle.ele('a:picLocks').att('noAdjustHandles', 1) : null;
        this.noChangeArrowheads === true ? cNvPicPrEle.ele('a:picLocks').att('noChangeArrowheads', 1) : null;
        this.noChangeShapeType === true ? cNvPicPrEle.ele('a:picLocks').att('noChangeShapeType', 1) : null;

        let blipFillEle = picEle.ele('xdr:blipFill');
        blipFillEle.ele('a:blip').att('r:embed', this.rId).att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        blipFillEle.ele('a:stretch').ele('a:fillRect');

        let spPrEle = picEle.ele('xdr:spPr');
        let xfrmEle = spPrEle.ele('a:xfrm');
        xfrmEle.ele('a:off').att('x', 0).att('y', 0);
        xfrmEle.ele('a:ext').att('cx', this.width).att('cy', this.height);

        let prstGeom = spPrEle.ele('a:prstGeom').att('prst', 'rect');
        prstGeom.ele('a:avLst');

        anchorEle.ele('xdr:clientData');
    }
}

module.exports = Picture;