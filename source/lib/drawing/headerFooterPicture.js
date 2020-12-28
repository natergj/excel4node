const Drawing = require('./drawing.js');
const path = require('path');
const imgsz = require('image-size');
const mime = require('mime');
const uniqueId = require('lodash.uniqueid');

const EMU = require('../classes/emu.js');
const xmlbuilder = require('xmlbuilder');

class headerFooterPicture extends Drawing {
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
        this.scale = opts.scale;

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
     * @param {xmlbuilder.Element} xmlele Element object of the xmlbuilder module
     */
    addToXMLele(xmlele) {
        const headerFooterPositions = ['LF','CF', 'RF', 'LH', 'CH', 'RH'];
        var ind = headerFooterPositions.indexOf(this.position);
        
        //Size calculations see: https://stackoverflow.com/a/14369133/14456373, http://lcorneliussen.de/raw/dashboards/ooxml/
        let _scale = 1.0; if(this.scale) _scale = this.scale; 
        let x_width = Math.round((this.width * _scale) / 12700); 
        let x_height = Math.round((this.height * _scale) / 12700);
        

        var sh = xmlele.ele('v:shape');
        sh.att('id', this.position).att('o:spid', '_x0000_s' + 1025 + ind).att('type', '#_x0000_t75')
            .att('style', `position:absolute;margin-left:0;margin-top:0;width:${x_width};height:${x_height};z-index:${ind+1}`);
        sh.ele('v:imagedata').att('o:relid', this.rId).att('o:title', 'image' + this.id)
        sh.ele('o:lock').att('v:ext', 'edit').att('rotation', 't');
    }
}

module.exports = headerFooterPicture;