var lodash = require('lodash');

module.exports = DxfItem;

function DxfItem(data) {
    this.id = null;
    this.data = data || {};
    return this;
}

// Construct a new DxfItem object from a given Style instance
// Various transformations and translations are applied,
// and only a subset of style parameters are featured at this point.

DxfItem.fromStyle = function (style) {
    var data = {};
    var fetchProp = function (path, fn) {
        var foundProp = lodash.get(style, path, null);
        if (foundProp) {
            fn(foundProp);
        }
    };
    fetchProp('xf.fontId', function (fontId) {
        // TODO introspecting font data from wb is awkward
        var fontData = style.wb.styleData.fonts[fontId];
        data['font'] = fontData.generateXMLObj().font;
    });
    fetchProp('xf.fillId', function (fillId) {
        var fillData = style.wb.styleData.fills[fillId];
        // TODO dxf needs fill color set as 'bgColor' as opposed to style default of 'fgColor'
        fillData.bgColor = fillData.fgColor;
        data['fill'] = fillData.generateXMLObj().fill;
    });
    fetchProp('xf.borderId', function (borderId) {
        var borderData = style.wb.styleData.borders[borderId];
        data['border'] = borderData.generateXMLObj().border;
    });
    return new DxfItem(data);
};


DxfItem.prototype.setId = function (id) {
    this.id = id;
    return this;
};

DxfItem.prototype.getId = function (id) {
    return this.id;
};


DxfItem.prototype.toElData = function () {
    return {
        dxf: this.data
    };
};
