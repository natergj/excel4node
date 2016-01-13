var DxfItem = require('./dxf_item');

module.exports = DxfCollection;

// Encapsulate all the "style differential formats" used by a Workbook.
// Since XML dxf data is referenced based on source order offset (as opposed to UID),
// we need to keep track of all such global items (and their order) here.

function DxfCollection() {
    this.items = [];
    return this;
}

DxfCollection.prototype.createFromStyle = function (style) {
    var id = this.items.length;
    var dxf = DxfItem.fromStyle(style);
    dxf.setId(id);
    this.items.push(dxf);
    return dxf;
};

DxfCollection.prototype.isEmpty = function () {
    return (this.items.length < 0);
};

DxfCollection.prototype.getContainerEl = function () {
    var itemEls = this.getBuilderElements();
    return {
        dxfs: {
            '@count': itemEls.length,
            '#list': itemEls
        }
    };
};

DxfCollection.prototype.getBuilderElements = function () {
    return this.items.map(function (item) {
        return item.toElData();
    });
};
