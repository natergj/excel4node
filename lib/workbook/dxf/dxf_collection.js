module.exports = DxfCollection;

// Encapsulate all the "style differential formats" used by a Workbook.
// Since XML dxf data is referenced based on source order offset (as opposed to UID),
// we need to keep track of all such global items (and their order) here.

function DxfCollection() {
    this.items = [];
    return this;
}

DxfCollection.prototype.createDxf = function (props) {
    var id = this.items.length;
    props.id = id;
    this.items.push(props);
    return props;
};

DxfCollection.prototype.getBuilderElements = function () {
    return this.items.map(function (item) {
        return {
            dxf: {
                font: {
                    b: { '@val': '0' },
                    i: { '@val': '0' }
                },
                fill: {
                    patternFill: {
                        bgColor: 'FFFFC7CE'
                    }
                }
            }
        };
    });
};
