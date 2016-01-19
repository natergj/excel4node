module.exports = columnAccessor;

// -----------------------------------------------------------------------------

function columnAccessor(colNum) {
    var thisWS = this;
    if (!thisWS.cols) {
        thisWS.cols = {};
    }
    if (!thisWS.cols[colNum]) {
        var newCol = new Column();
        newCol.setAttribute('max', colNum);
        newCol.setAttribute('min', colNum);
        newCol.setAttribute('customWidth', 0);
        newCol.setAttribute('width', thisWS.wb.defaults.colWidth);
        thisWS.cols[colNum] = newCol;
    }
    var thisCol = thisWS.cols[colNum];
    thisCol.ws = thisWS;
    return thisCol;
}

// -----------------------------------------------------------------------------

function Column() {
    return this;
}

// Methods ---------------------------------------------------------------------

Column.prototype.setAttribute = function (attr, val) {
    this[attr] = val;
};

Column.prototype.Hide = function () {
    this.setAttribute('hidden', 1);
    return this;
};

Column.prototype.Group = function (level, isHidden) {
    this.ws.hasGroupings = true;
    var hidden = isHidden ? 1 : 0;
    this.setAttribute('outlineLevel', level);
    this.setAttribute('hidden', hidden);
    return this;
};

Column.prototype.Width = function (w) {
    this.setAttribute('width', w);
    this.setAttribute('customWidth', 1);
    return this;
};

Column.prototype.Freeze = function (scrollTo) {
    var sTo = scrollTo ? scrollTo : this.min;
    var sv = this.ws.sheet.sheetViews[0].sheetView;
    var pane;
    var foundPane = false;

    sv.forEach(function (v, i) {
        if (Object.keys(v).indexOf('pane') >= 0) {
            pane = sv[i].pane;
            foundPane = true;
        }
    });

    if (!foundPane) {
        var l = sv.push({
            pane: {
                '@activePane': 'topRight',
                '@state': 'frozen',
                '@topLeftCell': sTo.toExcelAlpha() + '1',
                '@xSplit': this.min - 1
            }
        });
        pane = sv[l - 1].pane;
    } else {
        var curTopLeft = pane['@topLeftCell'];
        var points = curTopLeft.toExcelRowCol();
        pane['@activePane'] = 'bottomRight';
        pane['@topLeftCell'] = sTo.toExcelAlpha() + points.row;
        pane['@xSplit'] = this.min - 1;
    }

    return this;
};
