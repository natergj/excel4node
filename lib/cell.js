var Style = require('./style');

module.exports = cellAccessor;

//------------------------------------------------------------------------------

function cellAccessor(row1, col1, row2, col2, isMerged) {

    var thisWS = this;
    var theseCells = {
        cells: [],
        excelRefs: []
    };

    /******************************
        Cell Range Methods
     ******************************/
    theseCells.String = string;
    theseCells.Complex = complex;
    theseCells.Number = number;
    theseCells.Bool = bool;
    theseCells.Format = format();
    theseCells.Style = style;
    theseCells.Date = date;
    theseCells.Formula = formula;
    theseCells.Merge = mergeCells;
    theseCells.getCell = getCell;
    theseCells.Properties = propHandler;
    theseCells.Link = hyperlink;

    row2 = row2 ? row2 : row1;
    col2 = col2 ? col2 : col1;

    /******************************
        Add all cells in range to Cell definition
     ******************************/
    for (var r = row1; r <= row2; r++) {
        var thisRow = thisWS.Row(r);
        for (var c = col1; c <= col2; c++) {
            var thisCol = thisWS.Column(parseInt(c));
            if (!thisRow.cells[c]) {
                thisRow.cells[c] = new Cell(thisWS);
            }
            var thisCell = thisRow.cells[c];
            thisCell.attributes['r'] = c.toExcelAlpha() + r;
            thisRow.attributes['spans'] = '1:' + thisRow.cellCount();
            theseCells.cells.push(thisCell);
        }
    }
    theseCells.excelRefs = getAllCellsInNumericRange(row1, col1, row2, col2);
    theseCells.cells.sort(excelCellsSort);

    if (isMerged) {
        theseCells.Merge();
    }
    /******************************
        Cell Range Method Definitions
     ******************************/
    function string(val) {

        var chars, chr;
        chars = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;
        chr = val.match(chars);
        if (chr) {
            console.log('Invalid Character for XML "' + chr + '" in string "' + val + '"');
            val = val.replace(chr, '');
        }

        if (typeof(val) !== 'string') {
            console.log('Value sent to String function of cells %s was not a string, it has type of %s', JSON.stringify(theseCells.excelRefs), typeof(val));
            val = '';
        }

        val = val.toString();
        // Remove Control characters, they aren't understood by xmlbuilder
        val = val.replace(/[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/, '');

        if (!isMerged) {
            theseCells.cells.forEach(function (c, i) {
                c.String(thisWS.wb.getStringIndex(val));
            });
        } else {
            var c = theseCells.cells[0];
            c.String(thisWS.wb.getStringIndex(val));
        }
        return theseCells;
    }
    function complex(val) {
        thisWS.wb.workbook.sharedStrings.push(val);
        var index = thisWS.wb.workbook.sharedStrings.length - 1;
        if (!isMerged) {
            theseCells.cells.forEach(function (c, i) {
                c.String(index);
            });
        } else {
            var c = theseCells.cells[0];
            c.String(index);
        }
        return theseCells;
    }
    function format() {
        var methods = {
            'Number': formatter().number,
            'Date': formatter().date,
            'Font': {
                Family: formatter().font.family,
                Size: formatter().font.size,
                Bold: formatter().font.bold,
                Italics: formatter().font.italics,
                Underline: formatter().font.underline,
                Color: formatter().font.color,
                WrapText: formatter().font.wraptext,
                Alignment: {
                    Vertical: formatter().font.alignment.vertical,
                    Horizontal: formatter().font.alignment.horizontal
                }
            },
            Fill: {
                Color: formatter().fill.color,
                Pattern: formatter().fill.pattern
            },
            Border: formatter().border
        };
        return methods;
    }
    function style(sty) {

        // If style has a border, split excel cells into rows
        if (sty.xf.applyBorder > 0) {
            var cellRows = [];
            var curRow = [];
            var curCol = '';
            theseCells.excelRefs.forEach(function (cr, i) {
                var thisCol = cr.replace(/[0-9]/g, '');
                if (thisCol !== curCol) {
                    if (curRow.length > 0) {
                        cellRows.push(curRow);
                        curRow = [];
                    }
                    curCol = thisCol;
                }
                curRow.push(cr);
                if (i === theseCells.excelRefs.length - 1 && curRow.length > 0) {
                    cellRows.push(curRow);
                }
            });

            var borderEdges = {};
            borderEdges.left = cellRows[0][0].replace(/[0-9]/g, '');
            borderEdges.right = cellRows[cellRows.length - 1][0].replace(/[0-9]/g, '');
            borderEdges.top = cellRows[0][0].replace(/[a-zA-Z]/g, '');
            borderEdges.bottom = cellRows[0][cellRows[0].length - 1].replace(/[a-zA-Z]/g, '');
        }

        theseCells.cells.forEach(function (c, i) {
            if (theseCells.excelRefs.length === 1 || sty.xf.applyBorder === 0) {
                c.Style(sty.xf.xfId);
            } else {
                var curBorderId = sty.xf.borderId;
                var masterBorder = JSON.parse(JSON.stringify(c.ws.wb.styleData.borders[curBorderId]));

                var thisBorder = {};
                var cellRef = c.getAttribute('r');
                var cellCol = cellRef.replace(/[0-9]/g, '');
                var cellRow = cellRef.replace(/[a-zA-Z]/g, '');

                if (cellRow === borderEdges.top) {
                    thisBorder.top = masterBorder.top;
                }
                if (cellRow === borderEdges.bottom) {
                    thisBorder.bottom = masterBorder.bottom;
                }
                if (cellCol === borderEdges.left) {
                    thisBorder.left = masterBorder.left;
                }
                if (cellCol === borderEdges.right) {
                    thisBorder.right = masterBorder.right;
                }

                if (c.getAttribute('s') !== undefined) {
                    var curStyle = Style.getStyleById(c.ws.wb, c.getAttribute('s'));
                } else {
                    var curStyle = Style.getStyleById(c.ws.wb, sty.xf.xfId);
                }

                var newBorder = new Style.border(c.ws.wb, thisBorder);
                var curXF = JSON.parse(JSON.stringify(curStyle));
                curXF.applyBorder = 1;
                curXF.borderId = newBorder.borderId;
                var newXF = new Style.cellXfs(c.ws.wb, curXF);
                c.setAttribute('s', newXF.xfId);
            }
        });
        return theseCells;
    }
    function date(val) {
        if (!val || !val.toISOString) {
            if (val.toISOString() !== new Date(val).toISOString()) {
                val = new Date();
                console.log('Value sent to Date function of cells %s was not a date, it has type of %s', JSON.stringify(theseCells.excelRefs), typeof(val));
            }
        }
        if (!isMerged) {
            theseCells.cells.forEach(function (c, i) {
                var styleInfo = c.getStyleInfo();
                if (styleInfo.applyNumberFormat === 1) {
                    if (styleInfo.numFmt.formatCode.substr(0, 7) !== '[$-409]') {
                        console.log('Number format was already set for cell %s. It will be overridden with date format', thisCell.getAttribute('r'));
                        c.toCellRange().Format.Date();
                    }
                } else {
                    c.toCellRange().Format.Date();
                }
                c.Date(val);
            });
        } else {
            var c = theseCells.cells[0];
            var styleInfo = c.getStyleInfo();
            if (styleInfo.applyNumberFormat === 1) {
                if (styleInfo.numFmt.formatCode.substr(0, 7) !== '[$-409]') {
                    console.log('Number format was already set for cell %s. It will be overridden with date format', thisCell.getAttribute('r'));
                    c.toCellRange().Format.Date();
                }
            } else {
                c.toCellRange().Format.Date();
            }
            c.Date(val);
        }
        return theseCells;
    }
    function formula(val) {
        if (typeof(val) !== 'string') {
            console.log('Value sent to Formula function of cells %s was not a string, it has type of %s', JSON.stringify(theseCells.excelRefs), typeof(val));
            val = '';
        }
        if (!isMerged) {
            theseCells.cells.forEach(function (c, i) {
                c.Formula(val);
            });
        } else {
            var c = theseCells.cells[0];
            c.Formula(val);
        }

        return theseCells;
    }
    function number(val) {
        if (val === undefined || parseFloat(val) !== val) {
            console.log('Value sent to Number function of cells %s was not a number, it has type of %s and value of %s',
                JSON.stringify(theseCells.excelRefs),
                typeof(val),
                val
            );
            val = '';
        }
        val = parseFloat(val);

        if (!isMerged) {
            theseCells.cells.forEach(function (c, i) {
                c.Number(val);
                if (c.getAttribute('t')) {
                    c.deleteAttribute('t');
                }
            });
        } else {
            var c = theseCells.cells[0];
            c.Number(val);
            if (c.getAttribute('t')) {
                c.deleteAttribute('t');
            }
        }
        return theseCells;
    }
    function bool(val) {
        if (val === undefined || typeof (val.toString().toLowerCase() === 'true' || ((val.toString().toLowerCase() === 'false') ? false : val)) !== 'boolean') {
            console.log('Value sent to Bool function of cells %s was not a bool, it has type of %s and value of %s',
                JSON.stringify(theseCells.excelRefs),
                typeof(val),
                val
            );
            val = '';
        }
        val = val.toString().toLowerCase() === 'true';

        if (!isMerged) {
            theseCells.cells.forEach(function (c, i) {
                c.Bool(val);
            });
        } else {
            var c = theseCells.cells[0];
            c.Bool(val);
        }
        return theseCells;
    }
    function hyperlink(url, val) {
        if (!val) {
            val = url;
        }
        string(val);
        if (!thisWS.hyperlinks) {
            thisWS.hyperlinks = [];
        }
        var thisId = generateRId();

        thisWS.hyperlinks.push({ ref: theseCells.excelRefs[0], url: url, id: thisId });
        return theseCells;
    }
    function mergeCells() {
        if (!thisWS.mergeCells) {
            thisWS.mergeCells = [];
        }
        var cellRange = this.excelRefs[0] + ':' + this.excelRefs[this.excelRefs.length - 1];
        var rangeCells = this.excelRefs;
        var okToMerge = true;
        thisWS.mergeCells.forEach(function (cr) {
            // Check to see if currently merged cells contain cells in new merge request
            var curCells = getAllCellsInExcelRange(cr);
            var intersection = arrayIntersectSafe(rangeCells, curCells);
            if (intersection.length > 0) {
                okToMerge = false;
                console.log([
                    'Invalid Range for : ' + col1.toExcelAlpha() + row1 + ':' + col2.toExcelAlpha() + row2,
                    'Some cells in this range are already included in another merged cell range: ' + c['mergeCell'][0]['@ref'],
                    'The following are the intersection',
                    intersection
                ]);
            }
        });
        if (okToMerge) {
            thisWS.mergeCells.push(cellRange);
        }
    }
    function formatter() {

        var methods = {
            number: function (fmt) {
                theseCells.cells.forEach(function (c) {
                    setNumberFormat(c, 'formatCode', fmt);
                });
                return theseCells;
            },
            date: function (fmt) {
                theseCells.cells.forEach(function (c) {
                    setDateFormat(c, 'formatCode', fmt);
                });
                return theseCells;
            },
            font: {
                family: function (val) {
                    theseCells.cells.forEach(function (c) {
                        setCellFontAttribute(c, 'name', val);
                    });
                    return theseCells;
                },
                size: function (val) {
                    theseCells.cells.forEach(function (c) {
                        setCellFontAttribute(c, 'sz', val);
                    });
                    return theseCells;
                },
                bold: function () {
                    theseCells.cells.forEach(function (c) {
                        setCellFontAttribute(c, 'bold', true);
                    });
                    return theseCells;
                },
                italics: function () {
                    theseCells.cells.forEach(function (c) {
                        setCellFontAttribute(c, 'italics', true);
                    });
                    return theseCells;
                },
                underline: function () {
                    theseCells.cells.forEach(function (c) {
                        setCellFontAttribute(c, 'underline', true);
                    });
                    return theseCells;
                },
                color: function (val) {
                    theseCells.cells.forEach(function (c) {
                        setCellFontAttribute(c, 'color', Style.cleanColor(val));
                    });
                    return theseCells;
                },
                wraptext: function (val) {
                    theseCells.cells.forEach(function (c) {
                        setAlignmentAttribute(c, 'wrapText', 1);
                    });
                    return theseCells;
                },
                alignment: {
                    vertical: function (val) {
                        theseCells.cells.forEach(function (c) {
                            setAlignmentAttribute(c, 'vertical', val);
                        });
                        return theseCells;
                    },
                    horizontal: function (val) {
                        theseCells.cells.forEach(function (c) {
                            setAlignmentAttribute(c, 'horizontal', val);
                        });
                        return theseCells;
                    }
                }
            },
            fill: {
                color: function (val) {
                    theseCells.cells.forEach(function (c) {
                        setCellFill(c, 'fgColor', Style.cleanColor(val));
                    });
                    return theseCells;
                },
                pattern: function (val) {
                    theseCells.cells.forEach(function (c) {
                        setCellFill(c, 'patternType', val);
                    });
                    return theseCells;
                }
            },
            border: setCellsBorder
        };

        function setAlignmentAttribute(c, attr, val) {
            if (c.getAttribute('s') !== undefined) {
                var curStyle = Style.getStyleById(c.ws.wb, c.getAttribute('s'));
            } else {
                var curStyle = Style.getStyleById(c.ws.wb, 0);
            }

            var curXF = JSON.parse(JSON.stringify(curStyle));
            if (!curXF.alignment) {
                curXF.alignment = {};
            }
            curXF.applyAlignment = 1;
            curXF.alignment[attr] = val;

            var newXF = new Style.cellXfs(c.ws.wb, curXF);
            c.setAttribute('s', newXF.xfId);
        }

        function setNumberFormat(c, attr, val) {
            if (c.getAttribute('s') !== undefined) {
                var curStyle = Style.getStyleById(c.ws.wb, c.getAttribute('s'));
            } else {
                var curStyle = Style.getStyleById(c.ws.wb, 0);
            }

            if (curStyle.numFmtId !== 164 && curStyle.numFmtId !== 14) {
                var curNumFmt = JSON.parse(JSON.stringify(c.ws.wb.styleData.numFmts[curStyle.numFmtId - 165]));
            } else {
                var curNumFmt = {};
            }

            curNumFmt[attr] = val;
            var thisFmt = new Style.numFmt(c.ws.wb, curNumFmt);
            var curXF = JSON.parse(JSON.stringify(curStyle));
            curXF.applyNumberFormat = 1;
            curXF.numFmtId = thisFmt.numFmtId;
            var newXF = new Style.cellXfs(c.ws.wb, curXF);
            c.setAttribute('s', newXF.xfId);
        }

        function setDateFormat(c, attr, val) {
            if (c.getAttribute('s') !== undefined) {
                var curStyle = Style.getStyleById(c.ws.wb, c.getAttribute('s'));
            } else {
                var curStyle = Style.getStyleById(c.ws.wb, 0);
            }

            if (curStyle.numFmtId !== 164 && curStyle.numFmtId !== 14 && c.ws.wb.styleData.numFmts[curStyle.numFmtId - 165].formatCode.substr(0, 7) === '[$-409]') {
                var curNumFmt = JSON.parse(JSON.stringify(c.ws.wb.styleData.numFmts[curStyle.numFmtId - 165]));
            } else {
                var curNumFmt = {};
            }

            var curXF = JSON.parse(JSON.stringify(curStyle));
            curXF.applyNumberFormat = 1;
            if (val) {
                curNumFmt[attr] = val;
                var thisFmt = new Style.numFmt(c.ws.wb, curNumFmt);
                curXF.numFmtId = thisFmt.numFmtId;
            } else {
                curXF.numFmtId = 14;
            }
            var newXF = new Style.cellXfs(c.ws.wb, curXF);
            c.setAttribute('s', newXF.xfId);
        }

        function setCellsBorder(borderObj) {
            var cellRows = [];
            var curRow = [];
            var curCol = '';
            theseCells.excelRefs.forEach(function (cr, i) {
                var thisCol = cr.replace(/[0-9]/g, '');
                if (thisCol !== curCol) {
                    if (curRow.length > 0) {
                        cellRows.push(curRow);
                        curRow = [];
                    }
                    curCol = thisCol;
                }
                curRow.push(cr);
                if (i === theseCells.excelRefs.length - 1 && curRow.length > 0) {
                    cellRows.push(curRow);
                }
            });

            var borderEdges = {};
            borderEdges.left = cellRows[0][0].replace(/[0-9]/g, '');
            borderEdges.right = cellRows[cellRows.length - 1][0].replace(/[0-9]/g, '');
            borderEdges.top = cellRows[0][0].replace(/[a-zA-Z]/g, '');
            borderEdges.bottom = cellRows[0][cellRows[0].length - 1].replace(/[a-zA-Z]/g, '');

            theseCells.cells.forEach(function (c, i) {
                if (c.getAttribute('s') !== undefined) {
                    var curStyle = Style.getStyleById(c.ws.wb, c.getAttribute('s'));
                } else {
                    var curStyle = Style.getStyleById(c.ws.wb, 0);
                }

                var thisBorder = {};
                if (curStyle.applyBorder === 1) {
                    var curBorder = Style.getBorderById(c.ws.wb, curStyle.borderId);
                    thisBorder.left = curBorder.left;
                    thisBorder.right = curBorder.right;
                    thisBorder.top = curBorder.top;
                    thisBorder.bottom = curBorder.bottom;
                    thisBorder.diagonal = curBorder.diagonal;
                }
                var cellRef = c.getAttribute('r');
                var cellCol = cellRef.replace(/[0-9]/g, '');
                var cellRow = cellRef.replace(/[a-zA-Z]/g, '');

                if (cellRow === borderEdges.top) {
                    thisBorder.top = borderObj.top;
                }
                if (cellRow === borderEdges.bottom) {
                    thisBorder.bottom = borderObj.bottom;
                }
                if (cellCol === borderEdges.left) {
                    thisBorder.left = borderObj.left;
                }
                if (cellCol === borderEdges.right) {
                    thisBorder.right = borderObj.right;
                }


                var newBorder = new Style.border(c.ws.wb, thisBorder);
                var curXF = JSON.parse(JSON.stringify(curStyle));
                curXF.applyBorder = 1;
                curXF.borderId = newBorder.borderId;
                var newXF = new Style.cellXfs(c.ws.wb, curXF);
                c.setAttribute('s', newXF.xfId);
                if (!c.getValue().type) {
                    c.String('');
                }

            });
        }

        function setCellFill(c, attr, val) {
            if (c.getAttribute('s') !== undefined) {
                var curStyle = Style.getStyleById(c.ws.wb, c.getAttribute('s'));
            } else {
                var curStyle = Style.getStyleById(c.ws.wb, 0);
            }

            var curFill = JSON.parse(JSON.stringify(c.ws.wb.styleData.fills[curStyle.fillId]));

            curFill[attr] = val;
            var thisFill = new Style.fill(c.ws.wb, curFill);
            var curXF = JSON.parse(JSON.stringify(curStyle));
            curXF.applyFill = 1;
            curXF.fillId = thisFill.fillId;
            var newXF = new Style.cellXfs(c.ws.wb, curXF);
            c.setAttribute('s', newXF.xfId);
        }

        function setCellFontAttribute(c, attr, val) {
            if (c.getAttribute('s') !== undefined) {
                var curStyle = Style.getStyleById(c.ws.wb, c.getAttribute('s'));
            } else {
                var curStyle = Style.getStyleById(c.ws.wb, 0);
            }
            var curFont = JSON.parse(JSON.stringify(c.ws.wb.styleData.fonts[curStyle.fontId]));
            curFont[attr] = val;

            var thisFont = new Style.font(c.ws.wb, curFont);
            var curXF = JSON.parse(JSON.stringify(curStyle));
            curXF.applyFont = 1;
            curXF.fontId = thisFont.fontId;
            var newXF = new Style.cellXfs(c.ws.wb, curXF);
            c.setAttribute('s', newXF.xfId);
        }

        return methods;
    }
    function getCell(ref) {

        return theseCells.cells[theseCells.excelRefs.indexOf(ref)];
    }
    function propHandler() {
        var response = [];

        theseCells.cells.forEach(function (c, i) {
            response.push(c.getValue());
        });

        return response;
    }
    return theseCells;
}

// -----------------------------------------------------------------------------

function Cell(ws) {
    var thisCell = this;
    thisCell.ws = ws;
    thisCell.attributes = {};
    thisCell.children = {};
    thisCell.valueType;

    return thisCell;
}

/******************************
    Cell Methods
 ******************************/
Cell.prototype.setAttribute = setAttribute;
Cell.prototype.getAttribute = getAttribute;
Cell.prototype.deleteAttribute = deleteAttribute;
Cell.prototype.getStyleInfo = getStyleInfo;
Cell.prototype.toCellRange = getCellRange;
Cell.prototype.addChild = addChild;
Cell.prototype.deleteChild = deleteChild;
Cell.prototype.getChild = getChild;
Cell.prototype.String = string;
Cell.prototype.Complex = complex;
Cell.prototype.Number = number;
Cell.prototype.Bool = bool;
Cell.prototype.Date = date;
Cell.prototype.Formula = formula;
Cell.prototype.Style = styler;
Cell.prototype.getValue = getValue;


/******************************
    Cell Method Definitions
 ******************************/
function addChild(key, val) {

    return this.children[key] = val;
}

function getChild(key) {

    return this.children[key];
}

function deleteChild(key) {

    return delete this.children[key];
}

function setAttribute(attr, val) {

    return this.attributes[attr] = val;
}

function getAttribute(attr) {

    return this.attributes[attr];
}

function deleteAttribute(attr) {

    return delete this.attributes[attr];
}

function getStyleInfo() {

    var styleData = {};
    if (this.getAttribute('s')) {
        //UNDO var wbStyleData = this.ws.wb.styleData;
        var wbStyleData = this.ws.wb.styleData;
        var xf = wbStyleData.cellXfs[this.getAttribute('s')];
        styleData.xf = xf;
        styleData.applyAlignment = xf.applyAlignment;
        styleData.applyBorder = xf.applyBorder;
        styleData.applyNumberFormat = xf.applyNumberFormat;
        styleData.applyFill = xf.applyFill;
        styleData.applyFont = xf.applyFont;
        if (xf.applyAlignment !== 0) {
            styleData.alignment = xf.alignment;
        }
        if (xf.applyBorder !== 0) {
            styleData.border = wbStyleData.borders[xf.borderId];
        }
        if (xf.applyNumberFormat !== 0) {
            styleData.numFmt = wbStyleData.numFmts[xf.numFmtId - 165];
        }
        if (xf.applyFill !== 0) {
            styleData.fill = wbStyleData.fills[xf.fillId];
        }
        if (xf.applyFont !== 0) {
            styleData.font = wbStyleData.fonts[xf.fontId];
        }
    }
    return styleData;
}

function getCellRange() {

    //Since all formatting is done on cell ranges, convert cell to range of single cell
    var rc = this.getAttribute('r').toExcelRowCol();
    //UNDO return this.ws.Cell(rc.row,rc.col);
    return this.ws.Cell(rc.row, rc.col);
}

function string(strIndex) {

    this.setAttribute('t', 's');
    this.deleteChild('f');
    this.addChild('v', strIndex);
    this.valueType = 'string';
}

function complex(val) {
    this.deleteChild('f');
    this.addChild('v', val);
    this.valueType = 'complex';
}

function number(val) {

    this.deleteChild('f');
    this.addChild('v', val);
    this.valueType = 'number';
}

function bool(val) {

    this.setAttribute('t', 'b');
    this.deleteChild('f');
    this.addChild('v', val);
    this.valueType = 'bool';
}

function date(val) {

    this.deleteChild('f');
    val = new Date(val);
    var ts = val.getExcelTS();
    this.addChild('v', ts);
    this.valueType = 'date';
}

function formula(val) {

    this.deleteChild('v');
    this.addChild('f', val);
    this.valueType = 'formula';
}

function styler(style) {

    this.setAttribute('s', style);
}

function getAllCellsInNumericRange(row1, col1, row2, col2) {

    var response = [];
    row2 = row2 ? row2 : row1;
    col2 = col2 ? col2 : col1;
    for (var i = row1; i <= row2; i++) {
        for (var j = col1; j <= col2; j++) {
            response.push(j.toExcelAlpha() + i);
        }
    }
    return response.sort(excelRefSort);
}

function getAllCellsInExcelRange(range) {

    var cells = range.split(':');
    var cell1props = cells[0].toExcelRowCol();
    var cell2props = cells[1].toExcelRowCol();
    return getAllCellsInNumericRange(cell1props.row, cell1props.col, cell2props.row, cell2props.col);
}

function arrayIntersectSafe(a, b) {

    var ai = 0, bi = 0;
    var result = new Array();

    while (ai < a.length && bi < b.length) {
        if (a[ai] < b[bi]) {
            ai++;
        } else if (a[ai] > b[bi]) {
            bi++;
        } else {
            result.push(a[ai]);
            ai++;
            bi++;
        }
    }
    return result;
}

function excelRefSort(a, b) {
    if (a.replace(/[0-9]/g, '') === b.replace(/[0-9]/g, '')) {
        return a.replace(/[a-zA-Z]/g, '') - b.replace(/[a-zA-Z]/g, '');
    }
    return compareCharCodes(a, b);
}

function excelCellsSort(a, b) {
    var ar = a.attributes.r;
    var br = b.attributes.r;
    if (ar.replace(/[0-9]/g, '') === br.replace(/[0-9]/g, '')) {
        return ar.replace(/[a-zA-Z]/g, '') - br.replace(/[a-zA-Z]/g, '');
    }
    return compareCharCodes(ar, br);
}

function compareCharCodes(a, b) {
    var alphaOne = a.replace(/[0-9]/g, '').toUpperCase();
    var alphaTwo = b.replace(/[0-9]/g, '').toUpperCase();
    var numOne = '';
    var numTwo = '';

    for (var i = 0; i < alphaOne.length; i++) {
        numOne += alphaOne.charCodeAt(i);
    }
    for (var i = 0; i < alphaTwo.length; i++) {
        numTwo += alphaTwo.charCodeAt(i);
    }
    return Number(numOne) - Number(numTwo);
}

function getValue() {
    var obj = {};
    if (this.getChild('f')) {
        obj.value = this.getChild('f');
    }
    if (this.getChild('v')) {
        if (this.getAttribute('t') === 's') {
            obj.value = this.ws.wb.getStringFromIndex(this.getChild('v'));
        } else {
            obj.value = this.getChild('v');
        }
    }
    obj.ref = this.getAttribute('r');
    obj.type = this.valueType;

    return obj;
}

function generateRId() {
    var text = 'R';
    var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (var i = 0; i < 16; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
}
