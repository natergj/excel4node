const xml = require('xmlbuilder');
const utils = require('../utils.js');
const types = require('../types/index.js');
const hyperlinks = require('./classes/hyperlink');
const Picture = require('../drawing/picture.js');

let _addSheetPr = (promiseObj) => {
    // §18.3.1.82 sheetPr (Sheet Properties)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts;

        // Check if any option that would require the sheetPr element to be added exists
        if (
            o.pageSetup.fitToHeight !== null ||
            o.pageSetup.fitToWidth !== null ||
            o.outline.summaryBelow !== null ||
            o.outline.summaryRight !== null ||
            o.autoFilter.ref !== null
        ) {
            let ele = promiseObj.xml.ele('sheetPr');

            if (o.autoFilter.ref) {
                ele.att('enableFormatConditionsCalculation', 1);
                ele.att('filterMode', 1);
            }

            if (o.outline.summaryBelow !== null || o.outline.summaryRight !== null) {
                let outlineEle = ele.ele('outlinePr');
                outlineEle.att('applyStyles', 1);
                outlineEle.att('summaryBelow',  o.outline.summaryBelow === true ? 1 : 0);
                outlineEle.att('summaryRight',  o.outline.summaryRight === true ? 1 : 0);
                outlineEle.up();
            }

            // §18.3.1.65 pageSetUpPr (Page Setup Properties)
            if (o.pageSetup.fitToHeight !== null || o.pageSetup.fitToWidth !== null) {
                ele.ele('pageSetUpPr').att('fitToPage', 1).up();
            }
            ele.up();
        }

        resolve(promiseObj);
    });
};

let _addDimension = (promiseObj) => {
    // §18.3.1.35 dimension (Worksheet Dimensions)
    return new Promise((resolve, reject) => {
        let firstCell = 'A1';
        let lastCell = `${utils.getExcelAlpha(promiseObj.ws.lastUsedCol)}${promiseObj.ws.lastUsedRow}`;
        let ele = promiseObj.xml.ele('dimension');
        ele.att('ref', `${firstCell}:${lastCell}`);
        ele.up();

        resolve(promiseObj);
    });
};

let _addSheetViews = (promiseObj) => {
    // §18.3.1.88 sheetViews (Sheet Views)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.sheetView;
        let ele = promiseObj.xml.ele('sheetViews');
        let sv = ele.ele('sheetView')
        .att('showGridLines', o.showGridLines)
        .att('workbookViewId', o.workbookViewId)
        .att('rightToLeft', o.rightToLeft)
        .att('zoomScale', o.zoomScale)
        .att('zoomScaleNormal', o.zoomScaleNormal)
        .att('zoomScalePageLayoutView', o.zoomScalePageLayoutView);

        let modifiedPaneParams = [];
        Object.keys(o.pane).forEach((k) => {
            if (o.pane[k] !== null) {
                modifiedPaneParams.push(k);
            }
        });
        if (modifiedPaneParams.length > 0) {
            let pEle = sv.ele('pane');
            o.pane.xSplit !== null ? pEle.att('xSplit', o.pane.xSplit) : null;
            o.pane.ySplit !== null ? pEle.att('ySplit', o.pane.ySplit) : null;
            o.pane.topLeftCell !== null ? pEle.att('topLeftCell', o.pane.topLeftCell) : null;
            o.pane.activePane !== null ? pEle.att('activePane', o.pane.activePane) : null;
            o.pane.state !== null ? pEle.att('state', o.pane.state) : null;
            pEle.up();
        }
        sv.up();
        ele.up();
        resolve(promiseObj);
    });
};

let _addSheetFormatPr = (promiseObj) => {
    // §18.3.1.81 sheetFormatPr (Sheet Format Properties)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.sheetFormat;
        let ele = promiseObj.xml.ele('sheetFormatPr');

        o.baseColWidth !== null ? ele.att('baseColWidth', o.baseColWidth) : null;
        o.defaultColWidth !== null ? ele.att('defaultColWidth', o.defaultColWidth) : null;
        o.defaultRowHeight !== null ? ele.att('defaultRowHeight', o.defaultRowHeight) : ele.att('defaultRowHeight', 16);
        o.thickBottom !== null ? ele.att('thickBottom', utils.boolToInt(o.thickBottom)) : null;
        o.thickTop !== null ? ele.att('thickTop', utils.boolToInt(o.thickTop)) : null;


        if (typeof o.defaultRowHeight === 'number') {
            ele.att('customHeight', '1');
        }
        ele.up();
        resolve(promiseObj);
    });
};

let _addCols = (promiseObj) => {
    // §18.3.1.17 cols (Column Information)
    return new Promise((resolve, reject) => {

        if (promiseObj.ws.columnCount > 0) {
            let colsEle = promiseObj.xml.ele('cols');

            for (let colId in promiseObj.ws.cols) {
                let col = promiseObj.ws.cols[colId];
                let colEle = colsEle.ele('col');

                col.min !== null ? colEle.att('min', col.min) : null;
                col.max !== null ? colEle.att('max', col.max) : null;
                col.width !== null ? colEle.att('width', col.width) : null;
                col.style !== null ? colEle.att('style', col.style) : null;
                col.hidden !== null ? colEle.att('hidden', utils.boolToInt(col.hidden)) : null;
                col.customWidth !== null ? colEle.att('customWidth', utils.boolToInt(col.customWidth)) : null;
                col.outlineLevel !== null ? colEle.att('outlineLevel', col.outlineLevel) : null;
                col.collapsed !== null ? colEle.att('collapsed', utils.boolToInt(col.collapsed)) : null;
                colEle.up();
            }
            colsEle.up();
        }

        resolve(promiseObj);
    });
};

let _addSheetData = (promiseObj) => {
    // §18.3.1.80 sheetData (Sheet Data)
    return new Promise((resolve, reject) => {

        let ele = promiseObj.xml.ele('sheetData');
        let rows = Object.keys(promiseObj.ws.rows);

        let processRows = (theseRows) => {
            for (var r = 0; r < theseRows.length; r++) {
                let thisRow = promiseObj.ws.rows[theseRows[r]];
                thisRow.cellRefs.sort(utils.sortCellRefs);

                let rEle = ele.ele('row');

                rEle.att('r', thisRow.r);
                if (promiseObj.ws.opts.disableRowSpansOptimization !== true && thisRow.spans) {
                    rEle.att('spans', thisRow.spans);
                }
                thisRow.s !== null ? rEle.att('s', thisRow.s) : null;
                thisRow.customFormat !== null ? rEle.att('customFormat', thisRow.customFormat) : null;
                thisRow.ht !== null ? rEle.att('ht', thisRow.ht) : null;
                thisRow.hidden !== null ? rEle.att('hidden', thisRow.hidden) : null;
                thisRow.customHeight === true || typeof promiseObj.ws.opts.sheetFormat.defaultRowHeight === 'number' ? rEle.att('customHeight', 1) : null;
                thisRow.outlineLevel !== null ? rEle.att('outlineLevel', thisRow.outlineLevel) : null;
                thisRow.collapsed !== null ? rEle.att('collapsed', thisRow.collapsed) : null;
                thisRow.thickTop !== null ? rEle.att('thickTop', thisRow.thickTop) : null;
                thisRow.thickBot !== null ? rEle.att('thickBot', thisRow.thickBot) : null;

                for (var i = 0; i < thisRow.cellRefs.length; i++) {
                    promiseObj.ws.cells[thisRow.cellRefs[i]].addToXMLele(rEle);
                }

                rEle.up();
            }

            processNextRows();
        };

        let processNextRows = () => {
            let theseRows = rows.splice(0, 500);
            if (theseRows.length === 0) {
                ele.up();
                return resolve(promiseObj);
            }
            processRows(theseRows);
        };

        processNextRows();

    });
};

let _addSheetProtection = (promiseObj) => {
    // §18.3.1.85 sheetProtection (Sheet Protection Options)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.sheetProtection;
        let includeSheetProtection = false;
        Object.keys(o).forEach((k) =>  {
            if (o[k] !== null) {
                includeSheetProtection = true;
            }
        });

        if (includeSheetProtection) {
            // Set required fields with defaults if not specified
            o.sheet = o.sheet !== null ? o.sheet : true;
            o.objects = o.objects !== null ? o.objects : true;
            o.scenarios = o.scenarios !== null ? o.scenarios : true;

            let ele = promiseObj.xml.ele('sheetProtection');
            Object.keys(o).forEach((k) => {
                if (o[k] !== null) {
                    if (k === 'password') {
                        ele.att('password', utils.getHashOfPassword(o[k]));
                    } else {
                        ele.att(k, utils.boolToInt(o[k]));
                    }
                }
            });
            ele.up();
        }
        resolve(promiseObj);
    });
};

let _addAutoFilter = (promiseObj) => {
    // §18.3.1.2 autoFilter (AutoFilter Settings)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.autoFilter;

        if (typeof o.startRow === 'number') {
            let ele = promiseObj.xml.ele('autoFilter');
            let filterRow = promiseObj.ws.rows[o.startRow];

            o.startCol = typeof o.startCol === 'number' ? o.startCol : null;
            o.endCol = typeof o.endCol === 'number' ? o.endCol : null;

            if (typeof o.endRow !== 'number') {
                let firstEmptyRow = undefined;
                let curRow = o.startRow;
                while (firstEmptyRow === undefined) {
                    if (!promiseObj.ws.rows[curRow]) {
                        firstEmptyRow = curRow;
                    } else {
                        curRow++;
                    }
                }

                o.endRow = firstEmptyRow - 1;
            }

            // Columns to sort not manually set. filter all columns in this row containing data.
            if (typeof o.startCol !== 'number' || typeof o.endCol !== 'number') {
                o.startCol = filterRow.firstColumn;
                o.endCol = filterRow.lastColumn;
            }

            let startCell = utils.getExcelAlpha(o.startCol) + o.startRow;
            let endCell = utils.getExcelAlpha(o.endCol) + o.endRow;

            ele.att('ref', `${startCell}:${endCell}`);
            promiseObj.ws.wb.definedNameCollection.addDefinedName({
                hidden: 1,
                localSheetId: promiseObj.ws.localSheetId,
                name: '_xlnm._FilterDatabase',
                refFormula: '\'' + promiseObj.ws.name + '\'!' +
                    '$' + utils.getExcelAlpha(o.startCol) +
                    '$' + o.startRow +
                    ':' +
                    '$' + utils.getExcelAlpha(o.endCol) +
                    '$' + o.endRow
            });
            ele.up();
        }
        resolve(promiseObj);
    });
};

let _addMergeCells = (promiseObj) => {
    // §18.3.1.55 mergeCells (Merge Cells)
    return new Promise((resolve, reject) => {

        if (promiseObj.ws.mergedCells instanceof Array && promiseObj.ws.mergedCells.length > 0) {
            let ele = promiseObj.xml.ele('mergeCells').att('count', promiseObj.ws.mergedCells.length);
            promiseObj.ws.mergedCells.forEach((cr) => {
                ele.ele('mergeCell').att('ref', cr).up();
            });
            ele.up();
        }

        resolve(promiseObj);
    });
};

let _addConditionalFormatting = (promiseObj) => {
    // §18.3.1.18 conditionalFormatting (Conditional Formatting)
    return new Promise((resolve, reject) => {
        promiseObj.ws.cfRulesCollection.addToXMLele(promiseObj.xml);
        resolve(promiseObj);
    });
};

let _addHyperlinks = (promiseObj) => {
    // §18.3.1.48 hyperlinks (Hyperlinks)
    return new Promise((resolve, reject) => {
        promiseObj.ws.hyperlinkCollection.addToXMLele(promiseObj.xml);
        resolve(promiseObj);
    });
};

let _addDataValidations = (promiseObj) => {
    // §18.3.1.33 dataValidations (Data Validations)
    return new Promise((resolve, reject) => {
        if (promiseObj.ws.dataValidationCollection.length > 0) {
            promiseObj.ws.dataValidationCollection.addToXMLele(promiseObj.xml);
        }
        resolve(promiseObj);
    });
};

let _addPrintOptions = (promiseObj) => {
    // §18.3.1.70 printOptions (Print Options)
    return new Promise((resolve, reject) => {

        let addPrintOptions = false;
        let o = promiseObj.ws.opts.printOptions;
        Object.keys(o).forEach((k) => {
            if (o[k] !== null) {
                addPrintOptions = true;
            }
        });

        if (addPrintOptions) {
            let poEle = promiseObj.xml.ele('printOptions');
            o.centerHorizontal === true ? poEle.att('horizontalCentered', 1) : null;
            o.centerVertical === true ? poEle.att('verticalCentered', 1) : null;
            o.printHeadings === true ? poEle.att('headings', 1) : null;
            if (o.printGridLines === true) {
                poEle.att('gridLines', 1);
                poEle.att('gridLinesSet', 1);
            }
            poEle.up();
        }

        resolve(promiseObj);
    });
};

let _addPageMargins = (promiseObj) => {
    // §18.3.1.62 pageMargins (Page Margins)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.margins;

        promiseObj.xml.ele('pageMargins')
        .att('left', o.left)
        .att('right', o.right)
        .att('top', o.top)
        .att('bottom', o.bottom)
        .att('header', o.header)
        .att('footer', o.footer)
        .up();

        resolve(promiseObj);
    });
};

let _addLegacyDrawing = (promiseObj) => {
    return new Promise((resolve, reject) => {

        const rId = promiseObj.ws.relationships.indexOf('commentsVml') + 1;
        if(rId === 0) {
            resolve(promiseObj);
        } else {
            promiseObj.xml.ele('legacyDrawing')
            .att('r:id', 'rId' + rId)
            .up();

            resolve(promiseObj);
        }
    })
}

let _addPageSetup = (promiseObj) => {
    // §18.3.1.63 pageSetup (Page Setup Settings)
    return new Promise((resolve, reject) => {

        let addPageSetup = false;
        let o = promiseObj.ws.opts.pageSetup;
        Object.keys(o).forEach((k) => {
            if (o[k] !== null) {
                addPageSetup = true;
            }
        });

        if (addPageSetup === true) {
            let psEle = promiseObj.xml.ele('pageSetup');
            o.paperSize !== null ? psEle.att('paperSize', types.paperSize[o.paperSize]) : null;
            o.paperHeight !== null ? psEle.att('paperHeight', o.paperHeight) : null;
            o.paperWidth !== null ? psEle.att('paperWidth', o.paperWidth) : null;
            o.scale !== null ? psEle.att('scale', o.scale) : null;
            o.firstPageNumber !== null ? psEle.att('firstPageNumber', o.firstPageNumber) : null;
            o.fitToWidth !== null ? psEle.att('fitToWidth', o.fitToWidth) : null;
            o.fitToHeight !== null ? psEle.att('fitToHeight', o.fitToHeight) : null;
            o.pageOrder !== null ? psEle.att('pageOrder', o.pageOrder) : null;
            o.orientation !== null ? psEle.att('orientation', o.orientation) : null;
            o.usePrinterDefaults !== null ? psEle.att('usePrinterDefaults', utils.boolToInt(o.usePrinterDefaults)) : null;
            o.blackAndWhite !== null ? psEle.att('blackAndWhite', utils.boolToInt(o.blackAndWhite)) : null;
            o.draft !== null ? psEle.att('draft', utils.boolToInt(o.draft)) : null;
            o.cellComments !== null ? psEle.att('cellComments', o.cellComments) : null;
            o.useFirstPageNumber !== null ? psEle.att('useFirstPageNumber', utils.boolToInt(o.useFirstPageNumber)) : null;
            o.errors !== null ? psEle.att('errors', o.errors) : null;
            o.horizontalDpi !== null ? psEle.att('horizontalDpi', o.horizontalDpi) : null;
            o.verticalDpi !== null ? psEle.att('verticalDpi', o.verticalDpi) : null;
            o.copies !== null ? psEle.att('copies', o.copies) : null;
            psEle.up();
        }

        resolve(promiseObj);
    });
};

let _addPageBreaks = (promiseObj) => {
    // colBreaks (§18.3.1.14); rowBreaks (§18.3.1.74)
    const rowBreaks = promiseObj.ws.pageBreaks.row;
    if (rowBreaks.length > 0) {
        const rbEle = promiseObj.xml.ele('rowBreaks');
        rbEle.att('count', rowBreaks.length);
        rbEle.att('manualBreakCount', rowBreaks.length);
        rowBreaks.forEach(pos => {
            const bEle = rbEle.ele('brk');
            bEle.att('id', pos);
            bEle.att('man', 1);
            bEle.up();
        });
        rbEle.up();
    }
    const colBreaks = promiseObj.ws.pageBreaks.column;
    if (colBreaks.length > 0) {
        const cbEle = promiseObj.xml.ele('colBreaks');
        cbEle.att('count', colBreaks.length);
        cbEle.att('manualBreakCount', colBreaks.length);
        colBreaks.forEach(pos => {
            const bEle = cbEle.ele('brk');
            bEle.att('id', pos);
            bEle.att('man', 1);
            bEle.up();
        });
        cbEle.up();
    }
    return promiseObj;
}

let _addHeaderFooter = (promiseObj) => {
    // §18.3.1.46 headerFooter (Header Footer Settings)
    return new Promise((resolve, reject) => {

        let addHeaderFooter = false;
        let o = promiseObj.ws.opts.headerFooter;
        Object.keys(o).forEach((k) => {
            if (o[k] !== null) {
                addHeaderFooter = true;
            }
        });

        if (addHeaderFooter === true) {
            let hfEle = promiseObj.xml.ele('headerFooter');

            o.alignWithMargins !== null ? hfEle.att('alignWithMargins', utils.boolToInt(o.alignWithMargins)) : null;
            o.differentFirst !== null ? hfEle.att('differentFirst', utils.boolToInt(o.differentFirst)) : null;
            o.differentOddEven !== null ? hfEle.att('differentOddEven', utils.boolToInt(o.differentOddEven)) : null;
            o.scaleWithDoc !== null ? hfEle.att('scaleWithDoc', utils.boolToInt(o.scaleWithDoc)) : null;

            o.oddHeader !== null ? hfEle.ele('oddHeader').text(o.oddHeader).up() : null;
            o.oddFooter !== null ? hfEle.ele('oddFooter').text(o.oddFooter).up() : null;
            o.evenHeader !== null ? hfEle.ele('evenHeader').text(o.evenHeader).up() : null;
            o.evenFooter !== null ? hfEle.ele('evenFooter').text(o.evenFooter).up() : null;
            o.firstHeader !== null ? hfEle.ele('firstHeader').text(o.firstHeader).up() : null;
            o.firstFooter !== null ? hfEle.ele('firstFooter').text(o.firstFooter).up() : null;
            hfEle.up();
        }

        resolve(promiseObj);
    });
};

let _addDrawing = (promiseObj) => {
    // §18.3.1.36 drawing (Drawing)
    return new Promise((resolve, reject) => {
        if (!promiseObj.ws.drawingCollection.isEmpty) {
            let dId = promiseObj.ws.relationships.indexOf('drawing') + 1;
            promiseObj.xml.ele('drawing').att('r:id', 'rId' + dId).up();
        }
        resolve(promiseObj);
    });
};

let sheetXML = (ws) => {
    return new Promise((resolve, reject) => {

        let xmlProlog = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        let xmlString = '';
        let wsXML = xml.begin({
          'allowSurrogateChars': true,
        }, (chunk) => {
            xmlString += chunk;
        })

        .ele('worksheet')
        .att('mc:Ignorable', 'x14ac')
        .att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        .att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        .att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        .att('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

        // Excel complains if specific elements on not in the correct order in the XML doc as defined in §M.2.2
        let promiseObj = { xml: wsXML, ws: ws };

        _addSheetPr(promiseObj)
        .then(_addDimension)
        .then(_addSheetViews)
        .then(_addSheetFormatPr)
        .then(_addCols)
        .then(_addSheetData)
        .then(_addSheetProtection)
        .then(_addAutoFilter)
        .then(_addMergeCells)
        .then(_addConditionalFormatting)
        .then(_addDataValidations)
        .then(_addHyperlinks)
        .then(_addPrintOptions)
        .then(_addPageMargins)
        .then(_addLegacyDrawing)
        .then(_addPageSetup)
        .then(_addPageBreaks)
        .then(_addHeaderFooter)
        .then(_addDrawing)
        .then((promiseObj) => {
            return new Promise((resolve, reject) => {
                wsXML.end();
                resolve(xmlString);
            });
        })
        .then((xml) => {
            resolve(xml);
        })
        .catch((e) => {
            throw new Error(e.stack);
        });
    });
};

let relsXML = (ws) => {
    return new Promise((resolve, reject) => {
        let sheetRelRequired = false;
        if (ws.relationships.length > 0) {
            sheetRelRequired = true;
        }

        if (sheetRelRequired === false) {
            resolve();
        }

        let relXML = xml.create(
            'Relationships',
            {
                'version': '1.0',
                'encoding': 'UTF-8',
                'standalone': true,
                'allowSurrogateChars': true
            }
        );
        relXML.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        ws.relationships.forEach((r, i) => {
            let rId = 'rId' + (i + 1);
            if (r instanceof hyperlinks.Hyperlink) {
                relXML.ele('Relationship')
                .att('Id', rId)
                .att('Target', r.location)
                .att('TargetMode', 'External')
                .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
            } else if (r === 'drawing') {
                relXML.ele('Relationship')
                .att('Id', rId)
                .att('Target', '../drawings/drawing' + ws.sheetId + '.xml')
                .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing');
            } else if (r === 'comments') {
                relXML.ele('Relationship')
                .att('Id', rId)
                .att('Target', '../comments' + ws.sheetId + '.xml')
                .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments');
            } else if (r === 'commentsVml') {
                relXML.ele('Relationship')
                .att('Id', rId)
                .att('Target', '../drawings/commentsVml' + ws.sheetId + '.vml')
                .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing');
            }
        });
        let xmlString = relXML.doc().end();
        resolve(xmlString);
    });
};

let commentsXML = (ws) => {
    return new Promise((resolve, reject) => {
        const commentsXml = xml.create(
            'comments',
            {
                'version': '1.0',
                'encoding': 'UTF-8',
                'standalone': true,
                'allowSurrogateChars': true
            }
        );
        commentsXml.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        commentsXml.ele('authors').ele('author').text(ws.wb.author);

        const commentList = commentsXml.ele('commentList');
        Object.keys(ws.comments).forEach(ref => {
            commentList
                .ele('comment')
                .att('ref', ref)
                .att('authorId', '0')
                .att('guid', ws.comments[ref].uuid)
                .ele('text')
                .ele('t')
                .text(ws.comments[ref].comment);
        });
        let xmlString = commentsXml.doc().end();
        resolve(xmlString);
    });
}

let commentsVmlXML = (ws) => {
    return new Promise((resolve, reject) => {
        // do not add XML prolog to document
        const vmlXml = xml.begin().ele('xml');
        vmlXml.att('xmlns:v', 'urn:schemas-microsoft-com:vml')
        vmlXml.att('xmlns:o', 'urn:schemas-microsoft-com:office:office');
        vmlXml.att('xmlns:x', 'urn:schemas-microsoft-com:office:excel');

        const sl = vmlXml.ele('o:shapelayout').att('v:ext', 'edit');
        sl.ele('o:idmap').att('v:ext', 'edit').att('data', ws.sheetId);

        const st = vmlXml.ele('v:shapetype')
            .att('id', '_x0000_t202')
            .att('coordsize', '21600,21600')
            .att('o:spt', '202')
            .att('path', 'm,l,21600r21600,l21600,xe');
        st.ele('v:stroke').att('joinstyle', 'miter');
        st.ele('v:path').att('gradientshapeok', 't').att('o:connecttype', 'rect');

        Object.keys(ws.comments).forEach((ref) => {
            const {row, col, position, marginLeft, marginTop, width, height, zIndex, visibility, fillColor} = ws.comments[ref];
            const shape = vmlXml.ele('v:shape');
            shape.att('id', `_${ws.sheetId}_${row}_${col}`);
            shape.att('type', "#_x0000_t202");
            shape.att('style', `position:${position};margin-left:${marginLeft};margin-top:${marginTop};width:${width};height:${height};z-index:${zIndex};visibility:${visibility}`);
            shape.att('fillcolor', fillColor);
            shape.att('o:insetmode', 'auto');

            shape.ele('v:path').att('o:connecttype', 'none');

            const tb = shape.ele('v:textbox').att('style', 'mso-direction-alt:auto');
            tb.ele('div').att('style', 'text-align:left');

            const cd = shape.ele('x:ClientData').att('ObjectType', 'Note');
            cd.ele('x:MoveWithCells');
            cd.ele('x:SizeWithCells');
            cd.ele('x:AutoFill').text('False');
            cd.ele('x:Row').text(row - 1);
            cd.ele('x:Column').text(col - 1);
        });


        let xmlString = vmlXml.doc().end();
        resolve(xmlString);
    });
}

module.exports = { sheetXML, relsXML, commentsXML, commentsVmlXML };
