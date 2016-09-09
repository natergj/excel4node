'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol ? "symbol" : typeof obj; };

var xmlbuilder = require('xmlbuilder');
var JSZip = require('jszip');
var fs = require('fs');
var CTColor = require('../style/classes/ctColor.js');

var addRootContentTypesXML = function addRootContentTypesXML(promiseObj) {
    // Required as stated in §12.2
    return new Promise(function (resolve, reject) {
        var xml = xmlbuilder.create('Types', {
            'version': '1.0',
            'encoding': 'UTF-8',
            'standalone': true
        }).att('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        var contentTypesAdded = [];
        var extensionsAdded = [];
        promiseObj.wb.sheets.forEach(function (s, i) {
            if (s.drawingCollection.length > 0) {
                s.drawingCollection.drawings.forEach(function (d) {
                    if (extensionsAdded.indexOf(d.extension) < 0) {
                        var typeRef = d.contentType + '.' + d.extension;
                        if (contentTypesAdded.indexOf(typeRef) < 0) {
                            xml.ele('Default').att('ContentType', d.contentType).att('Extension', d.extension);
                        }
                        extensionsAdded.push(d.extension);
                    }
                });
            }
        });
        xml.ele('Default').att('ContentType', 'application/xml').att('Extension', 'xml');
        xml.ele('Default').att('ContentType', 'application/vnd.openxmlformats-package.relationships+xml').att('Extension', 'rels');
        xml.ele('Override').att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml').att('PartName', '/xl/workbook.xml');
        promiseObj.wb.sheets.forEach(function (s, i) {
            xml.ele('Override').att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml').att('PartName', '/xl/worksheets/sheet' + (i + 1) + '.xml');

            if (s.drawingCollection.length > 0) {
                xml.ele('Override').att('ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml').att('PartName', '/xl/drawings/drawing' + s.sheetId + '.xml');
            }
        });
        xml.ele('Override').att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml').att('PartName', '/xl/styles.xml');
        xml.ele('Override').att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml').att('PartName', '/xl/sharedStrings.xml');

        var xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.file('[Content_Types].xml', xmlString);
        resolve(promiseObj);
    });
};

var addRootRelsXML = function addRootRelsXML(promiseObj) {
    // Required as stated in §12.2
    return new Promise(function (resolve, reject) {
        var xml = xmlbuilder.create('Relationships', {
            'version': '1.0',
            'encoding': 'UTF-8',
            'standalone': true
        }).att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        xml.ele('Relationship').att('Id', 'rId1').att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument').att('Target', 'xl/workbook.xml');

        var xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('_rels').file('.rels', xmlString);
        resolve(promiseObj);
    });
};

var addWorkBookXML = function addWorkBookXML(promiseObj) {
    // Required as stated in §12.2
    return new Promise(function (resolve, reject) {

        var xml = xmlbuilder.create('workbook', {
            'version': '1.0',
            'encoding': 'UTF-8',
            'standalone': true
        });
        xml.att('mc:Ignorable', 'x15');
        xml.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        xml.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        xml.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        xml.att('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');

        var booksViewEle = xml.ele('bookViews');
        booksViewEle.ele('workbookView').att('xWindow', '240').att('yWindow', '15').att('windowWidth', '8505').att('windowHeight', '6240');

        var sheetsEle = xml.ele('sheets');
        promiseObj.wb.sheets.forEach(function (s, i) {
            sheetsEle.ele('sheet').att('name', s.name).att('sheetId', i + 1).att('r:id', 'rId' + (i + 1));
        });

        if (!promiseObj.wb.definedNameCollection.isEmpty) {
            promiseObj.wb.definedNameCollection.addToXMLele(xml);
        }

        var xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').file('workbook.xml', xmlString);
        resolve(promiseObj);
    });
};

var addWorkBookRelsXML = function addWorkBookRelsXML(promiseObj) {
    // Required as stated in §12.2
    return new Promise(function (resolve, reject) {

        var xml = xmlbuilder.create('Relationships', {
            'version': '1.0',
            'encoding': 'UTF-8',
            'standalone': true
        }).att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        xml.ele('Relationship').att('Id', 'rId' + (promiseObj.wb.sheets.length + 1)).att('Target', 'sharedStrings.xml').att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings');

        xml.ele('Relationship').att('Id', 'rId' + (promiseObj.wb.sheets.length + 2)).att('Target', 'styles.xml').att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');

        promiseObj.wb.sheets.forEach(function (s, i) {
            xml.ele('Relationship').att('Id', 'rId' + (i + 1)).att('Target', 'worksheets/sheet' + (i + 1) + '.xml').att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
        });

        var xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').folder('_rels').file('workbook.xml.rels', xmlString);
        resolve(promiseObj);
    });
};

var addWorkSheetsXML = function addWorkSheetsXML(promiseObj) {
    // Required as stated in §12.2
    return new Promise(function (resolve, reject) {

        var curSheet = 0;

        var processNextSheet = function processNextSheet() {
            var thisSheet = promiseObj.wb.sheets[curSheet];
            if (thisSheet) {
                curSheet++;
                thisSheet.generateXML().then(function (xml) {
                    return new Promise(function (resolve) {
                        // Add worksheet to zip
                        promiseObj.xlsx.folder('xl').folder('worksheets').file('sheet' + curSheet + '.xml', xml);

                        resolve();
                    });
                }).then(function () {
                    return thisSheet.generateRelsXML();
                }).then(function (xml) {
                    return new Promise(function (resolve) {
                        if (xml) {
                            promiseObj.xlsx.folder('xl').folder('worksheets').folder('_rels').file('sheet' + curSheet + '.xml.rels', xml);
                        }
                        resolve();
                    });
                }).then(processNextSheet).catch(function (e) {
                    promiseObj.wb.logger.error(e.stack);
                });
            } else {
                resolve(promiseObj);
            }
        };
        processNextSheet();
    });
};

/**
 * Generate XML for SharedStrings.xml file and add it to zip file. Called from _writeToBuffer()
 * @private
 * @memberof WorkBook
 * @param {Object} promiseObj object containing jszip instance, workbook intance and xmlvars
 * @return {Promise} Resolves with promiseObj
 */
var addSharedStringsXML = function addSharedStringsXML(promiseObj) {
    // §12.3.15 Shared String Table Part
    return new Promise(function (resolve, reject) {

        var xml = xmlbuilder.create('sst', {
            'version': '1.0',
            'encoding': 'UTF-8',
            'standalone': true
        }).att('count', promiseObj.wb.sharedStrings.length).att('uniqueCount', promiseObj.wb.sharedStrings.length).att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        promiseObj.wb.sharedStrings.forEach(function (s) {
            if (typeof s === 'string') {
                xml.ele('si').ele('t').txt(s);
            } else if (s instanceof Array) {
                (function () {

                    var thisSI = xml.ele('si');
                    var theseRuns = []; // §18.4.4 r (Rich Text Run)
                    var currProps = {};
                    var curRun = void 0;
                    var i = 0;
                    while (i < s.length) {
                        if (typeof s[i] === 'string') {
                            if (curRun === undefined) {
                                theseRuns.push({
                                    props: {},
                                    text: ''
                                });
                                curRun = theseRuns[theseRuns.length - 1];
                            }
                            curRun.text = curRun.text + s[i];
                        } else if (_typeof(s[i]) === 'object') {
                            theseRuns.push({
                                props: {},
                                text: ''
                            });
                            curRun = theseRuns[theseRuns.length - 1];
                            Object.keys(s[i]).forEach(function (k) {
                                currProps[k] = s[i][k];
                            });
                            Object.keys(currProps).forEach(function (k) {
                                curRun.props[k] = currProps[k];
                            });
                            if (s[i].value !== undefined) {
                                curRun.text = s[i].value;
                            }
                        }
                        i++;
                    }

                    theseRuns.forEach(function (run) {
                        if (Object.keys(run).length < 1) {
                            thisSI.ele('t', run.text).att('xml:space', 'preserve');
                        } else {
                            var thisRun = thisSI.ele('r');
                            var thisRunProps = thisRun.ele('rPr');
                            typeof run.props.name === 'string' ? thisRunProps.ele('rFont').att('val', run.props.name) : null;
                            run.props.bold === true ? thisRunProps.ele('b') : null;
                            run.props.italics === true ? thisRunProps.ele('i') : null;
                            run.props.strike === true ? thisRunProps.ele('strike') : null;
                            run.props.outline === true ? thisRunProps.ele('outline') : null;
                            run.props.shadow === true ? thisRunProps.ele('shadow') : null;
                            run.props.condense === true ? thisRunProps.ele('condense') : null;
                            run.props.extend === true ? thisRunProps.ele('extend') : null;
                            if (typeof run.props.color === 'string') {
                                var thisColor = new CTColor(run.props.color);
                                thisColor.addToXMLele(thisRunProps);
                            }
                            typeof run.props.size === 'number' ? thisRunProps.ele('sz').att('val', run.props.size) : null;
                            run.props.underline === true ? thisRunProps.ele('u') : null;
                            typeof run.props.vertAlign === 'string' ? thisRunProps.ele('vertAlign').att('val', run.props.vertAlign) : null;
                            thisRun.ele('t', run.text).att('xml:space', 'preserve');
                        }
                    });
                })();
            }
        });

        var xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').file('sharedStrings.xml', xmlString);

        resolve(promiseObj);
    });
};

var addStylesXML = function addStylesXML(promiseObj) {
    // §12.3.20 Styles Part
    return new Promise(function (resolve, reject) {

        var xml = xmlbuilder.create('styleSheet', {
            'version': '1.0',
            'encoding': 'UTF-8',
            'standalone': true
        }).att('mc:Ignorable', 'x14ac').att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main').att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006').att('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

        if (promiseObj.wb.styleData.numFmts.length > 0) {
            (function () {
                var nfXML = xml.ele('numFmts').att('count', promiseObj.wb.styleData.numFmts.length);
                promiseObj.wb.styleData.numFmts.forEach(function (nf) {
                    nf.addToXMLele(nfXML);
                });
            })();
        }

        var fontXML = xml.ele('fonts').att('count', promiseObj.wb.styleData.fonts.length);
        promiseObj.wb.styleData.fonts.forEach(function (f) {
            f.addToXMLele(fontXML);
        });

        var fillXML = xml.ele('fills').att('count', promiseObj.wb.styleData.fills.length);
        promiseObj.wb.styleData.fills.forEach(function (f) {
            var fXML = fillXML.ele('fill');
            f.addToXMLele(fXML);
        });

        var borderXML = xml.ele('borders').att('count', promiseObj.wb.styleData.borders.length);
        promiseObj.wb.styleData.borders.forEach(function (b) {
            b.addToXMLele(borderXML);
        });

        var cellXfsXML = xml.ele('cellXfs').att('count', promiseObj.wb.styles.length);
        promiseObj.wb.styles.forEach(function (s) {
            s.addXFtoXMLele(cellXfsXML);
        });

        if (promiseObj.wb.dxfCollection.length > 0) {
            promiseObj.wb.dxfCollection.addToXMLele(xml);
        }

        var xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').file('styles.xml', xmlString);

        resolve(promiseObj);
    });
};

var addDrawingsXML = function addDrawingsXML(promiseObj) {
    return new Promise(function (resolve) {
        if (!promiseObj.wb.mediaCollection.isEmpty) {

            promiseObj.wb.sheets.forEach(function (ws) {
                if (!ws.drawingCollection.isEmpty) {
                    (function () {

                        var drawingRelXML = xmlbuilder.create('Relationships', {
                            'version': '1.0',
                            'encoding': 'UTF-8',
                            'standalone': true
                        }).att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                        var drawingsXML = xmlbuilder.create('xdr:wsDr', {
                            'version': '1.0',
                            'encoding': 'UTF-8',
                            'standalone': true
                        });
                        drawingsXML.att('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main').att('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');

                        ws.drawingCollection.drawings.forEach(function (d) {

                            if (d.kind === 'image') {
                                var target = 'image' + d.id + '.' + d.extension;
                                promiseObj.xlsx.folder('xl').folder('media').file(target, fs.readFileSync(d.imagePath));

                                drawingRelXML.ele('Relationship').att('Id', d.rId).att('Target', '../media/' + target).att('Type', d.type);
                            }

                            d.addToXMLele(drawingsXML);
                        });

                        var drawingsXMLStr = drawingsXML.doc().end(promiseObj.xmlOutVars);
                        var drawingRelXMLStr = drawingRelXML.doc().end(promiseObj.xmlOutVars);
                        promiseObj.xlsx.folder('xl').folder('drawings').file('drawing' + ws.sheetId + '.xml', drawingsXMLStr);
                        promiseObj.xlsx.folder('xl').folder('drawings').folder('_rels').file('drawing' + ws.sheetId + '.xml.rels', drawingRelXMLStr);
                    })();
                }
            });
        }
        resolve(promiseObj);
    });
};

/**
 * Use JSZip to generate file to a node buffer
 * @private
 * @memberof WorkBook
 * @param {WorkBook} wb WorkBook instance
 * @return {Promise} resolves with Buffer 
 */
var writeToBuffer = function writeToBuffer(wb) {
    return new Promise(function (resolve, reject) {
        var promiseObj = {
            wb: wb,
            xlsx: new JSZip(),
            xmlOutVars: {}
        };

        if (promiseObj.wb.sheets.length === 0) {
            promiseObj.wb.WorkSheet();
        }

        addRootContentTypesXML(promiseObj).then(addWorkSheetsXML).then(addRootRelsXML).then(addWorkBookXML).then(addWorkBookRelsXML).then(addSharedStringsXML).then(addStylesXML).then(addDrawingsXML).then(function () {
            wb.opts.jszip.type = 'nodebuffer';
            promiseObj.xlsx.generateAsync(wb.opts.jszip).then(function (buf) {
                resolve(buf);
            }).catch(function (e) {
                reject(e);
            });
        }).catch(function (e) {
            wb.logger.error(e.stack);
            reject(e);
        });
    });
};

module.exports = { writeToBuffer: writeToBuffer };
//# sourceMappingURL=builder.js.map