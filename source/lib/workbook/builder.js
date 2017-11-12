const xmlbuilder = require('xmlbuilder');
const JSZip = require('jszip');
const fs = require('fs');
const CTColor = require('../style/classes/ctColor.js');
const utils = require('../utils');

let addRootContentTypesXML = (promiseObj) => {
    // Required as stated in §12.2
    return new Promise ((resolve, reject) => {
        let xml = xmlbuilder.create(
            'Types',
            {
                'version': '1.0', 
                'encoding': 'UTF-8', 
                'standalone': true
            }
        )
        .att('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        let contentTypesAdded = [];
        let extensionsAdded = [];
        promiseObj.wb.sheets.forEach((s, i) => {
            if (s.drawingCollection.length > 0) { 
                s.drawingCollection.drawings.forEach((d) => {
                    if (extensionsAdded.indexOf(d.extension) < 0) {
                        let typeRef = d.contentType + '.' + d.extension;
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
        promiseObj.wb.sheets.forEach((s, i) => {
            xml.ele('Override')
            .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
            .att('PartName', `/xl/worksheets/sheet${i + 1}.xml`);

            if (s.drawingCollection.length > 0) {              
                xml.ele('Override')
                .att('ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml')
                .att('PartName', '/xl/drawings/drawing' + s.sheetId + '.xml');  
            }
        });
        xml.ele('Override').att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml').att('PartName', '/xl/styles.xml');
        xml.ele('Override').att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml').att('PartName', '/xl/sharedStrings.xml');

        let xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.file('[Content_Types].xml', xmlString);
        resolve(promiseObj);
    });
};

let addRootRelsXML = (promiseObj) => {
    // Required as stated in §12.2
    return new Promise ((resolve, reject) => {
        let xml = xmlbuilder.create(
            'Relationships',
            {
                'version': '1.0', 
                'encoding': 'UTF-8', 
                'standalone': true
            }
        )
        .att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        xml
        .ele('Relationship')
        .att('Id', 'rId1')
        .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
        .att('Target', 'xl/workbook.xml');

        let xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('_rels').file('.rels', xmlString);
        resolve(promiseObj);

    });
};

let addWorkBookXML = (promiseObj) => {
    // Required as stated in §12.2
    return new Promise((resolve, reject) => {

        let xml = xmlbuilder.create(
            'workbook',
            {
                'version': '1.0', 
                'encoding': 'UTF-8', 
                'standalone': true
            }
        );
        xml.att('mc:Ignorable', 'x15');
        xml.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        xml.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        xml.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        xml.att('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');

        // bookViews (§18.2.1)
        if (promiseObj.wb.opts.workbookView) {
            const viewOpts = promiseObj.wb.opts.workbookView;
            let booksViewEle = xml.ele('bookViews');
            let workbookViewEle = booksViewEle.ele('workbookView');
            if (viewOpts.activeTab) {
                workbookViewEle.att('activeTab', viewOpts.activeTab);
            }
            if (viewOpts.autoFilterDateGrouping ) {
                workbookViewEle.att('autoFilterDateGrouping', utils.boolToInt(viewOpts.autoFilterDateGrouping));
            }
            if (viewOpts.firstSheet ) {
                workbookViewEle.att('firstSheet', viewOpts.firstSheet);
            }
            if (viewOpts.minimized ) {
                workbookViewEle.att('minimized', utils.boolToInt(viewOpts.minimized));
            }
            if (viewOpts.showHorizontalScroll ) {
                workbookViewEle.att('showHorizontalScroll', utils.boolToInt(viewOpts.showHorizontalScroll));
            }
            if (viewOpts.showSheetTabs ) {
                workbookViewEle.att('showSheetTabs', utils.boolToInt(viewOpts.showSheetTabs));
            }
            if (viewOpts.showVerticalScroll ) {
                workbookViewEle.att('showVerticalScroll', utils.boolToInt(viewOpts.showVerticalScroll));
            }
            if (viewOpts.tabRatio) {
                workbookViewEle.att('tabRatio', viewOpts.tabRatio);
            }
            if (viewOpts.visibility) {
                workbookViewEle.att('visibility', viewOpts.visibility);
            }
            if (viewOpts.windowWidth) {
                workbookViewEle.att('windowWidth', viewOpts.windowWidth);
            }
            if (viewOpts.windowHeight) {
                workbookViewEle.att('windowHeight', viewOpts.windowHeight);
            }
            if (viewOpts.xWindow) {
                workbookViewEle.att('xWindow', viewOpts.xWindow);
            }
            if (viewOpts.yWindow) {
                workbookViewEle.att('yWindow', viewOpts.yWindow);
            }
        }

        let sheetsEle = xml.ele('sheets');
        promiseObj.wb.sheets.forEach((s, i) => {
            sheetsEle.ele('sheet')
            .att('name', s.name)
            .att('sheetId', i + 1)
            .att('r:id', `rId${i + 1}`);
        });

        if (!promiseObj.wb.definedNameCollection.isEmpty) {
            promiseObj.wb.definedNameCollection.addToXMLele(xml);
        }

        let xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').file('workbook.xml', xmlString);
        resolve(promiseObj);

    });
};

let addWorkBookRelsXML = (promiseObj) => {
    // Required as stated in §12.2
    return new Promise((resolve, reject) => {

        let xml = xmlbuilder.create(
            'Relationships',
            {
                'version': '1.0', 
                'encoding': 'UTF-8', 
                'standalone': true
            }
        )
        .att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        xml
        .ele('Relationship')
        .att('Id', `rId${promiseObj.wb.sheets.length + 1}`)
        .att('Target', 'sharedStrings.xml')
        .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings');

        xml
        .ele('Relationship')
        .att('Id', `rId${promiseObj.wb.sheets.length + 2}`)
        .att('Target', 'styles.xml')
        .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');

        promiseObj.wb.sheets.forEach((s, i) => {
            xml
            .ele('Relationship')
            .att('Id', `rId${i + 1}`)
            .att('Target', `worksheets/sheet${i + 1}.xml`)
            .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
        });

        let xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').folder('_rels').file('workbook.xml.rels', xmlString);
        resolve(promiseObj);

    });
};

let addWorkSheetsXML = (promiseObj) => {
    // Required as stated in §12.2
    return new Promise ((resolve, reject) => {

        let curSheet = 0;
        
        let processNextSheet = () => {
            let thisSheet = promiseObj.wb.sheets[curSheet];
            if (thisSheet) {
                curSheet++;
                thisSheet.generateXML()
                .then((xml) => {
                    return new Promise((resolve) =>{
                        // Add worksheet to zip
                        promiseObj.xlsx.folder('xl').folder('worksheets').file(`sheet${curSheet}.xml`, xml); 
                        
                        resolve();
                    });
                })
                .then(() => {
                    return thisSheet.generateRelsXML();
                })
                .then((xml) => {
                    return new Promise((resolve) => {
                        if (xml) {
                            promiseObj.xlsx.folder('xl').folder('worksheets').folder('_rels').file(`sheet${curSheet}.xml.rels`, xml);
                        }
                        resolve();
                    });
                })
                .then(processNextSheet)
                .catch((e) => {
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
let addSharedStringsXML = (promiseObj) => {
    // §12.3.15 Shared String Table Part
    return new Promise ((resolve, reject) => {

        let xml = xmlbuilder.create(
            'sst',
            {
                'version': '1.0', 
                'encoding': 'UTF-8', 
                'standalone': true
            }
        )
        .att('count', promiseObj.wb.sharedStrings.length)
        .att('uniqueCount', promiseObj.wb.sharedStrings.length)
        .att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        promiseObj.wb.sharedStrings.forEach((s) => {
            if (typeof s === 'string') {
                xml.ele('si').ele('t').txt(s);
            } else if (s instanceof Array) {

                let thisSI = xml.ele('si');
                let theseRuns = []; // §18.4.4 r (Rich Text Run)
                let currProps = {};
                let curRun;
                let i = 0;
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
                    } else if (typeof s[i] === 'object') {
                        theseRuns.push({
                            props: {},
                            text: ''
                        });
                        curRun = theseRuns[theseRuns.length - 1];
                        Object.keys(s[i]).forEach((k) => {
                            currProps[k] = s[i][k];
                        });
                        Object.keys(currProps).forEach((k) => {
                            curRun.props[k] = currProps[k];
                        });
                        if (s[i].value !== undefined) {
                            curRun.text = s[i].value;
                        }
                    }
                    i++;
                }

                theseRuns.forEach((run) => {
                    if (Object.keys(run).length < 1) {
                        thisSI.ele('t', run.text).att('xml:space', 'preserve');
                    } else {
                        let thisRun = thisSI.ele('r');
                        let thisRunProps = thisRun.ele('rPr');
                        typeof run.props.name === 'string' ? thisRunProps.ele('rFont').att('val', run.props.name) : null;
                        run.props.bold === true ? thisRunProps.ele('b') : null;
                        run.props.italics === true ? thisRunProps.ele('i') : null;
                        run.props.strike === true ? thisRunProps.ele('strike') : null;
                        run.props.outline === true ? thisRunProps.ele('outline') : null;
                        run.props.shadow === true ? thisRunProps.ele('shadow') : null;
                        run.props.condense === true ? thisRunProps.ele('condense') : null;
                        run.props.extend === true ? thisRunProps.ele('extend') : null;
                        if (typeof run.props.color === 'string') {
                            let thisColor = new CTColor(run.props.color);
                            thisColor.addToXMLele(thisRunProps);
                        }
                        typeof run.props.size === 'number' ? thisRunProps.ele('sz').att('val', run.props.size) : null;
                        run.props.underline === true ? thisRunProps.ele('u') : null;
                        typeof run.props.vertAlign === 'string' ? thisRunProps.ele('vertAlign').att('val', run.props.vertAlign) : null;
                        thisRun.ele('t', run.text).att('xml:space', 'preserve');
                    }
                });

            }
        });

        let xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').file('sharedStrings.xml', xmlString);

        resolve(promiseObj);

    });
};

let addStylesXML = (promiseObj) => {
    // §12.3.20 Styles Part
    return new Promise ((resolve, reject) => {

        let xml = xmlbuilder.create(
            'styleSheet',
            {
                'version': '1.0', 
                'encoding': 'UTF-8', 
                'standalone': true
            }
        )
        .att('mc:Ignorable', 'x14ac')
        .att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        .att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        .att('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

        if (promiseObj.wb.styleData.numFmts.length > 0) {
            let nfXML = xml
            .ele('numFmts')
            .att('count', promiseObj.wb.styleData.numFmts.length);
            promiseObj.wb.styleData.numFmts.forEach((nf) => {
                nf.addToXMLele(nfXML);
            });
        }

        let fontXML = xml
        .ele('fonts')
        .att('count', promiseObj.wb.styleData.fonts.length);
        promiseObj.wb.styleData.fonts.forEach((f) => {
            f.addToXMLele(fontXML);
        });

        let fillXML = xml 
        .ele('fills')
        .att('count', promiseObj.wb.styleData.fills.length);
        promiseObj.wb.styleData.fills.forEach((f) => {
            let fXML = fillXML.ele('fill');
            f.addToXMLele(fXML);
        });

        let borderXML = xml 
        .ele('borders')
        .att('count', promiseObj.wb.styleData.borders.length);
        promiseObj.wb.styleData.borders.forEach((b) => {
            b.addToXMLele(borderXML);
        });


        let cellXfsXML = xml 
        .ele('cellXfs')
        .att('count', promiseObj.wb.styles.length);
        promiseObj.wb.styles.forEach((s) => {
            s.addXFtoXMLele(cellXfsXML);
        });

        if (promiseObj.wb.dxfCollection.length > 0) {
            promiseObj.wb.dxfCollection.addToXMLele(xml);
        }

        let xmlString = xml.doc().end(promiseObj.xmlOutVars);
        promiseObj.xlsx.folder('xl').file('styles.xml', xmlString);

        resolve(promiseObj);
    });
};

let addDrawingsXML = (promiseObj) => {
    return new Promise((resolve) => {
        if (!promiseObj.wb.mediaCollection.isEmpty) {

            promiseObj.wb.sheets.forEach((ws) => {
                if (!ws.drawingCollection.isEmpty) {

                    let drawingRelXML = xmlbuilder.create('Relationships', 
                        {
                            'version': '1.0', 
                            'encoding': 'UTF-8', 
                            'standalone': true
                        }
                    )
                    .att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                    let drawingsXML = xmlbuilder.create(
                        'xdr:wsDr',
                        {
                            'version': '1.0', 
                            'encoding': 'UTF-8', 
                            'standalone': true
                        }
                    );
                    drawingsXML
                    .att('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
                    .att('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');

                    ws.drawingCollection.drawings.forEach((d) => {

                        if (d.kind === 'image') {
                            let target = 'image' + d.id + '.' + d.extension;

                            let image = d.imagePath ? fs.readFileSync(d.imagePath) : d.image;
                            promiseObj.xlsx.folder('xl').folder('media').file(target, image);

                            drawingRelXML.ele('Relationship')
                            .att('Id', d.rId)
                            .att('Target', '../media/' + target)
                            .att('Type', d.type);

                        }



                        d.addToXMLele(drawingsXML);
                        
                    });

                    let drawingsXMLStr = drawingsXML.doc().end(promiseObj.xmlOutVars);
                    let drawingRelXMLStr = drawingRelXML.doc().end(promiseObj.xmlOutVars);
                    promiseObj.xlsx.folder('xl').folder('drawings').file('drawing' + ws.sheetId + '.xml', drawingsXMLStr);
                    promiseObj.xlsx.folder('xl').folder('drawings').folder('_rels').file('drawing' + ws.sheetId + '.xml.rels', drawingRelXMLStr);
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
let writeToBuffer = (wb) => {
    return new Promise ((resolve, reject) => {
        let promiseObj = {
            wb: wb, 
            xlsx: new JSZip(),
            xmlOutVars: {}
        };

        if (promiseObj.wb.sheets.length === 0) {
            promiseObj.wb.WorkSheet();
        }

        addRootContentTypesXML(promiseObj)
        .then(addWorkSheetsXML)
        .then(addRootRelsXML)
        .then(addWorkBookXML)
        .then(addWorkBookRelsXML)
        .then(addSharedStringsXML)
        .then(addStylesXML)
        .then(addDrawingsXML)
        .then(() => {
            wb.opts.jszip.type = 'nodebuffer';
            promiseObj.xlsx.generateAsync(wb.opts.jszip)
            .then((buf) => {
                resolve(buf);
            })
            .catch((e) => {
                reject(e);
            });
        })
        .catch((e) => {
            wb.logger.error(e.stack);
            reject(e);
        });

    });
};

module.exports = { writeToBuffer };