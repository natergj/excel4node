const tape = require('tape');
const _tape = require('tape-promise').default;
const test = _tape(tape);
const xl = require('../source/index');
const DOMParser = require('xmldom').DOMParser;

test('Generate multiple sheets', (t) => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let ws2 = wb.addWorksheet('test2');
    let ws3 = wb.addWorksheet('test3');

    t.ok(wb.sheets.length === 3, 'Correctly generated multiple sheets');

    wb.setSelectedTab(2);
    t.ok(
        wb.sheets[0].opts.sheetView.tabSelected === 0 &&
        wb.sheets[1].opts.sheetView.tabSelected === 1 &&
        wb.sheets[2].opts.sheetView.tabSelected === 0, '2nd Tab set to be default tab selected');

    t.end();
});

test('Set Worksheet options', (t) => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test 1', {
        'margins': { // Accepts a Double in Inches
            'bottom': 1.25,
            'footer': 1.5,
            'header': 2.0,
            'left': 0.5,
            'right': 0.75,
            'top': 1.0
        },
        'printOptions': {
            'centerHorizontal': true,
            'centerVertical': true,
            'printGridLines': true,
            'printHeadings': true

        },
        'headerFooter': { // Set Header and Footer strings and options. See note below
            'evenFooter': 'Even Footer String',
            'evenHeader': 'Even Header String',
            'firstFooter': 'First Footer String',
            'firstHeader': 'First Header String',
            'oddFooter': 'Odd Footer String',
            'oddHeader': 'Odd Header String',
            'alignWithMargins': true,
            'differentFirst': true,
            'differentOddEven': true,
            'scaleWithDoc': true
        },
        'pageSetup': {
            'blackAndWhite': true,
            'cellComments': 'none', // one of 'none', 'asDisplayed', 'atEnd'
            'copies': 1,
            'draft': false,
            'errors': 'NA', // One of 'displayed', 'blank', 'dash', 'NA'
            'firstPageNumber': 1,
            'fitToHeight': 10, // Number of vertical pages to fit to
            'fitToWidth': 1, // Number of horizontal pages to fit to
            'horizontalDpi': 96,
            'orientation': 'landscape', // One of 'default', 'portrait', 'landscape'
            'pageOrder': 'overThenDown', // One of 'downThenOver', 'overThenDown'
            'paperHeight': '11in', // Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. '10.5cm'
            'paperSize': 'A4_PAPER', // see lib/types/paperSize.js for all types and descriptions of types
            'paperWidth': '8.5in',
            'scale': 100,
            'useFirstPageNumber': true,
            'usePrinterDefaults': true,
            'verticalDpi': 96
        },
        'sheetView': {
            'pane': { // Note. Calling .freeze() on a row or column will adjust these values 
                'activePane': 'bottomLeft', // one of 'bottomLeft', 'bottomRight', 'topLeft', 'topRight'
                'state': 'frozen', // one of 'split', 'frozen', 'frozenSplit'
                'topLeftCell': 'E3', // i.e. 'A1'
                'xSplit': 5, // Horizontal position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of columns visible in the top pane.
                'ySplit': 3 // Vertical position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of rows visible in the left pane.
            },
            'rightToLeft': false, // Flag indicating whether the sheet is in 'right to left' display mode. When in this mode, Column A is on the far right, Column B ;is one column left of Column A, and so on. Also, information in cells is displayed in the Right to Left format.
            'showGridLines': false, // Flag indicating whether the sheet should have gridlines enabled or disabled during view
            'zoomScale': 110, // Defaults to 100
            'zoomScaleNormal': 120, // Defaults to 100
            'zoomScalePageLayoutView': 130 // Defaults to 100
        },
        'sheetFormat': {
            'baseColWidth': 12, // Defaults to 10. Specifies the number of characters of the maximum digit width of the normal style's font. This value does not include margin padding or extra padding for gridlines. It is only the number of characters.,
            'defaultColWidth': 16,
            'defaultRowHeight': 20,
            'thickBottom': false, // 'True' if rows have a thick bottom border by default.
            'thickTop': true // 'True' if rows have a thick top border by default.
        },
        'sheetProtection': { // same as 'Protect Sheet' in Review tab of Excel 
            'autoFilter': true, // True means that that user will be unable to modify this setting
            'deleteColumns': true,
            'deleteRows': true,
            'formatCells': true,
            'formatColumns': true,
            'formatRows': true,
            'insertColumns': true,
            'insertHyperlinks': false,
            'insertRows': true,
            'objects': true,
            'password': 'myPassword',
            'pivotTables': true,
            'scenarios': true,
            'selectLockedCells': true,
            'selectUnlockedCells': true,
            'sheet': true,
            'sort': true
        },
        'outline': {
            'summaryBelow': true, // Flag indicating whether summary rows appear below detail in an outline, when applying an outline/grouping.
            'summaryRight': true // Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline/grouping.
        },
        'hidden': true // Flag indicating whether to not hide the worksheet within the workbook.
    });
    ws.row(2).filter(1, 10);

    ws.generateXML().then((XML) => {
        let doc = new DOMParser().parseFromString(XML);
        let margins = doc.getElementsByTagName('pageMargins')[0];
        t.equals(margins.getAttribute('bottom'), '1.25', 'Bottom margin properly set');
        t.equals(margins.getAttribute('footer'), '1.5', 'Footer margin properly set');
        t.equals(margins.getAttribute('header'), '2', 'Header margin properly set');
        t.equals(margins.getAttribute('left'), '0.5', 'Left margin properly set');
        t.equals(margins.getAttribute('right'), '0.75', 'Right margin properly set');
        t.equals(margins.getAttribute('top'), '1', 'Top margin properly set');

        let printOptions = doc.getElementsByTagName('printOptions')[0];
        t.equals(printOptions.getAttribute('horizontalCentered'), '1', 'Horizontal print center option set');
        t.equals(printOptions.getAttribute('verticalCentered'), '1', 'Vertical print center option correctly not set.');
        t.equals(printOptions.getAttribute('gridLines'), '1', 'print gridlines print option correctly not set.');
        t.equals(printOptions.getAttribute('gridLinesSet'), '1', 'print gridlines print option correctly not set.');
        t.equals(printOptions.getAttribute('headings'), '1', 'printHeadings print option set');

        let headerFooter = doc.getElementsByTagName('headerFooter')[0];
        t.equals(headerFooter.getAttribute('alignWithMargins'), '1', 'headerFooter alignWithMargins set correctly');
        t.equals(headerFooter.getAttribute('differentFirst'), '1', 'headerFooter differentFirst set correctly');
        t.equals(headerFooter.getAttribute('differentOddEven'), '1', 'headerFooter differentOddEven set correctly');
        t.equals(headerFooter.getAttribute('scaleWithDoc'), '1', 'headerFooter scaleWithDoc set correctly');
        t.equals(headerFooter.getElementsByTagName('evenFooter')[0].childNodes[0].data, 'Even Footer String', 'Even footer string set correctly');
        t.equals(headerFooter.getElementsByTagName('evenHeader')[0].childNodes[0].data, 'Even Header String', 'Even header string set correctly');
        t.equals(headerFooter.getElementsByTagName('firstFooter')[0].childNodes[0].data, 'First Footer String', 'First footer string set correctly');
        t.equals(headerFooter.getElementsByTagName('firstHeader')[0].childNodes[0].data, 'First Header String', 'First header string set correctly');
        t.equals(headerFooter.getElementsByTagName('oddFooter')[0].childNodes[0].data, 'Odd Footer String', 'Odd footer string set correctly');
        t.equals(headerFooter.getElementsByTagName('oddHeader')[0].childNodes[0].data, 'Odd Header String', 'Odd header string set correctly');

        let pageSetup = doc.getElementsByTagName('pageSetup')[0];
        t.equals(pageSetup.getAttribute('paperSize'), '9', 'pageSetup paperSize was set correctly.');
        t.equals(pageSetup.getAttribute('paperHeight'), '11in', 'pageSetup paperHeight was set correctly.');
        t.equals(pageSetup.getAttribute('paperWidth'), '8.5in', 'pageSetup paperWidth was set correctly.');
        t.equals(pageSetup.getAttribute('scale'), '100', 'pageSetup scale was set correctly.');
        t.equals(pageSetup.getAttribute('firstPageNumber'), '1', 'pageSetup firstPageNumber was set correctly.');
        t.equals(pageSetup.getAttribute('fitToWidth'), '1', 'pageSetup fitToWidth was set correctly.');
        t.equals(pageSetup.getAttribute('fitToHeight'), '10', 'pageSetup fitToHeight was set correctly.');
        t.equals(pageSetup.getAttribute('pageOrder'), 'overThenDown', 'pageSetup pageOrder was set correctly.');
        t.equals(pageSetup.getAttribute('orientation'), 'landscape', 'pageSetup orientation was set correctly.');
        t.equals(pageSetup.getAttribute('usePrinterDefaults'), '1', 'pageSetup usePrinterDefaults was set correctly.');
        t.equals(pageSetup.getAttribute('blackAndWhite'), '1', 'pageSetup blackAndWhite was set correctly.');
        t.equals(pageSetup.getAttribute('draft'), '0', 'pageSetup draft was set correctly.');
        t.equals(pageSetup.getAttribute('cellComments'), 'none', 'pageSetup cellComments was set correctly.');
        t.equals(pageSetup.getAttribute('useFirstPageNumber'), '1', 'pageSetup useFirstPageNumber was set correctly.');
        t.equals(pageSetup.getAttribute('errors'), 'NA', 'pageSetup errors was set correctly.');
        t.equals(pageSetup.getAttribute('horizontalDpi'), '96', 'pageSetup horizontalDpi was set correctly.');
        t.equals(pageSetup.getAttribute('verticalDpi'), '96', 'pageSetup verticalDpi was set correctly.');
        t.equals(pageSetup.getAttribute('copies'), '1', 'pageSetup copies was set correctly.');

        let sheetView = doc.getElementsByTagName('sheetView')[0];
        t.equals(sheetView.getAttribute('workbookViewId'), '0', 'sheetView workbookViewId was set correctly');
        t.equals(sheetView.getAttribute('rightToLeft'), 'false', 'sheetView rightToLeft was set correctly');
        t.equals(sheetView.getAttribute('showGridLines'), 'false', 'sheetView showGridLines was set correctly');
        t.equals(sheetView.getAttribute('zoomScale'), '110', 'sheetView zoomScale was set correctly');
        t.equals(sheetView.getAttribute('zoomScaleNormal'), '120', 'sheetView zoomScaleNormal was set correctly');
        t.equals(sheetView.getAttribute('zoomScalePageLayoutView'), '130', 'sheetView zoomScalePageLayoutView was set correctly');
        t.equals(sheetView.getElementsByTagName('pane')[0].getAttribute('xSplit'), '5', 'sheetView xSplit set correctly');
        t.equals(sheetView.getElementsByTagName('pane')[0].getAttribute('ySplit'), '3', 'sheetView ySplit set correctly');
        t.equals(sheetView.getElementsByTagName('pane')[0].getAttribute('topLeftCell'), 'E3', 'sheetView topLeftCell set correctly');
        t.equals(sheetView.getElementsByTagName('pane')[0].getAttribute('activePane'), 'bottomLeft', 'sheetView activePane set correctly');
        t.equals(sheetView.getElementsByTagName('pane')[0].getAttribute('state'), 'frozen', 'sheetView state set correctly');

        let sheetFormat = doc.getElementsByTagName('sheetFormatPr')[0];
        t.equals(sheetFormat.getAttribute('baseColWidth'), '12', 'sheetFormat baseColWidth was correctly set');
        t.equals(sheetFormat.getAttribute('defaultColWidth'), '16', 'sheetFormat defaultColWidth was correctly set');
        t.equals(sheetFormat.getAttribute('defaultRowHeight'), '20', 'sheetFormat defaultRowHeight was correctly set');
        t.equals(sheetFormat.getAttribute('thickBottom'), '0', 'sheetFormat thickBottom was correctly set');
        t.equals(sheetFormat.getAttribute('thickTop'), '1', 'sheetFormat thickTop was correctly set');
        t.equals(sheetFormat.getAttribute('customHeight'), '1', 'sheetFormat customHeight was correctly set');

        let sheetProtection = doc.getElementsByTagName('sheetProtection')[0];
        t.equals(sheetProtection.getAttribute('autoFilter'), '1', 'sheetProtection autoFilter was correctly set');
        t.equals(sheetProtection.getAttribute('deleteColumns'), '1', 'sheetProtection deleteColumns was correctly set');
        t.equals(sheetProtection.getAttribute('deleteRows'), '1', 'sheetProtection deleteRows was correctly set');
        t.equals(sheetProtection.getAttribute('formatCells'), '1', 'sheetProtection formatCells was correctly set');
        t.equals(sheetProtection.getAttribute('formatColumns'), '1', 'sheetProtection formatColumns was correctly set');
        t.equals(sheetProtection.getAttribute('formatRows'), '1', 'sheetProtection formatRows was correctly set');
        t.equals(sheetProtection.getAttribute('insertColumns'), '1', 'sheetProtection insertColumns was correctly set');
        t.equals(sheetProtection.getAttribute('insertHyperlinks'), '0', 'sheetProtection insertHyperlinks was correctly set');
        t.equals(sheetProtection.getAttribute('insertRows'), '1', 'sheetProtection insertRows was correctly set');
        t.equals(sheetProtection.getAttribute('objects'), '1', 'sheetProtection objects was correctly set');
        t.equals(sheetProtection.getAttribute('password'), 'F9CD', 'sheetProtection password was correctly set');
        t.equals(sheetProtection.getAttribute('pivotTables'), '1', 'sheetProtection pivotTables was correctly set');
        t.equals(sheetProtection.getAttribute('scenarios'), '1', 'sheetProtection scenarios was correctly set');
        t.equals(sheetProtection.getAttribute('selectLockedCells'), '1', 'sheetProtection selectLockedCells was correctly set');
        t.equals(sheetProtection.getAttribute('selectUnlockedCells'), '1', 'sheetProtection selectUnlockedCells was correctly set');
        t.equals(sheetProtection.getAttribute('sheet'), '1', 'sheetProtection sheet was correctly set');
        t.equals(sheetProtection.getAttribute('sort'), '1', 'sheetProtection sort was correctly set');

        let sheetPr = doc.getElementsByTagName('sheetPr')[0];
        let outlinePr = sheetPr.getElementsByTagName('outlinePr')[0];
        t.equals(outlinePr.getAttribute('applyStyles'), '1', 'outline property applyStyles set correctly');
        t.equals(outlinePr.getAttribute('summaryBelow'), '1', 'outline property summaryBelow set correctly');
        t.equals(outlinePr.getAttribute('summaryRight'), '1', 'outline property summaryRight set correctly');

        t.end();
    });

    wb._generateXML().then((XML) => {
        let doc = new DOMParser().parseFromString(XML);
        let sheet = doc.getElementsByTagName('sheet')[0];
        t.equals(sheet.getAttribute('state'), 'hidden', 'Hidden state properly set');
    });
});

test('Verify Invalid Worksheet options fail type validation', (t) => {
    let wb = new xl.Workbook();

    try {
        let ws = wb.addWorksheet('sheet', {
            pageSetup: {
                cellComments: 'invalid option'
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid pageSetup.cellComments property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid pageSetup.cellComments property should throw an error'
        );
    }


    try {
        let ws = wb.addWorksheet('sheet', {
            pageSetup: {
                errors: 'invalid option'
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid pageSetup.pageSetup property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid pageSetup.errors property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            pageSetup: {
                orientation: 'invalid option'
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid pageSetup.orientation property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid pageSetup.orientation property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            pageSetup: {
                pageOrder: 'invalid option'
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid pageSetup.pageOrder property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid pageSetup.pageOrder property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            pageSetup: {
                paperHeight: 'invalid option'
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid pageSetup.paperHeight property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid pageSetup.paperHeight property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            pageSetup: {
                paperSize: 'invalid option'
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid pageSetup.paperSize property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid pageSetup.paperSize property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            pageSetup: {
                paperWidth: 'invalid option'
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid pageSetup.paperWidth property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid pageSetup.paperWidth property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            sheetView: {
                pane: {
                    activePane: 'invalid option'
                }
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid sheetView.pane.activePane property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid sheetView.pane.activePane property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            sheetView: {
                pane: {
                    state: 'invalid option'
                }
            }
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid sheetView.pane.state property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid sheetView.pane.state property should throw an error'
        );
    }

    try {
        let ws = wb.addWorksheet('sheet', {
            hidden: 'notBoolean'
        });
        t.notOk(typeof ws === 'object', 'Worksheet creation should fail when setting invalid hidden property');
    } catch (e) {
        t.ok(
            e instanceof TypeError,
            'setting invalid hidden property should throw an error'
        );
    }


    t.end();
});

test('Check worksheet defaultRowHeight behavior', (t) => {
    // The defaultRowHeight property effects more than its own tag in the XML. 
    // need to check output of sheetFormatPr.customHeight, sheetFormatPr.defaultRowHeight and each row's customHeight attribute
    // With no sheetFormat.customHeight set, the row height should scale to fit the text and ignore the sheetFormatPr.defaultRowHeight value

    let wb = new xl.Workbook({
        defaultFont: {
            size: 9,
            name: 'Arial'
        }
    });

    let ws1 = wb.addWorksheet('Sheet1', {
        sheetFormat: {
            defaultRowHeight: 12
        }
    });
    ws1.cell(1, 1).string('String');

    let ws2 = wb.addWorksheet('Sheet2');
    ws2.cell(1, 1).string('String');

    ws1.generateXML()
        .then((XML) => {
            let doc = new DOMParser().parseFromString(XML);
            let sheetFormatPr = doc.getElementsByTagName('sheetFormatPr')[0];
            t.equals(sheetFormatPr.getAttribute('defaultRowHeight'), '12', 'Required attribute sheetFormatPr.defalutRowHeight successfully updated with custom row height');
            t.equals(sheetFormatPr.getAttribute('customHeight'), '1', 'Optional sheetFormatPr.customHeight successfully set to be true');

            let firstRow = doc.getElementsByTagName('row')[0];
            t.equals(firstRow.getAttribute('customHeight'), '1', 'customHeight attribute on row successfully set to 1 since sheet default row height specified');
        })
        .then(() => {
            return ws2.generateXML();
        })
        .then((XML) => {
            let doc = new DOMParser().parseFromString(XML);
            let sheetFormatPr = doc.getElementsByTagName('sheetFormatPr')[0];
            t.equals(sheetFormatPr.getAttribute('defaultRowHeight'), '16', 'Required attribute sheetFormatPr.defalutRowHeight successfully set with default value');
            t.equals(sheetFormatPr.getAttribute('customHeight'), '', 'Optional sheetFormatPr.customHeight not set when sheetFormat.defaultRowHeight not specified');

            let firstRow = doc.getElementsByTagName('row')[0];
            t.equals(firstRow.getAttribute('customHeight'), '', 'customHeight attribute on row successfully not set since sheet default row height not specified');
        })
        .then(() => {
            t.end();
        });

});

test('Check worksheet addPageBreak behavior', (t) => {
    let wb = new xl.Workbook({
        defaultFont: {
            size: 9,
            name: 'Arial'
        }
    });

    let ws1 = wb.addWorksheet('Sheet1');
    ws1.cell(1, 1).string('String');
    ws1.addPageBreak('row', 1);
    ws1.addPageBreak('row', 6);
    ws1.addPageBreak('column', 8);

    ws1.generateXML()
        .then((XML) => {
            let doc = new DOMParser().parseFromString(XML);
            let rowBreakXml = doc.getElementsByTagName('rowBreaks')[0];
            t.equals(rowBreakXml.getAttribute('count'), '2', 'has 2 rowBreak2');
            t.equals(rowBreakXml.getAttribute('manualBreakCount'), '2', 'has 2 manualBreakCount');

            let firstBrk = rowBreakXml.getElementsByTagName('brk')[0];
            t.equals(firstBrk.getAttribute('id'), '1', 'first break is at correct position');
            let secondBrk = rowBreakXml.getElementsByTagName('brk')[1];
            t.equals(secondBrk.getAttribute('id'), '6', 'second break is at correct position');

            let colBreakXml = doc.getElementsByTagName('colBreaks')[0];
            t.equals(colBreakXml.getAttribute('count'), '1', 'has 1 column break');
            let firstColBreak= colBreakXml.getElementsByTagName('brk')[0];
            t.equals(firstColBreak.getAttribute('id'), '8', 'first column break in correct position');
        })
        .then(() => {
            t.end();
        });

});