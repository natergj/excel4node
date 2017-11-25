const fs = require('fs');
const path = require('path');

require('source-map-support').install();
var xl = require('./distribution');

function generateWorkbook() {
    var wb = new xl.Workbook({
        defaultFont: {
            name: 'Verdana',
            size: 12
        },
        dateFormat: 'mm/dd/yyyy hh:mm:ss',
        logLevel: 1,
    });

    /*****************************************
     * START Create a sample invoice
     *****************************************/

    // Create some styles to be used throughout
    var multiLineStyle = wb.createStyle({
        alignment: {
            wrapText: true,
            vertical: 'top'
        }
    });
    var largeText = wb.createStyle({
        font: {
            name: 'Cambria',
            size: 20
        }
    });
    var medText = wb.createStyle({
        font: {
            name: 'Cambria',
            size: 14,
            color: '#D4762C'
        },
        alignment: {
            vertical: 'center'
        }
    });

    var currencyStyle = wb.createStyle({
        numberFormat: '$##0.00; [Red]($##0.00); $0.00'
    });

    var invoiceWS = wb.addWorksheet('Invoice', {
        pageSetup: {
            fitToWidth: 1
        },
        headerFooter: {
            oddHeader: 'iAmNater invoice',
            oddFooter: 'Invoice Page &P'
        }
    });

    // Set some row and column properties
    invoiceWS.row(1).setHeight(25);
    invoiceWS.row(2).setHeight(45);
    invoiceWS.column(1).setWidth(3);
    invoiceWS.column(2).setWidth(10);
    invoiceWS.column(3).setWidth(35);
    invoiceWS.column(5).setWidth(25);
    invoiceWS.cell(2, 2).string('INVOICE').style(largeText);
    invoiceWS.cell(2, 3).string('809871').style(largeText).style({ font: { color: '#D4762C' } });

    // Add a company logo
    invoiceWS.addImage({
        path: './sampleFiles/logo.png',
        type: 'picture',
        position: {
            type: 'twoCellAnchor',
            from: {
                col: 4,
                colOff: 0,
                row: 2,
                rowOff: 0
            },
            to: {
                col: 6,
                colOff: 0,
                row: 3,
                rowOff: 0
            }
        }
    });

    // Add some borders to specific cells
    invoiceWS.cell(2, 2, 2, 5).style({ border: { bottom: { style: 'thick', color: '#000000' } } });

    // Add some data and adjust styles for specific cells
    invoiceWS.cell(3, 2, 3, 3, true).string('January 1, 2016').style({ border: { bottom: { style: 'thin', color: '#D4762C' } } });
    invoiceWS.cell(4, 2, 4, 3, true).string('PAYMENT DUE BY: March 1, 2016').style({ font: { bold: true } });

    // style methods can be chained. multiple styles will be merged with last style taking precedence if there is a conflict
    invoiceWS.cell(3, 5, 4, 5, true).formula('E31').style(currencyStyle).style({ font: { size: 20, color: '#D4762C' }, alignment: { vertical: 'center' } });
    invoiceWS.cell(4, 2, 4, 5).style({ border: { bottom: { style: 'thin', color: '#000000' } } });

    invoiceWS.row(6).setHeight(75);
    invoiceWS.cell(6, 2, 6, 5).style(multiLineStyle);

    // set some strings to have multiple font formats within a single cell
    invoiceWS.cell(6, 2, 6, 3, true).string([
        {
            bold: true
        },
        'Client Name\n',
        {
            bold: false
        },
        'Company Name Inc.\n1234 First Street\nSomewhere, OR 12345'
    ]);

    invoiceWS.cell(6, 4, 6, 5, true).string([
        {
            bold: true
        },
        'iAmNater.com\n',
        {
            bold: false
        },
        '123 Nowhere Lane\nSomewhere, OR 12345'
    ]).style({ alignment: { horizontal: 'right' } });

    invoiceWS.cell(8, 2, 8, 5).style({ border: { bottom: { style: 'thick', color: '#000000' } } });

    invoiceWS.cell(10, 2).string('QUANTITY');
    invoiceWS.cell(10, 3).string('DETAILS');
    invoiceWS.cell(10, 4).string('UNIT PRICE').style({ alignment: { horizontal: 'right' } });
    invoiceWS.cell(10, 5).string('LINE TOTAL').style({ alignment: { horizontal: 'right' } });

    var items = require('./sampleFiles/invoiceData.json').items;
    var i = 0;
    var rowOffset = 11;
    var oddBackgroundColor = '#F8F5EE';
    while (i <= 10) {
        var item = items[i];
        var curRow = rowOffset + i;
        if (item !== undefined) {
            invoiceWS.cell(curRow, 2).number(item.units).style({ alignment: { horizontal: 'left' } });
            invoiceWS.cell(curRow, 3).string(item.description);
            invoiceWS.cell(curRow, 4).number(item.unitCost).style(currencyStyle);
            invoiceWS.cell(curRow, 5).formula(xl.getExcelCellRef(rowOffset + i, 2) + '*' + xl.getExcelCellRef(rowOffset + 1, 4)).style(currencyStyle);
        }
        if (i % 2 === 0) {
            invoiceWS.cell(curRow, 2, curRow, 5).style({
                fill: {
                    type: 'pattern',
                    patternType: 'solid',
                    fgColor: oddBackgroundColor
                }
            });
        }
        i++;
    }
    invoiceWS.cell(21, 2, 21, 5).style({ border: { bottom: { style: 'thin', color: '#DCD1B3' } } });

    invoiceWS.cell(22, 4).string('Discount');
    invoiceWS.cell(22, 5).number(0.00).style(currencyStyle);

    invoiceWS.cell(23, 4).string('Net Total');
    invoiceWS.cell(23, 5).formula('SUM(E11:E21)').style(currencyStyle);

    invoiceWS.cell(23, 2, 23, 5).style({ border: { bottom: { style: 'thin', color: '#000000' } } });

    invoiceWS.row(24).setHeight(20);
    invoiceWS.cell(24, 4, 25, 4, true).string('USD TOTAL').style(medText);
    invoiceWS.cell(24, 5, 25, 5, true).formula('SUM(E22:E23)').style(medText).style(currencyStyle);
    /*****************************************
     * END Create a sample invoice
     *****************************************/


    /*****************************************
     * START Create a filterable list
     *****************************************/

    var filterSheet = wb.addWorksheet('Filters');

    for (var i = 1; i <= 10; i++) {
        filterSheet.cell(1, i).string('Header' + i);
    }
    filterSheet.row(1).filter(1, 10);

    for (var r = 2; r <= 30; r++) {
        for (var c = 1; c <= 10; c++) {
            filterSheet.cell(r, c).number(parseInt(Math.random() * 100));
        }
    }
     /*****************************************
     * END Create a filterable list
     *****************************************/

    /*****************************************
     * START Create collapsable lists
     *****************************************/


    var collapseSheet = wb.addWorksheet('Collapsables', {
        pageSetup: {
            fitToWidth: 1
        },
        outline: {
            summaryBelow: true
        }
    });

    var rowOffset = 0;
    for (var r = 1; r <= 10; r++) {
        for (var c = 1; c <= 10; c++) {
            collapseSheet.cell(r + rowOffset, c).number(parseInt(Math.random() * 100));
        }
        collapseSheet.row(r + rowOffset).group(1, true);
    }
    for (var i = 1; i <= 10; i++) {
        collapseSheet.cell(11, i).formula('SUM(' + xl.getExcelCellRef(rowOffset + 1, i) + ':' + xl.getExcelCellRef(rowOffset + 10, i) + ')');
    }
    collapseSheet.cell(11, 1, 11, 10).style({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#C2D6EC'
        }
    });

    var rowOffset = 11;
    for (var r = 1; r <= 10; r++) {
        for (var c = 1; c <= 10; c++) {
            collapseSheet.cell(r + rowOffset, c).number(parseInt(Math.random() * 100));
        }
        collapseSheet.row(r + rowOffset).group(1, true);
    }
    for (var i = 1; i <= 10; i++) {
        collapseSheet.cell(22, i).formula('SUM(' + xl.getExcelCellRef(rowOffset + 1, i) + ':' + xl.getExcelCellRef(rowOffset + 10, i) + ')');
    }
    collapseSheet.cell(22, 1, 22, 10).style({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#4273B0'
        }
    });



    var rowOffset = 22;
    for (var r = 1; r <= 10; r++) {
        for (var c = 1; c <= 10; c++) {
            collapseSheet.cell(r + rowOffset, c).number(parseInt(Math.random() * 100));
        }
        collapseSheet.row(r + rowOffset).group(1, true);
    }
    for (var i = 1; i <= 10; i++) {
        collapseSheet.cell(33, i).formula('SUM(' + xl.getExcelCellRef(rowOffset + 1, i) + ':' + xl.getExcelCellRef(rowOffset + 10, i) + ')');
    }
    collapseSheet.cell(33, 1, 33, 10).style({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#C2D6EC'
        }
    });

    var rowOffset = 33;
    for (var r = 1; r <= 10; r++) {
        for (var c = 1; c <= 10; c++) {
            collapseSheet.cell(r + rowOffset, c).number(parseInt(Math.random() * 100));
        }
        collapseSheet.row(r + rowOffset).group(1, true);
    }
    for (var i = 1; i <= 10; i++) {
        collapseSheet.cell(44, i).formula('SUM(' + xl.getExcelCellRef(rowOffset + 1, i) + ':' + xl.getExcelCellRef(rowOffset + 10, i) + ')');
    }
    collapseSheet.cell(44, 1, 44, 10).style({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#4273B0'
        }
    });



    var rowOffset = 44;
    for (var r = 1; r <= 10; r++) {
        for (var c = 1; c <= 10; c++) {
            collapseSheet.cell(r + rowOffset, c).number(parseInt(Math.random() * 100));
        }
        collapseSheet.row(r + rowOffset).group(1, true);
    }
    for (var i = 1; i <= 10; i++) {
        collapseSheet.cell(55, i).formula('SUM(' + xl.getExcelCellRef(rowOffset + 1, i) + ':' + xl.getExcelCellRef(rowOffset + 10, i) + ')');
    }
    collapseSheet.cell(55, 1, 55, 10).style({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#C2D6EC'
        }
    });

    var rowOffset = 55;
    for (var r = 1; r <= 10; r++) {
        for (var c = 1; c <= 10; c++) {
            collapseSheet.cell(r + rowOffset, c).number(parseInt(Math.random() * 100));
        }
        collapseSheet.row(r + rowOffset).group(1, true);
    }
    for (var i = 1; i <= 10; i++) {
        collapseSheet.cell(66, i).formula('SUM(' + xl.getExcelCellRef(rowOffset + 1, i) + ':' + xl.getExcelCellRef(rowOffset + 10, i) + ')');
    }
    collapseSheet.cell(66, 1, 66, 10).style({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#4273B0'
        }
    });
    /*****************************************
     * START Create collapsable lists
     *****************************************/

    /*****************************************
     * START Create Frozen lists
     *****************************************/

    var frozenSheet = wb.addWorksheet('Frozen');

    for (var i = 2; i <= 21; i++) {
        frozenSheet.cell(1, i).string('Column' + i);
    }
    frozenSheet.row(1).freeze();

    for (var r = 2; r <= 30; r++) {
        frozenSheet.cell(r, 1).string('Row' + r);
        for (var c = 2; c <= 22; c++) {
            frozenSheet.cell(r, c).number(parseInt(Math.random() * 100));
        }
    }
    frozenSheet.column(1).freeze();

     /*****************************************
     * END Create Frozen lists
     *****************************************/

    /*****************************************
     * START Create Split sheet
     *****************************************/

    var splitSheet = wb.addWorksheet('SplitSheet', {
        'sheetView': {
            'pane': {
                'activePane': 'bottomRight',
                'state': 'split',
                'xSplit': 2000,
                'ySplit': 3000
            }
        }
    });

    for (var r = 1; r <= 30; r++) {
        for (var c = 1; c <= 20; c++) {
            splitSheet.cell(r, c).number(parseInt(Math.random() * 100));
        }
    }

     /*****************************************
     * END Create Split
     *****************************************/

    /*****************************************
     * START Create Selectable Options list
     *****************************************/
    var optionsSheet = wb.addWorksheet('Selectable Options');

    optionsSheet.cell(1, 1).string('Booleans');
    optionsSheet.cell(1, 2).string('Option List');
    optionsSheet.cell(1, 3).string('Numbers 1-10');

    optionsSheet.addDataValidation({
        type: 'list',
        allowBlank: true,
        prompt: 'Choose from dropdown',
        error: 'Invalid choice was chosen',
        sqref: 'A2:A10',
        formulas: [
            'true,false'
        ]
    });

    optionsSheet.addDataValidation({
        type: 'list',
        allowBlank: true,
        prompt: 'Choose from dropdown',
        promptTitle: 'Choose from dropdown',
        error: 'Invalid choice was chosen',
        showInputMessage: true,
        sqref: 'B2:B10',
        formulas: [
            'option 1,option 2,option 3'
        ]
    });

    optionsSheet.addDataValidation({
        errorStyle: 'stop',
        error: 'Number must be between 1 and 10',
        type: 'whole',
        operator: 'between',
        allowBlank: 1,
        sqref: 'C2:C10',
        formulas: [1, 10]
    });
    /*****************************************
     * END Create Selectable Options list
     *****************************************/

    /*****************************************
     * START date check sheet
     *****************************************/

    var datesSheet = wb.addWorksheet('dates');

    datesSheet.column(1).setWidth(24);
    datesSheet.column(2).setWidth(20);
    datesSheet.column(3).setWidth(45);

    datesSheet.cell(1, 1).string('Dates');
    datesSheet.cell(1, 2).string('Serial Value');
    datesSheet.cell(1, 3).string('Date String');

    datesSheet.cell(2, 1).date(new Date('1910-02-03T10:05:54Z'));
    datesSheet.cell(2, 2).string('3687.4207639');
    datesSheet.cell(2, 3).string('Correctly translated 1910-02-03T10:05:54Z');
    datesSheet.cell(3, 1).date(new Date('1900-01-01T12:00:00Z'));
    datesSheet.cell(3, 2).string('1.5000000');
    datesSheet.cell(3, 3).string('Correctly translated 1900-01-01T12:00:00Z');
    datesSheet.cell(4, 1).date(new Date('9999-12-31T23:59:59Z'));
    datesSheet.cell(4, 2).string('2958465.9999884');
    datesSheet.cell(4, 3).string('Correctly translated 9999-12-31T23:59:59Z');


    datesSheet.cell(5, 1).date(new Date('1900-01-01T00:00:00Z'));
    datesSheet.cell(5, 2).string('1');
    datesSheet.cell(5, 3).string('Correctly translated 1900-01-01');
    datesSheet.cell(6, 1).date(new Date('1910-02-03T00:00:00Z'));
    datesSheet.cell(6, 2).string('3687');
    datesSheet.cell(6, 3).string('Correctly translated 1910-02-03');
    datesSheet.cell(7, 1).date(new Date('2006-02-01T00:00:00Z'))
    datesSheet.cell(7, 2).string('38749');
    datesSheet.cell(7, 3).string('Correctly translated 2006-02-01');
    datesSheet.cell(8, 1).date(new Date('9999-12-31T00:00:00Z'));
    datesSheet.cell(8, 2).string('2958465');
    datesSheet.cell(8, 3).string('Correctly translated 9999-12-31');


    /*****************************************
     * END date check sheet
     *****************************************/

    /*****************************************
     * START final sheet
     *****************************************/

    var imageSheet = wb.addWorksheet('images', {
        pageSetup: {
            orientation: 'landscape'
        },
        sheetView: {
            zoomScale: 120
        }
    });

    imageSheet.cell(1, 1).string('Images');
    imageSheet.addImage({
        path: './sampleFiles/thumbsUp.jpg',
        type: 'picture',
        position: {
            type: 'absoluteAnchor',
            x: 0,
            y: '10mm'
        }
    });
    imageSheet.addImage({
        image: fs.readFileSync(path.resolve(__dirname, './sampleFiles/thumbsUp.jpg')),
        type: 'picture',
        position: {
            type: 'oneCellAnchor',
            from: {
                col: 1,
                colOff: 0,
                row: 32,
                rowOff: 0
            }
        }
    });

    /*****************************************
     * END final sheet
     *****************************************/

    return wb;
}


var wb = generateWorkbook();

/*
wb.write('Excel1.xlsx');
*/

wb.write('Excel.xlsx', function (err, stats) {
    if (err) {
        console.log(err);
    } else {
        console.log('Excel.xlsx written and has the following stats');
        console.log(stats);
    }
});

/*
var http = require('http');
http.createServer(function (req, res) {
    wb.write('MyExcel.xlsx', res);
}).listen(3000, function () {
    console.log('Go to http://localhost:3000 to download a copy of this Workbook');
});
*/
