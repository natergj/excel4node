require('source-map-support').install();
var xl = require('./distribution');
var wb = new xl.WorkBook({
    defaultFont: {
        name: 'Verdana',
        size: 12
    }
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

wb.write('Excel.xlsx', function (err, stats) {
    console.log(err);
    console.log(stats);
});
