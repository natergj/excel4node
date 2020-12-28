const fs = require('fs');
const path = require('path');
const unzipper = require('unzipper')

require('source-map-support').install();
var xl = require('./distribution');

var wb = new xl.Workbook({
    defaultFont: {
        name: 'Verdana',
        size: 12
    },
    dateFormat: 'mm/dd/yyyy hh:mm:ss',
    logLevel: 1,
    workbookView: {
        windowWidth: 28800,
        windowHeight: 17620,
        xWindow: 240,
        yWindow: 480,
    },
    author: 'Mohanad Ahmed'
});

var ws = wb.addWorksheet("Sheet1", {
    pageSetup: {
        fitToWidth: 1,
        paperSize: 'A4_PAPER', 
        orientation: 'landscape'
    },
    headerFooter: {
        oddHeader: '&L&G &C&G &R&G',
        oddFooter: '&L&G &C&G &R&G'
    },
    printOptions: { centerHorizontal: true },
    sheetView: { rightToLeft: true }
});

ws.cell(1,1).string('Mohanad says hi');

ws.addHeaderFooterImage({
    image: fs.readFileSync(path.resolve(__dirname, './sampleFiles/logo copy.png')),
    type: 'picture'
}, 'LF')
ws.addHeaderFooterImage({
    image: fs.readFileSync(path.resolve(__dirname, './sampleFiles/logo copy.png')),
    type: 'picture'
}, 'CF')
ws.addHeaderFooterImage({
    image: fs.readFileSync(path.resolve(__dirname, './sampleFiles/signatures2.png')),
    type: 'picture',
    scale: 0.05
}, 'RF')
ws.addHeaderFooterImage({
    image: fs.readFileSync(path.resolve(__dirname, './sampleFiles/signatures2.png')),
    type: 'picture',
    scale: 0.23
}, 'CH')

wb.write('Excel.xlsx', function (err, stats) {
    if (err) {
        console.log(err);
    } else {
        console.log('Excel.xlsx written and has the following stats');
        console.log(stats);
        var dir = require('path').join(__dirname, '/testzip')
        fs.createReadStream('Excel.xlsx').pipe(unzipper.Extract({ path: dir }));
    }
});
