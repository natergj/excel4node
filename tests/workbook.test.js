var test = require('tape');
var jszip = require('jszip');

var XmlTestDoc = require('./lib/xml_test_doc');

// load prototype extensions
// TODO fix prototype extensions and remove this
require('../lib/index');

var WorkBook = require('../lib/WorkBook');

test('WorkBook init', function (t) {
    t.plan(1);
    var wb = new WorkBook();
    t.ok(wb);
});

// Initial test to cover lib at a high level
test('WorkBook coverage', function (t) {
    t.plan(2);

    var wb = new WorkBook();

    var ws = wb.WorkSheet('Test Worksheet');

    var myCell = ws.Cell(1, 1);
    myCell.String('Test Value');

    var outBuffer = wb.writeToBuffer();

    t.ok(
        Buffer.isBuffer(outBuffer),
        'WorkBook#writeToBuffer() returns a Buffer'
    );

    var outZip = new jszip;
    outZip.load(outBuffer);

    var workbookXml = outZip.folder('xl').file('workbook.xml').asText();
    var doc = new XmlTestDoc(workbookXml);

    t.equal(
        doc.select('//workbook/sheets/sheet[@name="Test Worksheet"]').length,
        1,
        'XML output should have a valid <sheet/> tag'
    );

    // console.log(doc.prettyPrint());
});


test('WorkBook style sheet xml', function (t) {
    t.plan(5);

    var wb = new WorkBook();
    var ws = wb.WorkSheet('Test Worksheet');

    var styleSheetXml = wb.createStyleSheetXML();
    t.ok(styleSheetXml);

    var doc = new XmlTestDoc(styleSheetXml);

    ['fonts', 'fills', 'borders', 'cellXfs'].forEach(function (name) {
        t.equal(
            doc.select('//styleSheet/' + name).length,
            1,
            'XML output should have a valid <' + name + '/> tag'
        );
    });

    // console.log(doc.prettyPrint());
});

test('WorkBook default font', function (t) {
    t.plan(13);
    var wb = new WorkBook();
    var defaultFont = wb.styleData.fonts[0];
    t.equal(defaultFont.bold, false, 'Bold is false');
    t.equal(defaultFont.italics, false, 'Italics is false');
    t.equal(defaultFont.underline, false, 'Underline is false');
    t.equal(defaultFont.color, 'FF000000', 'Color is FF000000');
    t.equal(defaultFont.name, 'Calibri', 'Font name is Calibri');
    t.equal(defaultFont.sz, 12, 'Size 12');

    var df = wb.updateDefaultFont({
        bold: true,
        italics: true,
        underline: true,
        color: '0000FF',
        font: 'Arial',
        size: 10
    });

    t.equal(defaultFont, df, 'Return is default font');

    t.equal(defaultFont.bold, true, 'Bold is true');
    t.equal(defaultFont.italics, true, 'Italics is true');
    t.equal(defaultFont.underline, true, 'Underline is true');
    t.equal(defaultFont.color, '0000FF', 'Color is 0000FF');
    t.equal(defaultFont.name, 'Arial', 'Font name is Arial');
    t.equal(defaultFont.sz, 10, 'Size 10');
});