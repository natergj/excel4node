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

