var test = require('tape');
var XmlTestDoc = require('./lib/xml_test_doc');

var xl = require('../lib/index');

function makeWorkSheet() {
    var wb = new xl.WorkBook();
    return wb.WorkSheet('test');
}

test('WorkSheet coverage', function (t) {
    t.plan(3);
    var ws = makeWorkSheet();
    t.ok(ws.Column(1));
    t.ok(ws.Row(1));
    t.ok(ws.Cell(1, 1));
});

test('WorkSheet setValidation()', function (t) {
    t.plan(1);
    var ws = makeWorkSheet();
    ws.setValidation({
        type: 'list',
        allowBlank: 1,
        sqref: 'B2:B10',
        formulas: ['=sheet2!$A$1:$A$2']
    });
    t.ok(ws);
});

test('WorkSheet addConditionalFormattingRule()', function (t) {
    t.plan(3);
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test');

    var style = wb.Style();
    style.Font.Bold();
    style.Fill.Color('FFDDDD');
    style.Fill.Pattern('solid');

    ws.addConditionalFormattingRule('A1:A10', {
        type: 'expression',
        priority: 1,
        formula: 'NOT(ISERROR(SEARCH("??", A1)))'
    });

    ws.addConditionalFormattingRule('B1:B10', {
        type: 'expression',
        priority: 2,
        formula: 'NOT(ISERROR(SEARCH("??", A1)))'
    });

    ws.addConditionalFormattingRule('B1:B10', {
        type: 'expression',
        priority: 3,
        formula: 'NOT(ISERROR(SEARCH("!!", B1)))'
    });

    var doc = new XmlTestDoc(ws.toXML());

    t.equal(
        doc.select('//conditionalFormatting/@sqref').length,
        2,
        'there should be two valid <conditionalFormatting/> tags created'
    );

    t.equal(
        doc.select('//conditionalFormatting[@sqref="B1:B10"]/cfRule').length,
        2,
        'there should be two rules for the sqref B1:B10'
    );

    var wbssDoc = new XmlTestDoc(wb.createStyleSheetXML());

    t.equal(
        wbssDoc.select('//styleSheet/dxfs').length,
        1,
        'there should be one <dxfs/> element in the workbook stylesheet output'
    );

    // console.log(doc.prettyPrint());
    // console.log(wbssDoc.prettyPrint());
});

