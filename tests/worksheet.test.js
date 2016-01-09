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
    t.plan(2);
    var ws = makeWorkSheet();

    ws.addConditionalFormattingRule('A1:A10', {
        type: 'containsText',
        priority: 1,
        operator: 'containsText',
        text: '??',
        formula: 'NOT(ISERROR(SEARCH("??", A1)))'
    });

    ws.addConditionalFormattingRule('B1:B10', {
        type: 'containsText',
        priority: 2,
        operator: 'containsText',
        text: '??',
        formula: 'NOT(ISERROR(SEARCH("??", A1)))'
    });

    ws.addConditionalFormattingRule('B1:B10', {
        type: 'containsText',
        priority: 3,
        operator: 'containsText',
        text: '!!',
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

    // console.log(doc.prettyPrint());
});

