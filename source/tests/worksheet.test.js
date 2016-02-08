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
    t.plan(5);
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test');

    var style = wb.Style();
    style.Font.Bold();
    style.Fill.Color('FFDDDD');
    style.Fill.Pattern('solid');

    ws.addConditionalFormattingRule('A1:A10', {
        type: 'expression',
        priority: 1,
        formula: 'NOT(ISERROR(SEARCH("??", A1)))',
        style: style
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
        doc.count('//conditionalFormatting/@sqref'),
        2,
        'there should be two valid <conditionalFormatting/> tags created'
    );

    t.equal(
        doc.count('//conditionalFormatting[@sqref="B1:B10"]/cfRule'),
        2,
        'there should be two rules for the sqref B1:B10'
    );

    var wbssDoc = new XmlTestDoc(wb.createStyleSheetXML());

    t.equal(
        wbssDoc.count('//styleSheet/dxfs'),
        1,
        'there should be one <dxfs/> element in the workbook stylesheet output'
    );

    t.equal(
        wbssDoc.count('//styleSheet/dxfs/dxf/font/color[@rgb="FF000000"]'),
        1,
        'there should be font color embedded in dxf style'
    );

    t.equal(
        wbssDoc.count('//styleSheet/dxfs/dxf/fill/patternFill/bgColor[@rgb="FFFFDDDD"]'),
        1,
        'there should be font color embedded in dxf style'
    );

    // console.log(doc.prettyPrint());
    // console.log(wbssDoc.prettyPrint());
});

test('WorkSheet test printScaling', function (t) {
    t.plan(15);
    var wb = new xl.WorkBook();
    var ws = wb.WorkSheet('test', {
        fitToPage: {
            orientation: 'landscape'
        }
    });
    var opts = {
        fitToWidth: 200,
        fitToHeight: 300,
        horizontalDpi: 12345,
        verticalDpi: 67890
    };

    t.equal(JSON.stringify(ws.printScaling(wb.Print.NO_SCALING)), '{"scale":0}', 'NO_SCALING') ;
    t.equal(ws.sheet.sheetPr.pageSetUpPr['@fitToPage'], undefined);
    t.equal(JSON.stringify(ws.printScaling(wb.Print.FIT_ONE_PAGE)), '{"scale":1}', 'FIT_ONE_PAGE') ;
    t.equal(ws.sheet.sheetPr.pageSetUpPr['@fitToPage'], 1);
    t.equal(JSON.stringify(ws.printScaling(wb.Print.FIT_ALL_COLUMNS)), '{"scale":2,"fitToHeight":0}', 'FIT_ALL_COLUMNS') ;
    t.equal(ws.sheet.sheetPr.pageSetUpPr['@fitToPage'], 1);
    t.equal(JSON.stringify(ws.printScaling(wb.Print.FIT_ALL_ROWS)), '{"scale":3,"fitToWidth":0}', 'FIT_ALL_ROWS') ;
    t.equal(ws.sheet.sheetPr.pageSetUpPr['@fitToPage'], 1);
    t.equal(JSON.stringify(ws.printScaling(wb.Print.CUSTOM_SCALING)), '{"scale":4}', 'CUSTOM_SCALING') ;
    t.equal(ws.sheet.sheetPr.pageSetUpPr['@fitToPage'], 1);

    opts.scale = wb.Print.NO_SCALING;
    ws.printScaling(JSON.parse(JSON.stringify(opts)));
    t.equal(JSON.stringify(ws.sheet.pageSetup), '[{"@orientation":"landscape"}]', 'NO_SCALING') ;

    opts.scale = wb.Print.FIT_ONE_PAGE;
    ws.printScaling(JSON.parse(JSON.stringify(opts)));
    t.equal(JSON.stringify(ws.sheet.pageSetup), '[{"@horizontalDpi":12345},{"@verticalDpi":67890},{"@orientation":"landscape"}]', 'FIT_ONE_PAGE') ;

    opts.scale = wb.Print.FIT_ALL_COLUMNS;
    ws.printScaling(JSON.parse(JSON.stringify(opts)));
    t.equal(JSON.stringify(ws.sheet.pageSetup), '[{"@fitToHeight":0},{"@horizontalDpi":12345},{"@verticalDpi":67890},{"@orientation":"landscape"}]', 'FIT_ALL_COLUMNS') ;

    opts.scale = wb.Print.FIT_ALL_ROWS;
    ws.printScaling(JSON.parse(JSON.stringify(opts)));
    t.equal(JSON.stringify(ws.sheet.pageSetup), '[{"@fitToWidth":0},{"@horizontalDpi":12345},{"@verticalDpi":67890},{"@orientation":"landscape"}]', 'FIT_ALL_ROWS') ;

    opts.scale = wb.Print.CUSTOM_SCALING;
    ws.printScaling(JSON.parse(JSON.stringify(opts)));
    t.equal(JSON.stringify(ws.sheet.pageSetup), '[{"@fitToHeight":300},{"@fitToWidth":200},{"@horizontalDpi":12345},{"@verticalDpi":67890},{"@orientation":"landscape"}]', 'CUSTOM_SCALING') ;
});

test('WorkSheet headerFooter()', function (t) {
    t.plan(1);
    var ws = makeWorkSheet();
    var headerFooter = ws.headerFooter({
        oddHeader: '&LDavid Gofman&R&D',
        oddFooter: '&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N'
    });
    t.equal(JSON.stringify(headerFooter), 
    '{' +
        '"@differentOddEven":false,' +
        '"@differentFirst":false,' +
        '"@scaleWithDoc":true,' +
        '"@alignWithMargins":true,' +
        '"oddHeader":"&LDavid Gofman&R&D",' +
        '"oddFooter":"&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N"' +
    '}');
});