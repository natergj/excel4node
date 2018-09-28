let test = require('tape');
let xl = require('../source');

test('Escape Unicode Cell Values', (t) => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');
    let cellIndex = 1;
    /**
     * To test that unicode is escaped properly, provide an unescaped source string, and then our
     * expected escaped string.
     *
 * See the following literature:
 * https://stackoverflow.com/questions/43094662/excel-accepts-some-characters-whereas-openxml-has-error/43141040#43141040
 * https://stackoverflow.com/questions/43094662/excel-accepts-some-characters-whereas-openxml-has-error
 * https://www.ecma-international.org/publications/standards/Ecma-376.htm
     */
    function testUnicode(strVal, testVal) {
        let cellAccessor = ws.cell(1, cellIndex);
        let cell = cellAccessor.string(strVal);
        let thisCell = ws.cells[cell.excelRefs[0]];
        cellIndex++;
        t.ok(wb.sharedStrings[thisCell.v] === testVal, 'Unicode "' + strVal + '" correctly escaped in cell');
    }

    testUnicode('Hi <>', 'Hi &lt;&gt;');
    testUnicode('😂', '&#x1f602;');
    testUnicode('hello! 😂', 'hello! &#x1f602;');
    testUnicode('☕️', '☕️'); // ☕️ is U+2615 which is within the valid range.
    testUnicode('😂☕️', '&#x1f602;☕️');
    testUnicode('Good 🤞🏼 Luck', 'Good &#x1f91e;&#x1f3fc; Luck');
    testUnicode('Fist 🤜🏻🤛🏿 bump', 'Fist &#x1f91c;&#x1f3fb;&#x1f91b;&#x1f3ff; bump');
    testUnicode('㭩', '㭩');
    testUnicode('I am the Α and the Ω', 'I am the Α and the Ω');
    testUnicode('𐤶', '&#x10936;'); // Lydian Letter En U+10936
    testUnicode('𠁆', '&#x20046;'); // Ideograph bik6
    testUnicode('🚵', '&#x1f6b5'); // Mountain Bicyclist

    t.end();
});