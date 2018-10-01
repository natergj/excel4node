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

    testUnicode('Hi <>', 'Hi <>');
    testUnicode('😂', '😂');
    testUnicode('hello! 😂', 'hello! 😂');
    testUnicode('☕️', '☕️'); // ☕️ is U+2615 which is within the valid range.
    testUnicode('😂☕️', '😂☕️');
    testUnicode('Good 🤞🏼 Luck', 'Good 🤞🏼 Luck');
    testUnicode('Fist 🤜🏻🤛🏿 bump', 'Fist 🤜🏻🤛🏿 bump');
    testUnicode('㭩', '㭩');
    testUnicode('I am the Α and the Ω', 'I am the Α and the Ω');
    testUnicode('𐤶', '𐤶'); // Lydian Letter En U+10936
    testUnicode('𠁆', '𠁆'); // Ideograph bik6
    testUnicode('\u000b', ''); // tab should be removed

    t.end();
});