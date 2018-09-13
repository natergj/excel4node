let test = require('tape');
let xl = require('../source/index');

function testEmoji(t, wb, ws, cellIndex, strVal) {
    let cellAccessor = ws.cell(1, cellIndex);
    let cell = cellAccessor.string(strVal);
    let thisCell = ws.cells[cell.excelRefs[0]];
    t.ok(wb.sharedStrings[thisCell.v] === strVal, 'Emoji exists in cell');
}
test('Cell coverage', (t) => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('test');

    testEmoji(t, wb, ws, 1, '😂');
    testEmoji(t, wb, ws, 2, 'hello! 😂');
    testEmoji(t, wb, ws, 3, '😂☕️');

    t.end();
});