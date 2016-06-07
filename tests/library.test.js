let test = require('tape');
let xl = require('../distribution/index');

test('Test library functions', (t) => {
    t.equals(xl.getExcelRowCol('A1').row, 1, 'Returned correct row from ref lookup');
    t.equals(xl.getExcelRowCol('C10').row, 10, 'Returned correct row from ref lookup');
    t.equals(xl.getExcelRowCol('AA14').row, 14, 'Returned correct row from ref lookup');
    t.equals(xl.getExcelRowCol('ABA999').row, 999, 'Returned correct row from ref lookup');
    t.equals(xl.getExcelRowCol('A1').col, 1, 'Returned correct column from ref lookup');
    t.equals(xl.getExcelRowCol('AA1').col, 27, 'Returned correct column from ref lookup');
    t.equals(xl.getExcelRowCol('ZA1').col, 677, 'Returned correct column from ref lookup');
    t.equals(xl.getExcelRowCol('ABA1').col, 729, 'Returned correct column from ref lookup');

    t.equals(xl.getExcelAlpha(1), 'A', 'Returned correct column alpha');
    t.equals(xl.getExcelAlpha(27), 'AA', 'Returned correct column alpha');
    t.equals(xl.getExcelAlpha(677), 'ZA', 'Returned correct column alpha');
    t.equals(xl.getExcelAlpha(729), 'ABA', 'Returned correct column alpha');

    t.equals(xl.getExcelCellRef(1, 1), 'A1', 'Returned correct excel cell reference');
    t.equals(xl.getExcelCellRef(10, 3), 'C10', 'Returned correct excel cell reference');
    t.equals(xl.getExcelCellRef(14, 27), 'AA14', 'Returned correct excel cell reference');
    t.equals(xl.getExcelCellRef(999, 729), 'ABA999', 'Returned correct excel cell reference');

    t.equals(xl.getExcelTS(new Date('2015-01-01T00:00:00.0000Z')), 42004.791666666664, 'Correctly translated date in standard time to excel timestamp');
    t.equals(xl.getExcelTS(new Date('2015-06-01T00:00:00.0000Z')), 42155.833333333336, 'Correctly translated date in daylight time to excel timestamp');

    t.end();
});