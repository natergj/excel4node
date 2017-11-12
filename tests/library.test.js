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

    /**
     * Tests as defined in §18.17.4.3 of ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference
     * The serial value 3687.4207639... represents 1910-02-03T10:05:54Z
     * The serial value 1.5000000... represents 1900-01-01T12:00:00Z
     * The serial value 2958465.9999884... represents 9999-12-31T23:59:59Z
     */
    t.equals(xl.getExcelTS(new Date('1910-02-03T10:05:54Z')), 3687.4207639, 'Correctly translated date 1910-02-03T10:05:54Z');
    t.equals(xl.getExcelTS(new Date('1900-01-01T12:00:00Z')), 1.5000000, 'Correctly translated date 1900-01-01T12:00:00Z');
    t.equals(xl.getExcelTS(new Date('9999-12-31T23:59:59Z')), 2958465.9999884, 'Correctly translated date 9999-12-31T23:59:59Z');

    /**
     * Tests as defined in §18.17.4.1 of ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference
     * The serial value 2.0000000... represents 1900-01-01
     * The serial value 3687.0000000... represents 1910-02-03
     * The serial value 38749.0000000... represents 2006-02-01
     * The serial value 2958465.0000000... represents 9999-12-31
     */
    t.equals(xl.getExcelTS(new Date('1900-01-01T00:00:00Z')), 1, 'Correctly translated 1900-01-01');
    t.equals(xl.getExcelTS(new Date('1910-02-03T00:00:00Z')), 3687, 'Correctly translated 1910-02-03');
    t.equals(xl.getExcelTS(new Date('2006-02-01T00:00:00Z')), 38749, 'Correctly translated 2006-02-01');
    t.equals(xl.getExcelTS(new Date('9999-12-31T00:00:00Z')), 2958465, 'Correctly translated 9999-12-31');

    t.equals(xl.getExcelTS(new Date('2017-06-01T00:00:00.000Z')), 42887, 'Correctly translated 2017-06-01');

    t.end();
});