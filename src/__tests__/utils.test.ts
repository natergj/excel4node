import utils from '../utils';
import { expect } from 'chai';

describe('Utils', () => {
  describe('generateRId', () => {
    it('should return a string of length 17', () => {
      expect(utils.generateRId().length).to.equal(17);
    });

    it('should return a string beginning with R', () => {
      expect(utils.generateRId().substr(0, 1)).to.equal('R');
    });
  });

  describe('getExcelRowCol', () => {
    const testCases = [
      { ref: 'A1', row: 1, col: 1 },
      { ref: 'C10', row: 10, col: 3 },
      { ref: 'AC14', row: 14, col: 29 },
      { ref: 'AQA999', row: 999, col: 1119 },
      { ref: 'AA12', row: 12, col: 27 },
      { ref: 'ZA50', row: 50, col: 677 },
      { ref: 'ABA121', row: 121, col: 729 },
    ];
    testCases.forEach(t => {
      it(`should return row: ${t.row}, col: ${t.col} for ref: ${t.ref}`, () => {
        expect(utils.getExcelRowCol(t.ref)).to.eql({ row: t.row, col: t.col });
      });
    });
  });

  describe('getExcelAlpha', () => {
    // prettier-ignore
    const testCases = [
      { col: 1, colStr: 'A' },
      { col: 27, colStr: 'AA' },
      { col: 677, colStr: 'ZA' },
      { col: 729, colStr: 'ABA' },
    ];
    testCases.forEach(t => {
      it(`should return column string ${t.colStr} for column ${t.col}`, () => {
        expect(utils.getExcelAlpha(t.col)).to.equal(t.colStr);
      });
    });
  });

  describe('getExcelCellRef', () => {
    // prettier-ignore
    const testCases = [
      { row: 1, col: 1, ref: 'A1'},
      { row: 10, col: 3, ref: 'C10'},
      { row: 14, col: 27, ref: 'AA14'},
      { row: 999, col: 729, ref: 'ABA999'},
    ];
    testCases.forEach(t => {
      it(`should return excel ref of ${t.ref} for row ${t.row} and col ${t.col}`, () => {
        expect(utils.getExcelCellRef(t.row, t.col)).to.equal(t.ref);
      });
    });
  });

  describe('getExcelTS', () => {
    /**
     * Tests as defined in §18.17.4.3 of ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference
     * The serial value 3687.4207639... represents 1910-02-03T10:05:54Z
     * The serial value 1.5000000... represents 1900-01-01T12:00:00Z
     * The serial value 2958465.9999884... represents 9999-12-31T23:59:59Z
     */
    // prettier-ignore
    const testCases = [
      { str: '1910-02-03T10:05:54Z', ts: 3687.4207639},
      { str: '1900-01-01T12:00:00Z', ts: 1.5 },
      { str: '9999-12-31T23:59:59Z', ts: 2958465.9999884 },
      { str: '1900-01-01T00:00:00Z', ts: 1 },
      { str: '1910-02-03T00:00:00Z', ts: 3687 },
      { str: '2006-02-01T00:00:00Z', ts: 38749 },
      { str: '9999-12-31T00:00:00Z', ts: 2958465 },
      { str: '2017-06-01T00:00:00.000Z', ts: 42887 },
    ];
    testCases.forEach(t => {
      it(`should return ${t.ts} for date string: ${t.str}`, () => {
        expect(utils.getExcelTS(t.str)).to.equal(t.ts);
      });
    });
  });
});
