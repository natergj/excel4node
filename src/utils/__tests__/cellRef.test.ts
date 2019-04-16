import { getExcelAlpha, getExcelCellRef, getExcelRowCol } from '../cellRef';
import { expect } from 'chai';

describe('cellRef utils', () => {
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
        expect(getExcelRowCol(t.ref)).to.eql({ row: t.row, col: t.col });
      });
    });
  });

  describe('getExcelAlpha', () => {
    const testCases = [
      { col: 1, colStr: 'A' },
      { col: 27, colStr: 'AA' },
      { col: 677, colStr: 'ZA' },
      { col: 729, colStr: 'ABA' },
    ];
    testCases.forEach(t => {
      it(`should return column string ${t.colStr} for column ${t.col}`, () => {
        expect(getExcelAlpha(t.col)).to.equal(t.colStr);
      });
    });
  });

  describe('getExcelCellRef', () => {
    const testCases = [
      { row: 1, col: 1, ref: 'A1' },
      { row: 10, col: 3, ref: 'C10' },
      { row: 14, col: 27, ref: 'AA14' },
      { row: 999, col: 729, ref: 'ABA999' },
    ];
    testCases.forEach(t => {
      it(`should return excel ref of ${t.ref} for row ${t.row} and col ${t.col}`, () => {
        expect(getExcelCellRef(t.row, t.col)).to.equal(t.ref);
      });
    });
  });
});
