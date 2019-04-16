import { getExcelTS } from '../date';
import { expect } from 'chai';

describe('date utils', () => {
  describe('getExcelTS', () => {
    /**
     * Tests as defined in §18.17.4.3 of ECMA-376, Second Edition, Part 1 - Fundamentals And Markup Language Reference
     * The serial value 3687.4207639... represents 1910-02-03T10:05:54Z
     * The serial value 1.5000000... represents 1900-01-01T12:00:00Z
     * The serial value 2958465.9999884... represents 9999-12-31T23:59:59Z
     */
    const testCases = [
      { str: '1910-02-03T10:05:54Z', ts: 3687.4207639 },
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
        expect(getExcelTS(t.str)).to.equal(t.ts);
      });
    });
  });
});
