import {  getHashOfPassword } from '../password';
import { expect } from 'chai';

describe('password utils', () => {
  describe('getHashOfPassword', () => {
    const testCases = [
      { str: 'password', passwd: '83AF' },
      { str: 'passw0rd', passwd: '946F' },
      { str: 'pa$sword', passwd: '8117' },
    ];

    testCases.forEach(t => {
      it(`should return ${t.passwd} for string: ${t.str}`, () => {
        expect(getHashOfPassword(t.str)).to.equal(t.passwd);
      });
    });
  });
});
