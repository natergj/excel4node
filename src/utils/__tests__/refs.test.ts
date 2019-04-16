import { generateRId } from '../refs';
import { expect } from 'chai';

describe('cellRef utils', () => {
  describe('generateRId', () => {
    it('should return a string of length 17', () => {
      expect(generateRId().length).to.equal(17);
    });

    it('should return a string beginning with R', () => {
      expect(generateRId().substr(0, 1)).to.equal('R');
    });
  });
});
