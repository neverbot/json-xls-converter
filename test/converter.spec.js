import arrayData from './array-data.json' assert { type: 'json' };
import { converter } from '../src/converter.js';
import fs from 'fs/promises';
import { expect } from 'chai';

describe('Converter Tests', () => {
  describe('Creating an Excel file', () => {
    before(async () => {
      try {
        await fs.unlink('output.xlsx');
      } catch (e) {
        // do nothing
      }
    });

    it('should create a file', async () => {
      const xls = await converter(arrayData, {});

      await fs.writeFile('output.xlsx', xls, 'binary');

      try {
        let res = await fs.stat('output.xlsx');
        expect(res).to.be.an('object');
      } catch (e) {
        // this should not happen, only if the file does not exist
        expect(false).to.be.true;
      }
    });
  });
});
