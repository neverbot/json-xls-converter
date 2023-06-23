import { expect } from 'chai';
import nodeExcel from '../src/excel.js';
import fs from 'fs/promises';

import { oaDate } from '../src/dates.js';

describe('Simple Excel xlsx Export', () => {
  describe('Export', () => {
    it('returns xlsx', async () => {
      let conf = {};
      conf.name = 'xxxxxx';
      conf.cols = [
        {
          caption: 'string',
          type: 'string',
        },
        {
          caption: 'date',
          type: 'date',
        },
        {
          caption: 'bool',
          type: 'bool',
        },
        {
          caption: 'number 2',
          type: 'number',
        },
      ];

      conf.rows = [
        ['pi', oaDate(new Date(Date.UTC(2013, 4, 1))), true, 3.14],
        ['e', oaDate(new Date(2012, 4, 1)), false, 2.7182],
        ["M&M<>'", oaDate(new Date(Date.UTC(2013, 6, 9))), false, 1.2],
        ['null', null, null, null],
      ];

      const result = await nodeExcel.execute(conf);
      //console.log(result);

      await fs.writeFile('single.xlsx', result, 'binary');

      try {
        let res = await fs.stat('single.xlsx');
        expect(res).to.be.an('object');
      } catch (e) {
        // this should not happen, only if the file does not exist
        expect(false).to.be.true;
      }
    });

    it('returns multisheet xlsx', async () => {
      let confs = [];
      let conf = {};
      conf.cols = [
        {
          caption: 'string',
          type: 'string',
        },
        {
          caption: 'date',
          type: 'date',
        },
        {
          caption: 'bool',
          type: 'bool',
        },
        {
          caption: 'number 2',
          type: 'number',
        },
      ];

      conf.rows = [
        ['hahai', oaDate(new Date(Date.UTC(2013, 4, 1))), true, 3.14],
        ['e', oaDate(new Date(2012, 4, 1)), false, 2.7182],
        ["M&M<>'", oaDate(new Date(Date.UTC(2013, 6, 9))), false, 1.2],
        ['null', null, null, null],
      ];

      for (var i = 0; i < 3; i++) {
        conf = JSON.parse(JSON.stringify(conf)); //clone
        conf.name = 'sheet' + i;
        confs.push(conf);
      }

      const result = await nodeExcel.execute(confs);

      await fs.writeFile('multi.xlsx', result, 'binary');

      try {
        let res = await fs.stat('multi.xlsx');
        expect(res).to.be.an('object');
      } catch (e) {
        // this should not happen, only if the file does not exist
        expect(false).to.be.true;
      }
    });
  });
});
