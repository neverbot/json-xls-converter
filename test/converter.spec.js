import arrayData from './array-data.json' assert { type: 'json' };
import objectData from './object-data.json' assert { type: 'json' };
import weirdData from './weird-data.json' assert { type: 'json' };

import { transform } from '../src/converter.js';

import { expect } from 'chai';

function removeBeforeCellWrite(list) {
  return list.map((item) => {
    delete item.beforeCellWrite;
    return item;
  });
}

describe('converter tests', () => {
  /*
  beforeEach( () => {
    this.addMatchers({
      toEqualFields:  () => {
        return {
          compare: function (actual, expected) {
            const res =
              expected &&
              expected.all &&
              expected.all(function (item, i) {
                return (
                  actual[i] &&
                  Object.keys(item).all(function (field) {
                    return actual[i][field] === item[field];
                  })
                );
              });
            return {
              pass: res,
            };
          },
        };
      },
    });
  });
  */

  describe('handling illegal xml characters', () => {
    it('should remove vertical tabs', () => {
      const res = transform.prepareJson(weirdData);
      expect(res.rows[0][1]).to.equal(' foo bar ');
    });
  });

  describe('when the data is an empty array', () => {
    it('should create an empty config', () => {
      const res = transform.prepareJson([]);
      expect(res.cols).to.be.empty;
      expect(res.rows).to.be.empty;
    });
  });

  describe('when the data is an empty object', () => {
    it('should create a config with one empty row', () => {
      const res = transform.prepareJson({});
      expect(res.cols).to.be.empty;
      expect(res.rows).to.be.an('array');
      expect(res.rows[0]).to.be.empty;
    });
  });

  describe('when the data is an array', () => {
    describe('cols', () => {
      it('should create a cols part', () => {
        const res = transform.prepareJson(arrayData);
        expect(res.cols).to.exist;
      });

      it('should create the correct cols', () => {
        let res = transform.prepareJson(arrayData);

        res.cols = removeBeforeCellWrite(res.cols);

        expect(res.cols).to.include.deep.members([
          {
            caption: 'name',
            type: 'string',
          },
          {
            caption: 'date',
            type: 'string',
          },
          {
            caption: 'number',
            type: 'number',
          },
        ]);
      });

      it('should create the correct cols when fields are given as array', () => {
        let res = transform.prepareJson(arrayData, {
          fields: ['date', 'name'],
        });

        res.cols = removeBeforeCellWrite(res.cols);

        expect(res.cols).to.include.deep.members([
          {
            caption: 'date',
            type: 'string',
          },
          {
            caption: 'name',
            type: 'string',
          },
        ]);
      });

      it('should create the correct cols when fields are given as object', () => {
        let res = transform.prepareJson(arrayData, {
          fields: {
            number: 'string',
            name: 'string',
          },
        });

        res.cols = removeBeforeCellWrite(res.cols);

        expect(res.cols).to.include.deep.members([
          {
            caption: 'number',
            type: 'string',
          },
          {
            caption: 'name',
            type: 'string',
          },
        ]);
      });

      it('should create caption and type field', () => {
        const cols = transform.prepareJson(arrayData).cols;
        expect(cols[0].caption).to.exist;
        expect(cols[0].type).to.exist;
      });
    });

    describe('rows', () => {
      it('should create a rows part', () => {
        const res = transform.prepareJson(arrayData);
        expect(res.rows).to.exist;
      });

      it('should create rows with data in the correct order', () => {
        const res = transform.prepareJson(arrayData);
        expect(res.rows[0]).to.deep.equal(['Ivy Dickson', '2013-05-27T11:04:15-07:00', 10]);
        expect(res.rows[1]).to.deep.equal(['Walker Lynch', '2014-02-07T22:09:58-08:00', 2]);
        expect(res.rows[2]).to.deep.equal(['Maxwell U. Holden', '2013-06-16T05:29:13-07:00', 5]);
      });
    });

    describe('style', () => {
      it('should have the provided style xml file', () => {
        const fn = 'test.xml';
        const res = transform.prepareJson(arrayData, {
          style: fn,
        });

        expect(res.stylesXmlFile).to.equal(fn);
      });
    });
  });

  describe('when the data is an object', () => {
    describe('cols', () => {
      it('should create a cols part', () => {
        const res = transform.prepareJson(objectData);
        expect(res.cols).to.exist;
      });
      it('should create caption and type field', () => {
        var cols = transform.prepareJson(objectData).cols;
        expect(cols[0].caption).to.exist;
        expect(cols[0].type).to.exist;
      });
    });

    describe('rows', () => {
      it('should create a rows part', () => {
        const res = transform.prepareJson(objectData);
        expect(res.rows).to.exist;
      });
    });

    describe('style', () => {
      it('should have the provided style xml file', () => {
        var fn = 'test.xml';
        const res = transform.prepareJson(objectData, {
          style: fn,
        });
        expect(res.stylesXmlFile).to.equal(fn);
      });
    });

    describe('display of nested fields', () => {
      it('should write nested fields as json', () => {
        const res = transform.prepareJson(objectData);
        expect(res.rows[0][3]).to.equal('{"field":"foo"}');
      });
    });
  });

  describe('working with missing fields', () => {
    it('should leave missing fields blank', () => {
      const res = transform.prepareJson([
        {
          firma: 'transportabel',
          internet: 'http://www.transportabel.de',
          Branche: 'Möbel',
          STRASSE: 'Messingweg 49',
          ort: 'Münster-Sprakel',
          TEL_ZENTRALE: '(0251) 29 79 46',
        },
        {
          firma: 'Soziale Möbelbörse & mehr e.V.',
          internet: 'http://www.gersch-ms.de',
          Branche: 'Möbel',
          STRASSE: 'Nienkamp 80',
          ort: 'Münster-Wienburg',
          TEL_ZENTRALE: '(0251) 53 40 76',
        },
        {
          firma: 'Bald Eckhart e.K.',
          Branche: 'Möbel',
          STRASSE: 'Weseler Str. 628',
          ort: 'Münster-Mecklenbeck',
          TEL_ZENTRALE: '(0251) 53 40 76',
        },
      ]);

      expect(res.rows[2][1]).to.be.null;
    });
  });

  describe('prepping with config', () => {
    it('should get a nested field', () => {
      const res = transform.prepareJson(objectData, {
        fields: ['nested.field'],
      });

      expect(res.rows[0]).to.deep.equal(['foo']);
    });
  });
});
