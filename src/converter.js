/* eslint no-control-regex: 0 */

import nodeExcel from './excel.js';

const converter = async (json, config) => {
  const conf = converter.prepareJson(json, config);
  const result = await nodeExcel.execute(conf);
  return result;
};

// get a xls type based on js type
function getType(obj, type) {
  if (type) {
    return type;
  }
  const t = typeof obj;
  switch (t) {
    case 'string':
    case 'number':
      return t;
    case 'boolean':
      return 'bool';
    default:
      return 'string';
  }
}

// get a nested property from a JSON object given its key, i.e 'a.b.c'
function getByString(object, path) {
  path = path.replace(/\[(\w+)\]/g, '.$1'); // convert indexes to properties
  path = path.replace(/^\./, ''); // strip a leading dot
  const a = path.split('.');

  while (a.length) {
    const n = a.shift();

    if (n in object) {
      object = object[n] == undefined ? null : object[n];
    } else {
      return null;
    }
  }

  return object;
}

// prepare json to be in the correct format for excel-export
converter.prepareJson = (json, config) => {
  let res = {};
  const conf = config || {};
  const jsonArr = [].concat(json);
  let fields = conf.fields || Object.keys(jsonArr[0] || {});
  let types = [];

  if (!(fields instanceof Array)) {
    types = Object.keys(fields).map((key) => {
      return fields[key];
    });

    fields = Object.keys(fields);
  }

  // cols
  res.cols = fields.map((key, i) => {
    return {
      caption: key,
      type: getType(jsonArr[0][key], types[i]),
      beforeCellWrite: (row, cellData, eOpt) => {
        eOpt.cellType = getType(cellData, types[i]);
        return cellData;
      },
    };
  });

  // rows
  res.rows = jsonArr.map((row) => {
    return fields.map((key) => {
      let value = getByString(row, key);

      // stringify objects
      if (value && value.constructor == Object) {
        value = JSON.stringify(value);
      }
      // replace illegal xml characters with a square
      // see http://www.w3.org/TR/xml/#charsets
      // #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
      if (typeof value === 'string') {
        value = value.replace(
          /[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u10000-\u10FFFF]/g,
          ''
        );
      }

      return value;
    });
  });
  // add style xml if given
  if (conf.style) {
    res.stylesXmlFile = conf.style;
  }
  return res;
};

converter.middleware = (req, res, next) => {
  res.xls = async (fn, data, config) => {
    const xls = await converter(data, config);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader('Content-Disposition', 'attachment; filename=' + fn);
    res.end(xls, 'binary');
  };

  next();
};

export { converter };
