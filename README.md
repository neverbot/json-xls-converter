# json-xls-converter

[![npm](https://img.shields.io/npm/dt/json-xls-converter)](https://www.npmjs.com/package/json-xls-converter)
[![npm](https://img.shields.io/npm/dw/json-xls-converter)](https://www.npmjs.com/package/json-xls-converter)
[![GitHub license](https://img.shields.io/github/license/neverbot/json-xls-converter)](https://github.com/neverbot/json-xls-converter/blob/master/LICENSE)
[![npm](https://img.shields.io/npm/v/json-xls-converter)](https://www.npmjs.com/package/json-xls-converter)

**Utility to convert json to an Excel file.**

This is an updated version of [json2xls](https://github.com/rikkertkoppes/json2xls) which seems to be abandoned, and have some important vulnerabilities in its dependencies (some of them abandoned too).

This project is based in:
- [`json2xls`](https://github.com/rikkertkoppes/json2xls), by [Rikkert Koppes](https://github.com/rikkertkoppes).
- [`excel-export`](https://github.com/functionscope/Node-Excel-Export), by [functionscope](https://github.com/functionscope).

None of those projects included any kind of license, I'm distributing this new version with the MIT license.

## Installation

`npm i json-xls-converter`

## Usage 

To save as a file:

```javascript
import { converter } from 'json-xls-converter';
import fs from 'fs/promises';

const json = {
  foo: 'bar',
  qux: 'moo',
  poo: 123,
  stux: new Date()
}

const xls = await converter(json);

await fs.writeFile('data.xlsx', xls, 'binary');
```

Or use as an express middleware. It adds a convenience xls method to the response object to immediately output an excel file as a download.

```javascript
const jsonArr = [{
  foo: 'bar',
  qux: 'moo',
  poo: 123,
  stux: new Date()
},
{
  foo: 'bar',
  qux: 'moo',
  poo: 345,
  stux: new Date()
}];

app.use(converter.middleware);

app.get('/', (req, res) => {
  res.xls('data.xlsx', jsonArr);
});
```

## Migrating from json2xls to json-xls-converter

1. `npm remove json2xls && npm install json-xls-converter`

2. Change every import from:

   `import json2xls from 'json2xls';`

   to:

   `import { converter } from 'json-xls-converter';`

3. Change every invocation from:

   `const xls = json2xls(data);`

   to:

   `const xls = await converter(data);`

4. If you are using it as an express middleware, change from:

   `app.use(json2xls.middleware);`

   to:
   
   `app.use(converter.middleware);`





