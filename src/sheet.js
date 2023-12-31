/* eslint no-useless-escape: 0 */

import fs from 'fs/promises';

let sheetFront =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  ' <x:sheetPr/><x:sheetViews><x:sheetView tabSelected="1" workbookViewId="0" /></x:sheetViews>' +
  ' <x:sheetFormatPr defaultRowHeight="15" />';
const sheetBack =
  ' <x:pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75" />' +
  ' <x:headerFooter /></x:worksheet>';

class Sheet {
  constructor(config, xlsx, shareStrings, convertedShareStrings) {
    this.config = config;
    this.xlsx = xlsx;
    this.shareStrings = shareStrings;
    this.convertedShareStrings = convertedShareStrings;
  }

  async generate() {
    let cols = this.config.cols,
      data = this.config.rows,
      colsLength = cols.length,
      rows = '',
      row = '',
      colsWidth = '',
      styleIndex,
      k;

    this.config.fileName =
      'xl/worksheets/' + (this.config.name || 'sheet').replace(/[*?\]\[\/\/]/g, '') + '.xml';

    if (this.config.stylesXmlFile) {
      let path = this.config.stylesXmlFile;
      let styles = await fs.readFile(path, 'utf8');

      if (styles) {
        this.xlsx.file('xl/styles.xml', styles);
      }
    }

    // first row for column caption
    row = '<x:row r="1" spans="1:' + colsLength + '">';
    let colStyleIndex;

    for (k = 0; k < colsLength; k++) {
      colStyleIndex = cols[k].captionStyleIndex || 0;

      row += this.addStringCell(this.getColumnLetter(k + 1) + 1, cols[k].caption, colStyleIndex);

      if (cols[k].width) {
        colsWidth +=
          '<col customWidth = "1" width="' +
          cols[k].width +
          '" max = "' +
          (k + 1) +
          '" min="' +
          (k + 1) +
          '"/>';
      }
    }

    row += '</x:row>';
    rows += row;

    // fill in data
    let i,
      j,
      r,
      cellData,
      currRow,
      cellType,
      dataLength = data.length;

    for (i = 0; i < dataLength; i++) {
      (r = data[i]), (currRow = i + 2);
      row = '<x:row r="' + currRow + '" spans="1:' + colsLength + '">';

      for (j = 0; j < colsLength; j++) {
        styleIndex = null;
        cellData = r[j];
        cellType = cols[j].type;

        if (typeof cols[j].beforeCellWrite === 'function') {
          let e = {
            rowNum: currRow,
            styleIndex: null,
            cellType: cellType,
          };
          cellData = cols[j].beforeCellWrite(r, cellData, e);
          styleIndex = e.styleIndex || styleIndex;
          cellType = e.cellType;
          // Parsing error: Deleting local variable in strict mode.
          // delete e;
          e = null;
        }

        switch (cellType) {
          case 'number':
            row += this.addNumberCell(this.getColumnLetter(j + 1) + currRow, cellData, styleIndex);
            break;
          case 'date':
            row += this.addDateCell(this.getColumnLetter(j + 1) + currRow, cellData, styleIndex);
            break;
          case 'bool':
            row += this.addBoolCell(this.getColumnLetter(j + 1) + currRow, cellData, styleIndex);
            break;
          default:
            row += this.addStringCell(this.getColumnLetter(j + 1) + currRow, cellData, styleIndex);
        }
      }
      row += '</x:row>';
      rows += row;
    }

    if (colsWidth !== '') {
      sheetFront += '<cols>' + colsWidth + '</cols>';
    }

    this.xlsx.file(
      this.config.fileName,
      sheetFront + '<x:sheetData>' + rows + '</x:sheetData>' + sheetBack,
    );
  }

  addNumberCell(cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;

    if (value === null) {
      return '';
    } else
      return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
  }

  addDateCell(cellRef, value, styleIndex) {
    styleIndex = styleIndex || 1;

    if (value === null) {
      return '';
    } else
      return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
  }

  addBoolCell(cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;

    if (value === null) {
      return '';
    }
    if (value) {
      value = 1;
    } else {
      value = 0;
    }

    return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="b"><x:v>' + value + '</x:v></x:c>';
  }

  addStringCell(cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;

    if (value === null) {
      return '';
    }

    if (typeof value === 'string') {
      value = value
        .replace(/&/g, '&amp;')
        .replace(/'/g, '&apos;')
        .replace(/>/g, '&gt;')
        .replace(/</g, '&lt;');
    }

    let i = this.shareStrings.getElementByKey(value);

    if (typeof i === 'undefined') {
      i = this.shareStrings.length;
      this.shareStrings.setElement(value, i);
      this.convertedShareStrings += '<x:si><x:t>' + value + '</x:t></x:si>';
    }

    return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="s"><x:v>' + i + '</x:v></x:c>';
  }

  getColumnLetter(col) {
    if (col <= 0) {
      throw 'col must be more than 0';
    }

    let array = new Array();

    while (col > 0) {
      let remainder = col % 26;
      col /= 26;
      col = Math.floor(col);

      if (remainder === 0) {
        remainder = 26;
        col--;
      }

      array.push(64 + remainder);
    }

    return String.fromCharCode.apply(null, array.reverse());
  }
}

/*
let startTag = function (obj, tagName, closed) {
  let result = '<' + tagName,
    p;
  for (p in obj) {
    result += ' ' + p + '=' + obj[p];
  }
  if (!closed) result += '>';
  else result += '/>';
  return result;
};

let endTag = function (tagName) {
  return '</' + tagName + '>';
};
*/

export default Sheet;
