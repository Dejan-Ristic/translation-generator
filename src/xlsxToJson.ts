import ExcelJS from 'exceljs';
import { writeFileSync } from 'fs';
import {
  TranslationConstants,
  type TranslationData,
} from './types/TranslationData.types.ts';

const translationError = (errorMsg = '') => {
  throw new Error(errorMsg || 'Translation not formatted');
};

const importFromExcel = (langCode: string) => {
  const worksheet = workbook.getWorksheet(TranslationConstants.WORKBOOK_NAME);
  if (!worksheet) return;

  const translatedObj: TranslationData = {};

  let propertyPathNumber = 0;
  let translatedlangNumber = 0;
  let firstRowIndex: null | number = null;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (!firstRowIndex) firstRowIndex = rowNumber;

    if (rowNumber === firstRowIndex) {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.value === 'propertyPath') propertyPathNumber = colNumber;
        if (cell.value === langCode) translatedlangNumber = colNumber;
      });
      if (!propertyPathNumber || !translatedlangNumber) translationError();
    } else {
      const transProp = row.getCell(propertyPathNumber).value;
      const transVal = row.getCell(translatedlangNumber).value;

      if (!transProp || !transVal) translationError();

      if (typeof transProp === 'string' && typeof transVal === 'string') {
        const nestedProps = transProp.split('.');

        let current = translatedObj;
        nestedProps.forEach((key, index) => {
          if (!current[key])
            current[key] = index === nestedProps.length - 1 ? transVal : {};
          if (typeof current[key] !== 'string') current = current[key];
        });
      }
    }
  });

  const jsonOutput = JSON.stringify(translatedObj, null, 2);
  writeFileSync(`src/translations/${langCode}.json`, jsonOutput);
};

const workbook = new ExcelJS.Workbook();

// TODO: ovo kasnije da se dinamicki uzme iz args
const translatedlang = 'en';

// TODO: ovo kasnije da se dinamicki uzme iz args
workbook.xlsx
  .readFile('src/translations/Translation-test.xlsx')
  .then(() => importFromExcel(translatedlang))
  .catch((error) => console.error('\x1b[31m%s\x1b[0m', error));
