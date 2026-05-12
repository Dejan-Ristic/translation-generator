import ExcelJS from 'exceljs';
import { writeFileSync } from 'fs';
import {
  TranslationConstants,
  type TranslationData,
} from './types/TranslationData.types.ts';
import console from 'console';

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

      if (typeof transProp === 'string') {
        const nestedProps = transProp.split('.');
        let current = translatedObj;

        nestedProps.forEach((key, index) => {
          if (!current[key]) {
            if (index === nestedProps.length - 1) {
              if (typeof transVal === 'string') {
                const textFormat = row.getCell(translatedlangNumber).font;
                if (textFormat.bold || textFormat.italic)
                  return (current[key] = [
                    {
                      font: {
                        bold: textFormat.bold,
                        italic: textFormat.italic,
                      },
                      text: transVal,
                    },
                  ]);
                else return (current[key] = transVal);
              } else
                return (current[key] = transVal.richText.map((textEl) => {
                  const filteredText = {};
                  if (textEl.font.bold || textEl.font.italic)
                    filteredText.font = {
                      bold: textEl?.font?.bold,
                      italic: textEl?.font?.italic,
                    };
                  filteredText.text = textEl.text;
                  return filteredText;
                }));
            }
            current[key] = {};
          }
          current = current[key];
        });
      }
    }
  });

  const checkArrayFields = (transArr) => {
    transArr.forEach((arrEl) => {
      if (typeof arrEl === 'object') checkObjectFields(arrEl);
    });
  };

  const checkObjectFields = (transObj) => {
    Object.entries(transObj).forEach(([key, val]) => {
      if (typeof val === 'object' && !Array.isArray(val)) {
        const objectKeys = Object.keys(val);
        if (objectKeys.every((objSubkey) => !Number.isNaN(Number(objSubkey)))) {
          const orderedKeys = objectKeys.sort((a, b) => Number(a) - Number(b));
          const dataArray = [];
          orderedKeys.forEach((orderedKey) => dataArray.push(val[orderedKey]));
          transObj[key] = dataArray;
          checkArrayFields(transObj[key]);
        } else checkObjectFields(transObj[key]);
      }
    });
  };

  checkObjectFields(translatedObj);

  const jsonOutput = JSON.stringify(translatedObj, null, 2);
  writeFileSync(`translations/${langCode}.json`, jsonOutput);
};

const workbook = new ExcelJS.Workbook();

// TODO: ovo kasnije da se dinamicki uzme iz args
const translatedlang = 'ba';

// TODO: ovo kasnije da se dinamicki uzme iz args
workbook.xlsx
  .readFile(`translations/${TranslationConstants.XLSX_FILE}.xlsx`)
  .then(() => importFromExcel(translatedlang))
  .catch((error) => console.error('\x1b[31m%s\x1b[0m', error));
