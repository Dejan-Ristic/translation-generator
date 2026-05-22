import ExcelJS from 'exceljs';
import { writeFileSync } from 'fs';
import { RichTextPartial, type TranslationData } from './types';
import { ERROR_MESSAGES, TRANSLATION_PARAMS } from './constants';

const xlsxToJsonError = (errorMsg = '') => {
  throw new Error(errorMsg || ERROR_MESSAGES.XLSX_TO_JSON_ERROR_DEFAULT);
};

const importFromExcel = (langCode: string) => {
  const worksheet = workbook.getWorksheet(TRANSLATION_PARAMS.WORKBOOK_NAME);
  if (!worksheet) return;

  const translatedObj: TranslationData = {};

  let propertyPathNumber = 0;
  let translatedlangNumber = 0;
  let firstRowIndex: null | number = null;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (!firstRowIndex) firstRowIndex = rowNumber;

    if (rowNumber === firstRowIndex) {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.value === TRANSLATION_PARAMS.COLUMN_PROPERTY)
          propertyPathNumber = colNumber;
        if (cell.value === langCode) translatedlangNumber = colNumber;
      });
      if (!propertyPathNumber || !translatedlangNumber) xlsxToJsonError();
    } else {
      const transProp = row.getCell(propertyPathNumber).value;
      const transVal = row.getCell(translatedlangNumber).value as
        | string
        | RichTextPartial;

      if (!transProp || !transVal) xlsxToJsonError();

      if (typeof transProp === 'string') {
        const nestedProps = transProp.split('.');
        let current: TranslationData = translatedObj;

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
              } else if (transVal.richText && Array.isArray(transVal.richText))
                return (current[key] = transVal.richText.map((textEl) => {
                  const filteredText: RichTextPartial = {};
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
          current = current[key] as TranslationData;
        });
      }
    }
  });

  const checkArrayFields = (
    transArr: Array<string | TranslationData | RichTextPartial>
  ) => {
    transArr.forEach((arrEl) => {
      if (typeof arrEl === 'object')
        checkObjectFields(arrEl as TranslationData);
    });
  };

  const checkObjectFields = (transObj: TranslationData) => {
    Object.entries(transObj).forEach(([key, val]) => {
      if (typeof val === 'object' && !Array.isArray(val)) {
        const objectKeys = Object.keys(val);
        if (objectKeys.every((objSubkey) => !Number.isNaN(Number(objSubkey)))) {
          const orderedKeys = objectKeys.sort((a, b) => Number(a) - Number(b));
          const dataArray: Array<string | TranslationData | RichTextPartial> =
            [];
          orderedKeys.forEach((orderedKey) =>
            dataArray.push(
              val[orderedKey] as string | TranslationData | RichTextPartial
            )
          );
          transObj[key] = dataArray;
          checkArrayFields(transObj[key]);
        } else checkObjectFields(transObj[key] as TranslationData);
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
  .readFile(`translations/${TRANSLATION_PARAMS.XLSX_FILE}.xlsx`)
  .then(() => importFromExcel(translatedlang))
  .catch((error) => {
    console.error('\x1b[31m%s\x1b[0m', error);
    xlsxToJsonError(ERROR_MESSAGES.XLSX_TO_JSON_ERROR_READ_FILE);
  });
