import { writeFileSync } from 'fs';
import { parseArgs } from 'node:util';
import ExcelJS from 'exceljs';
import { type RichTextPartial, type TranslationData } from './types/index.ts';
import {
  ERROR_MESSAGES,
  PARAMS_CONFIG,
  TRANSLATION_PARAMS,
} from './constants/index.ts';
import { parseAsText } from './utils/index.ts';

const xlsxToJsonError = (errorMsg = '') => {
  throw new Error(errorMsg || ERROR_MESSAGES.XLSX_TO_JSON_ERROR_DEFAULT);
};

// parameters configuration
const paramsConfig = {
  options: {
    ...PARAMS_CONFIG.options,
    'trans-lang': { type: 'string' },
  },
} as const;

// script parameters init
const scriptParams = parseArgs(paramsConfig);
const sourceFile =
  scriptParams.values['source-file'] ??
  xlsxToJsonError(ERROR_MESSAGES.XLSX_TO_JSON_ERROR_SOURCE_FILE);
const translationLanguage =
  scriptParams.values['trans-lang'] ??
  xlsxToJsonError(ERROR_MESSAGES.XLSX_TO_JSON_ERROR_LANGUAGE);
const saveAsFile =
  scriptParams.values['save-as'] ?? scriptParams.values['trans-lang'];
const workbookName =
  scriptParams.values['wb-name'] ?? TRANSLATION_PARAMS.WORKBOOK_NAME;
const parseAs = [
  TRANSLATION_PARAMS.CONTENT_TEXT,
  TRANSLATION_PARAMS.CONTENT_RICHTEXT,
].includes(scriptParams.values['parse-as'] as string)
  ? scriptParams.values['parse-as']
  : TRANSLATION_PARAMS.CONTENT_TEXT;

const checkArrayFields = (
  transArr: Array<string | TranslationData | RichTextPartial>
) => {
  transArr.forEach((arrEl) => {
    if (typeof arrEl === 'object') checkObjectFields(arrEl as TranslationData);
  });
};

const checkObjectFields = (transObj: TranslationData) => {
  Object.entries(transObj).forEach(([key, val]) => {
    if (typeof val === 'object' && !Array.isArray(val)) {
      const objectKeys = Object.keys(val);
      if (objectKeys.every((objSubkey) => !Number.isNaN(Number(objSubkey)))) {
        const orderedKeys = objectKeys.sort((a, b) => Number(a) - Number(b));
        const dataArray: Array<string | TranslationData | RichTextPartial> = [];
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

const importFromExcel = (langCode: string) => {
  const worksheet = workbook.getWorksheet(workbookName);
  if (!worksheet)
    return xlsxToJsonError(ERROR_MESSAGES.XLSX_TO_JSON_ERROR_WORKBOOK);

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
      if (!propertyPathNumber || !translatedlangNumber)
        return xlsxToJsonError(ERROR_MESSAGES.XLSX_TO_JSON_ERROR_READ_FILE);
    } else {
      const transProp = row.getCell(propertyPathNumber).value;
      const transVal = row.getCell(translatedlangNumber).value as
        | string
        | RichTextPartial;

      if (!transProp || !transVal)
        return xlsxToJsonError(ERROR_MESSAGES.XLSX_TO_JSON_ERROR_READ_FILE);

      if (typeof transProp === 'string') {
        const nestedProps = transProp.split('.');
        let current: TranslationData = translatedObj;

        nestedProps.forEach((key, index) => {
          if (!current[key]) {
            if (index === nestedProps.length - 1) {
              if (typeof transVal === 'string') {
                const textFormat = row.getCell(translatedlangNumber).font;
                if (textFormat.bold || textFormat.italic) {
                  if (parseAs === TRANSLATION_PARAMS.CONTENT_RICHTEXT)
                    return (current[key] = [
                      {
                        font: {
                          bold: textFormat.bold,
                          italic: textFormat.italic,
                        },
                        text: transVal,
                      },
                    ]);
                  else if (parseAs === TRANSLATION_PARAMS.CONTENT_TEXT)
                    return (current[key] = transVal);
                } else return (current[key] = transVal);
              } else if (
                transVal.richText &&
                Array.isArray(transVal.richText)
              ) {
                if (parseAs === TRANSLATION_PARAMS.CONTENT_RICHTEXT) {
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
                } else if (parseAs === TRANSLATION_PARAMS.CONTENT_TEXT) {
                  return (current[key] = parseAsText(
                    transVal.richText as RichTextPartial[]
                  ));
                }
              }
            }
            current[key] = {};
          }
          current = current[key] as TranslationData;
        });
      }
    }
  });

  checkObjectFields(translatedObj);

  // saving parsed data to json file
  const jsonOutput = JSON.stringify(translatedObj, null, 2);
  writeFileSync(`${saveAsFile}.json`, jsonOutput);
};

const workbook = new ExcelJS.Workbook();
workbook.xlsx
  .readFile(`${sourceFile}.xlsx`)
  .then(() => importFromExcel(translationLanguage))
  .catch((error) => {
    console.error('\x1b[31m%s\x1b[0m', error);
    xlsxToJsonError(ERROR_MESSAGES.XLSX_TO_JSON_ERROR_READ_FILE);
  });
