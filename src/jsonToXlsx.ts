import { readFile } from 'fs/promises';
import { parseArgs } from 'node:util';
import ExcelJS from 'exceljs';
import {
  type RichTextPartial,
  type TranslationData,
  type TranslationValue,
} from './types/index.ts';
import {
  ERROR_MESSAGES,
  PARAMS_CONFIG,
  TRANSLATION_PARAMS,
} from './constants/index.ts';
import { checkIfRichTextContent, parseAsText } from './utils/index.ts';

const jsonToXlsxError = (errorMsg = '') => {
  throw new Error(errorMsg || ERROR_MESSAGES.JSON_TO_XLSX_ERROR_DEFAULT);
};

// script parameters init
const scriptParams = parseArgs(PARAMS_CONFIG);
const sourceFile =
  scriptParams.values['source-file'] ??
  jsonToXlsxError(ERROR_MESSAGES.JSON_TO_XLSX_ERROR_SOURCE_FILE);
const saveAsFile =
  scriptParams.values['save-as'] ?? TRANSLATION_PARAMS.XLSX_FILE;
const workbookName =
  scriptParams.values['wb-name'] ?? TRANSLATION_PARAMS.WORKBOOK_NAME;
const parseAs = [
  TRANSLATION_PARAMS.CONTENT_TEXT,
  TRANSLATION_PARAMS.CONTENT_RICHTEXT,
].includes(scriptParams.values['parse-as'] as string)
  ? scriptParams.values['parse-as']
  : TRANSLATION_PARAMS.CONTENT_TEXT;

const jsonToXlsx = (data: Buffer<ArrayBuffer>) => {
  // workbook and worksheet init
  const translationData = JSON.parse(data.toString());
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(workbookName);
  worksheet.columns = [
    { header: TRANSLATION_PARAMS.COLUMN_PROPERTY, key: 'prop', width: 60 },
    { header: sourceFile.split('/').pop(), key: 'val', width: 200 },
  ];

  const traverseTranslationEntries = (key: string, value: TranslationValue) => {
    // value is typeof string on bottom level and is written to worksheet row
    if (typeof value === 'string') {
      if (!value)
        return jsonToXlsxError(
          ERROR_MESSAGES.JSON_TO_XLSX_ERROR_CONTENT_MISSING
        );
      return worksheet.addRow({
        prop: key,
        val: value,
      });
    } else if (Array.isArray(value)) {
      // if data in array conforms to RichTextPartial interface
      if (
        checkIfRichTextContent(
          value as Array<string | TranslationData | RichTextPartial>
        )
      ) {
        let content;
        // parses as rich text or plain text
        if (parseAs === TRANSLATION_PARAMS.CONTENT_TEXT)
          content = parseAsText(value as RichTextPartial[]);
        if (parseAs === TRANSLATION_PARAMS.CONTENT_RICHTEXT)
          content = { richText: value };
        return worksheet.addRow({
          prop: key,
          val: content,
        });
      } else
        // recursively traverses the array as if it is object - indices as keys
        value.forEach((element, index) => {
          traverseTranslationEntries(
            `${key}.${index}`,
            element as TranslationValue
          );
        });
    } else if (typeof value === 'object') {
      // every value in iteration that is an object is recursively traversed,
      // updated value for property string passed
      Object.entries(value).map(([nestedKey, nestedValue]) =>
        traverseTranslationEntries(`${key}.${nestedKey}`, nestedValue)
      );
    }
  };

  // data object traversing on top level
  const translationEntries = Object.entries(translationData) as Array<
    [string, TranslationValue]
  >;
  translationEntries.forEach(([key, value]) =>
    traverseTranslationEntries(key, value)
  );

  // saving parsed data to excel file
  worksheet.getRow(1).font = { bold: true };
  const saveName = `${saveAsFile}.xlsx`;
  workbook.xlsx.writeFile(saveName);
};

readFile(`${sourceFile}.json`)
  .then((data) => jsonToXlsx(data))
  .catch((error) => {
    console.error('\x1b[31m%s\x1b[0m', error);
    jsonToXlsxError(ERROR_MESSAGES.JSON_TO_XLSX_ERROR_READ_FILE);
  });
