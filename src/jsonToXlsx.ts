import { readFile } from 'fs/promises';
import { parseArgs } from 'node:util';
import ExcelJS from 'exceljs';
import {
  type RichTextPartial,
  type TranslationData,
  type TranslationValue,
} from './types';
import { ERROR_MESSAGES, TRANSLATION_PARAMS } from './constants';
import { checkIfRichTextContent, parseAsText } from './utils';

const jsonToXlsxError = (errorMsg = '') => {
  throw new Error(errorMsg || ERROR_MESSAGES.JSON_TO_XLSX_ERROR_DEFAULT);
};

const argsConfig = {
  options: {
    'source-file': { type: 'string' },
    'save-as': { type: 'string' },
    'parse-as': {
      type: 'string',
      choices: [
        TRANSLATION_PARAMS.CONTENT_TEXT,
        TRANSLATION_PARAMS.CONTENT_RICHTEXT,
      ],
    },
    'wb-name': { type: 'string' },
  },
} as const;

const scriptArgs = parseArgs(argsConfig);
const sourceFile =
  scriptArgs.values['source-file'] ??
  jsonToXlsxError(ERROR_MESSAGES.JSON_TO_XLSX_ERROR_SOURCE_FILE);
const saveAsFile = scriptArgs.values['save-as'] ?? TRANSLATION_PARAMS.XLSX_FILE;
const workbookName =
  scriptArgs.values['wb-name'] ?? TRANSLATION_PARAMS.WORKBOOK_NAME;
const parseAs =
  scriptArgs.values['parse-as'] ?? TRANSLATION_PARAMS.CONTENT_TEXT;

const jsonToXlsx = (data: Buffer<ArrayBuffer>) => {
  const translationData = JSON.parse(data.toString());

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(workbookName);

  worksheet.columns = [
    { header: TRANSLATION_PARAMS.COLUMN_PROPERTY, key: 'prop', width: 60 },
    { header: sourceFile.split('/').pop(), key: 'val', width: 200 },
  ];

  const traverseTranslationEntries = (key: string, value: TranslationValue) => {
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
      if (
        checkIfRichTextContent(
          value as Array<string | TranslationData | RichTextPartial>
        )
      ) {
        let content;
        if (parseAs === TRANSLATION_PARAMS.CONTENT_TEXT)
          content = parseAsText(value as RichTextPartial[]);
        if (parseAs === TRANSLATION_PARAMS.CONTENT_RICHTEXT)
          content = { richText: value };
        return worksheet.addRow({
          prop: key,
          val: content,
        });
      } else
        value.forEach((element, index) => {
          traverseTranslationEntries(
            `${key}.${index}`,
            element as TranslationValue
          );
        });
    } else if (typeof value === 'object') {
      Object.entries(value).map(([nestedKey, nestedValue]) =>
        traverseTranslationEntries(`${key}.${nestedKey}`, nestedValue)
      );
    }
  };

  const translationEntries = Object.entries(translationData) as Array<
    [string, TranslationValue]
  >;

  translationEntries.forEach(([key, value]) =>
    traverseTranslationEntries(key, value)
  );

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
