import { readFile } from 'fs/promises';
import { parseArgs } from 'node:util';
import ExcelJS from 'exceljs';
import {
  TranslationConstants,
  type TranslationData,
} from './types/TranslationData.types.ts';

const exportError = (errorMsg = '') => {
  throw new Error(errorMsg || 'Translation export error');
};

const argsConfig = {
  options: {
    'source-file': { type: 'string' },
    'save-as': { type: 'string' },
    'parse-as': {
      type: 'string' as const,
      choices: ['text', 'richText'] as const,
    },
    'wb-name': { type: 'string' },
  },
} as const;

const scriptArgs = parseArgs(argsConfig);
const sourceFile =
  scriptArgs.values['source-file'] ?? exportError('no source .json file');
const saveAsFile =
  scriptArgs.values['save-as'] ?? TranslationConstants.XLSX_FILE;
const workbookName =
  scriptArgs.values['wb-name'] ?? TranslationConstants.WORKBOOK_NAME;
const parseAs = scriptArgs.values['parse-as'] ?? 'text';

const jsonToXlsx = (data: Buffer<ArrayBuffer>) => {
  const translationData = JSON.parse(data.toString());

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(workbookName);

  worksheet.columns = [
    { header: 'propertyPath', key: 'prop', width: 60 },
    { header: sourceFile.split('/').pop(), key: 'val', width: 200 },
  ];

  const traverseTranslationEntries = (
    key: string,
    value: string | Array<string | TranslationData> | TranslationData
  ) => {
    if (typeof value === 'string') {
      if (!value) return exportError('no translation content');
      if (parseAs === 'text')
        return worksheet.addRow({
          prop: key,
          val: value,
        });
      else exportError();
    } else if (Array.isArray(value)) {
      if (value[0] === TranslationConstants.ARTHRYS_CONTENT) {
        return worksheet.addRow({
          prop: key,
          val: { richText: value.slice(1) },
        });
      } else
        value.forEach((element, index) => {
          traverseTranslationEntries(`${key}.${index}`, element);
        });
    } else if (typeof value === 'object') {
      Object.entries(value).map(([nestedKey, nestedValue]) =>
        traverseTranslationEntries(`${key}.${nestedKey}`, nestedValue)
      );
    }
  };

  const translationEntries = Object.entries(translationData) as Array<
    [string, string | Array<TranslationData> | TranslationData]
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
    exportError('.json import error');
  });
