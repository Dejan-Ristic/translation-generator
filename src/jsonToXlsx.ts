import { readFile } from 'fs/promises';
import { parseArgs } from 'node:util';
import ExcelJS from 'exceljs';
import type { TranslationData } from './types/TranslationData.types.ts';
import { htmlToRichText } from './utils/htmlToRichText.ts';

const exportError = (errorMsg = '') => {
  throw new Error(errorMsg || 'Translation export error');
};

const argsConfig = {
  options: {
    'source-file': { type: 'string' },
    'save-as': { type: 'string' },
    'parse-as': { type: 'string' as const, choices: ['text', 'html'] as const },
    'wb-name': { type: 'string' },
  },
} as const;

const scriptArgs = parseArgs(argsConfig);
const sourceFile =
  scriptArgs.values['source-file'] ?? exportError('no source .json file');
const saveAsFile = scriptArgs.values['save-as'] ?? 'translation';
const workbookName = scriptArgs.values['wb-name'] ?? 'translation';
const parseAs = scriptArgs.values['parse-as'] ?? 'text';

const jsonToXlsx = (data: Buffer<ArrayBuffer>) => {
  const translationData = JSON.parse(data.toString());

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(workbookName);

  worksheet.columns = [
    { header: 'propertyPath', key: 'prop', width: 40 },
    { header: sourceFile.split('/').pop(), key: 'val', width: 60 },
  ];

  const traverseTranslationEntries = (
    key: string,
    value: string | TranslationData
  ) => {
    if (typeof value === 'string') {
      let content: unknown = '';
      if (parseAs === 'text') content = value;
      if (parseAs === 'html') content = { richText: htmlToRichText(value) };
      if (content)
        return worksheet.addRow({
          prop: key,
          val: content,
        });
      else exportError('no translation content');
    }
    if (typeof value === 'object') {
      Object.entries(value).map(([nestedKey, nestedValue]) =>
        traverseTranslationEntries(`${key}.${nestedKey}`, nestedValue)
      );
    }
  };

  const translationEntries = Object.entries(translationData) as Array<
    [string, string | TranslationData]
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
