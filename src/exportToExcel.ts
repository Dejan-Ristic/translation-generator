import { readFile } from 'fs/promises';
import { parseArgs } from 'node:util';
import ExcelJS from 'exceljs';
import type { TranslationData } from './types/TranslationData.types.ts';

const argsConfig = {
  options: {
    'source-file': { type: 'string' },
    'save-as': { type: 'string' },
  },
} as const;

const scriptArgs = parseArgs(argsConfig);
const sourceLang = scriptArgs.values['source-file'];
const saveAsFile = scriptArgs.values['save-as'];

const exportToExcel = (data: Buffer<ArrayBuffer>) => {
  const translationData = JSON.parse(data.toString());

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('translation');

  worksheet.columns = [
    { header: 'propertyPath', key: 'prop', width: 40 },
    { header: sourceLang, key: 'val', width: 60 },
  ];

  const traverseTranslationEntries = (
    key: string,
    value: string | TranslationData
  ) => {
    if (typeof value === 'string')
      return worksheet.addRow({ prop: key, val: value });
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

const exportFrom = `${sourceLang}.json`;
const dataPath = new URL(exportFrom, import.meta.url);

readFile(dataPath)
  .then((data) => exportToExcel(data))
  .catch((error) => console.error('\x1b[31m%s\x1b[0m', error));
