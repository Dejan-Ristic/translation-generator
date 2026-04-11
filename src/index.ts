import ExcelJS from 'exceljs';
import { readFile } from 'fs/promises';
import type { TranslationData } from './types/translationData.ts';

const writeToExcel = (data: Buffer<ArrayBuffer>) => {
  const translationData = JSON.parse(data.toString());

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('User Data');

  worksheet.columns = [
    { header: 'Property', key: 'prop', width: 20 },
    { header: 'Value', key: 'val', width: 30 },
  ];

  const traverseTranslationObject = (
    key: string,
    value: string | TranslationData
  ) => {
    if (typeof value === 'string')
      return worksheet.addRow({ prop: key, val: value });
    if (typeof value === 'object') {
      Object.entries(value).map(([nestedKey, nestedValue]) =>
        traverseTranslationObject(`${key}.${nestedKey}`, nestedValue)
      );
    }
  };

  const translationEntries = Object.entries(translationData) as Array<
    [string, string | TranslationData]
  >;

  translationEntries.forEach(([key, value]) =>
    traverseTranslationObject(key, value)
  );

  worksheet.getRow(1).font = { bold: true };
  workbook.xlsx.writeFile('UserData.xlsx');
};

const dataPath = new URL('./test-data.json', import.meta.url);

readFile(dataPath)
  .then((data) => writeToExcel(data))
  .catch(() => {});
