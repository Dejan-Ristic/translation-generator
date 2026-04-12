import ExcelJS from 'exceljs';
import { writeFileSync } from 'fs';

// TODO: ovo kasnije da se dinamicki uzme iz args
const translatedlang = 'en';

const translationError = (errorMsg = '') => {
  throw new Error(errorMsg || 'Translation not formatted');
};

async function convertColumnToJson() {
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile('./Translation-test.xlsx');

  const worksheet = workbook.getWorksheet('translation');
  if (!worksheet) return;

  const translatedObj: Record<string, string> = {};

  let propertyPathNumber = 0;
  let translatedlangNumber = 0;
  let firstRowIndex: null | number = null;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (!firstRowIndex) firstRowIndex = rowNumber;

    if (rowNumber === firstRowIndex) {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.value === 'propertyPath') propertyPathNumber = colNumber;
        if (cell.value === translatedlang) translatedlangNumber = colNumber;
      });
      if (!propertyPathNumber || !translatedlangNumber) translationError();
    } else {
      const transProp = row.getCell(propertyPathNumber).value;
      const transVal = row.getCell(translatedlangNumber).value;

      if (!transProp || !transVal) translationError();
      if (typeof transVal === 'string' && typeof transProp === 'string')
        translatedObj[transProp] = transVal;
    }
  });

  const jsonOutput = JSON.stringify(translatedObj, null, 2);
  writeFileSync(`${translatedlang}.json`, jsonOutput);
}

convertColumnToJson().catch((err) => console.error(err));
