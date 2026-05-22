export const TRANSLATION_PARAMS: Record<string, string> = Object.freeze({
  WORKBOOK_NAME: 'arthrys translations',
  XLSX_FILE: 'translations',
  CONTENT_TEXT: 'text',
  CONTENT_RICHTEXT: 'richtext',
  COLUMN_PROPERTY: 'property',
});

export const RICHTEXT_FIELDS: Record<string, string> = Object.freeze({
  TEXT: 'text',
  FONT: 'font',
});

export const RICHTEXT_FONT_FIELDS: Record<string, string> = Object.freeze({
  BOLD: 'bold',
  ITALIC: 'italic',
});

export const ERROR_MESSAGES: Record<string, string> = Object.freeze({
  JSON_TO_XLSX_ERROR_DEFAULT: 'Translation excel file error',
  JSON_TO_XLSX_ERROR_SOURCE_FILE: 'Source json file not specified',
  JSON_TO_XLSX_ERROR_CONTENT_MISSING: 'Translation content missing',
  JSON_TO_XLSX_ERROR_READ_FILE: 'Source json read file error',
  XLSX_TO_JSON_ERROR_DEFAULT: 'Translation json file error',
  XLSX_TO_JSON_ERROR_READ_FILE: 'Source xlsx read file error',
});
