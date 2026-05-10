export const TranslationConstants = {
  WORKBOOK_NAME: 'arthrys_translations',
  XLSX_FILE: 'translations',
  TEXT: 'text',
  RICHTEXT: 'richtext',
};

export type TranslationData = {
  [key: string]: string | Array<TranslationData> | TranslationData;
};
