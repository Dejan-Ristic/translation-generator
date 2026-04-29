export const TranslationConstants = {
  ARTHRYS_CONTENT: 'arthrys_content',
  WORKBOOK_NAME: 'arthrys_translations',
  XLSX_FILE: 'translations',
};

export type TranslationData = {
  [key: string]: string | Array<TranslationData> | TranslationData;
};
