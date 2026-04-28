export enum TranslationConstants {
  ARTHRYS_CONTENT = 'arthrys_content',
  WORKBOOK_NAME = 'arthrys_translations',
  XLSX_FILE = 'translations',
}

export type TranslationData = {
  [key: string]: string | TranslationData;
};
