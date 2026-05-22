import { RICHTEXT_FIELDS, RICHTEXT_FONT_FIELDS } from '../constants';

export interface RichTextPartial {
  [RICHTEXT_FIELDS.TEXT]: string;
  [RICHTEXT_FIELDS.FONT]?: {
    [RICHTEXT_FONT_FIELDS.BOLD]?: boolean;
    [RICHTEXT_FONT_FIELDS.ITALIC]?: boolean;
  };
}

export type TranslationData = {
  [key: string]:
    | string
    | Array<string | TranslationData | RichTextPartial>
    | TranslationData;
};

export type TranslationValue = TranslationData[string];
