import { RICHTEXT_FIELDS } from '../constants';
import { RichTextPartial, TranslationData } from '../types';

export const checkIfRichTextContent = (
  contentArr: Array<string | TranslationData | RichTextPartial>
) => {
  return contentArr.every((arrItem) => {
    if (!arrItem || typeof arrItem !== 'object' || Array.isArray(arrItem))
      return false;

    const objKeys = Object.keys(arrItem);
    if (
      objKeys.length === 0 ||
      objKeys.length > 2 ||
      objKeys.some(
        (key) => ![RICHTEXT_FIELDS.TEXT, RICHTEXT_FIELDS.FONT].includes(key)
      )
    )
      return false;

    if (
      !Object.hasOwn(arrItem, RICHTEXT_FIELDS.TEXT) ||
      typeof arrItem[RICHTEXT_FIELDS.TEXT] !== 'string'
    )
      return false;

    if (
      Object.hasOwn(arrItem, RICHTEXT_FIELDS.FONT) &&
      (!arrItem[RICHTEXT_FIELDS.FONT] ||
        typeof arrItem[RICHTEXT_FIELDS.FONT] !== 'object' ||
        Array.isArray(arrItem[RICHTEXT_FIELDS.FONT]))
    )
      return false;

    return true;
  });
};

export const parseAsText = (contentArr: Array<RichTextPartial>) => {
  return contentArr.reduce(
    (accumulator, currentItem) =>
      `${accumulator}${currentItem[RICHTEXT_FIELDS.TEXT]}`,
    ''
  );
};
