import { RICHTEXT_FIELDS } from '../constants/index.ts';
import { type RichTextPartial, type TranslationData } from '../types/index.ts';

// checks if every element of the contentArr array conforms to RichTextPartial interface
export const checkIfRichTextContent = (
  contentArr: Array<string | TranslationData | RichTextPartial>
) => {
  return contentArr.every((arrItem) => {
    // must be (not array) object
    if (!arrItem || typeof arrItem !== 'object' || Array.isArray(arrItem))
      return false;

    const objKeys = Object.keys(arrItem);

    // must have 1 or 2 keys, all are 'text' or 'font'
    if (
      objKeys.length === 0 ||
      objKeys.length > 2 ||
      objKeys.some(
        (key) => ![RICHTEXT_FIELDS.TEXT, RICHTEXT_FIELDS.FONT].includes(key)
      )
    )
      return false;

    // must have 'text' key and its value is a string
    if (
      !Object.hasOwn(arrItem, RICHTEXT_FIELDS.TEXT) ||
      typeof arrItem[RICHTEXT_FIELDS.TEXT] !== 'string'
    )
      return false;

    // can have 'font' key and it is (not array) object
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

// creates text only out of a rich text data array
export const parseAsText = (contentArr: Array<RichTextPartial>) => {
  return contentArr.reduce(
    (accumulator, currentItem) =>
      `${accumulator}${currentItem[RICHTEXT_FIELDS.TEXT]}`,
    ''
  );
};
