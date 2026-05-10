export const checkIfRichTextContent = (contentArr: Array<any>) => {
  return contentArr.every((arrItem) => {
    if (!arrItem || typeof arrItem !== 'object' || Array.isArray(arrItem))
      return false;

    const objKeys = Object.keys(arrItem);
    if (
      objKeys.length === 0 ||
      objKeys.length > 2 ||
      objKeys.some((key) => !['text', 'font'].includes(key))
    )
      return false;

    if (!Object.hasOwn(arrItem, 'text') || typeof arrItem.text !== 'string')
      return false;

    if (
      Object.hasOwn(arrItem, 'font') &&
      (!arrItem.font ||
        typeof arrItem.font !== 'object' ||
        Array.isArray(arrItem.font))
    )
      return false;

    return true;
  });
};

export const parseAsText = (contentArr: Array<any>) => {
  return contentArr.reduce(
    (accumulator, currentItem) => `${accumulator}${currentItem.text}`,
    ''
  );
};
