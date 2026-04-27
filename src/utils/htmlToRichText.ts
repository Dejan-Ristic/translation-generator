const parseTextBreaks = (text: string) => text.replace(/<br\s*\/?>/gi, '\n');

const checkBoldText = (text: string) => {
  if (text.startsWith('<strong>'))
    return {
      font: { bold: true },
      text: text.replace(/<\/?strong>/gi, ''),
    };
  if (text.startsWith('<b>'))
    return {
      font: { bold: true },
      text: text.replace(/<\/?b>/gi, ''),
    };
};

const checkItalicText = (text: string) => {
  if (text.startsWith('<i>'))
    return {
      font: { italic: true },
      text: text.replace(/<\/?i>/gi, ''),
    };
  if (text.startsWith('<em>'))
    return {
      font: { italic: true },
      text: text.replace(/<\/?em>/gi, ''),
    };
};

const parseHtmlElements = (text: string): { text: string } => {
  let formattedText: { text: string } | undefined = checkBoldText(text);
  if (!formattedText) formattedText = checkItalicText(text);
  return formattedText ?? { text };
};

export const htmlToRichText = (htmlText: string) => {
  const partitionedText = htmlText.split(
    /(<strong>.*?<\/strong>|<b>.*?<\/b>|<i>.*?<\/i>|<em>.*?<\/em>)/gi
  );
  if (partitionedText.length > 1)
    return partitionedText.map((part) => {
      const richText = parseHtmlElements(part);
      return { ...richText, text: parseTextBreaks(richText.text) };
    });
  else return [{ text: parseTextBreaks(partitionedText[0]) }];
};
