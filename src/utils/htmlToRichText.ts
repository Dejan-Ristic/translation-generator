const parseTextBreaks = (text: string) => text.replace(/<br\s*\/?>/gi, '\n');

export const htmlToRichText = (htmlText: string) => {
  return htmlText
    .split(/(<strong>.*?<\/strong>|<b>.*?<\/b>|<i>.*?<\/i>|<em>.*?<\/em>)/gi)
    .map((part) => {
      if (part.startsWith('<strong>'))
        return {
          font: { bold: true },
          text: parseTextBreaks(part.replace(/<\/?strong>/gi, '')),
        };
      if (part.startsWith('<b>'))
        return {
          font: { bold: true },
          text: parseTextBreaks(part.replace(/<\/?b>/gi, '')),
        };
      if (part.startsWith('<i>'))
        return {
          font: { italic: true },
          text: parseTextBreaks(part.replace(/<\/?i>/gi, '')),
        };
      if (part.startsWith('<em>'))
        return {
          font: { italic: true },
          text: parseTextBreaks(part.replace(/<\/?em>/gi, '')),
        };
      return { text: parseTextBreaks(part) };
    });
};
