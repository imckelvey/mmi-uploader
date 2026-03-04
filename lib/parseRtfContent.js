/**
 * Parse plain text (from RTF) into structured sections for MMI HTML.
 */
function parseRtfContent(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim());
  const result = {
    summary: '',
    email: '',
    quoteText: '',
    quoteName: '',
    sectionTitle: 'U.S. Markets',
    introParagraphs: [],
    subsections: [], // { title, paragraphs }[]
    usRecapRows: [], // { name, month, ytd }[]
    yahooDate: '',
    whatInvestorsTitle: '',
    whatInvestorsBody: '',
    worldNarrative: '',
    worldRecapRows: [],
    worldYahooDate: '',
    indicatorsNote: '',
    indicatorsSections: [], // { title, body }[]
    fedParagraphs: [],
    byTheNumbersTitle: 'By the Numbers',
    byTheNumbersItems: [], // { value, label, ref }[]
    footnotes: [], // array of "1. Source, date" strings
    copyrightYear: new Date().getFullYear().toString(),
    monthLabel: 'November',
    monthAbbrev: 'nov',
    year: '2025',
    worldMonthLabel: 'November',
  };

  let i = 0;
  let seenCopyright = false;
  while (i < lines.length) {
    const line = lines[i];
    if (line.startsWith('Summary:')) {
      result.summary = line.replace(/^Summary:\s*/, '').trim();
      i++;
      continue;
    }
    if (line.startsWith('Email:')) {
      result.email = line.replace(/^Email:\s*/, '').trim();
      i++;
      continue;
    }
    if (line.startsWith('Quote:')) {
      let raw = line.replace(/^Quote:\s*/, '').trim();
      const quoteChars = /^[""\u201C\u201D]+|[""\u201C\u201D]+$/g;
      // If quote and attribution are on the same line (e.g. ...life changers." Tony Dungy, who won...), split there
      const sameLine = raw.match(/^(.+?[.!?])[""\u201C\u201D]+\s+([A-Z].+)$/);
      if (sameLine) {
        result.quoteText = sameLine[1].replace(quoteChars, '').trim();
        result.quoteName = sameLine[2].trim();
        i++;
        continue;
      }
      result.quoteText = raw.replace(quoteChars, '').trim();
      i++;
      if (i < lines.length && lines[i] && !lines[i].startsWith('U.S.') && !lines[i].startsWith('Markets')) {
        result.quoteName = lines[i].trim();
        i++;
      }
      continue;
    }
    if (line === 'U.S. and Canadian Markets' || line === 'U.S. Markets' || line === 'U.S. and Canadian markets') {
      result.sectionTitle = line;
      i++;
      const intro = [];
      const subsections = [];
      let currentSub = null;
      while (i < lines.length && lines[i] !== 'Markets Recap') {
        const l = lines[i];
        if (!l) { i++; continue; }
        // Subheading: short line, title case, no period
        if (l.length < 50 && !l.endsWith('.') && /^[A-Z]/.test(l) && !l.match(/^\d/) && l !== 'The Standard & Poor\'s 500 Index advanced 1.37 percent') {
          if (currentSub) subsections.push(currentSub);
          currentSub = { title: l, paragraphs: [] };
          i++;
          continue;
        }
        if (currentSub) {
          currentSub.paragraphs.push(l);
        } else {
          intro.push(l);
        }
        i++;
      }
      if (currentSub) subsections.push(currentSub);
      result.introParagraphs = intro.filter(p => p.length > 10);
      result.subsections = subsections;
      continue;
    }
    if (line === 'Markets Recap' && i < lines.length) {
      i++;
      while (i < lines.length && !lines[i].match(/^Yahoo Finance/)) {
        const l = lines[i];
        if (l && l.match(/^[A-Za-z&\/\s]+$/) && l.length > 2 && l.length < 50) {
          const name = l.trim();
          i++;
          const next1 = lines[i];
          const next2 = lines[i + 1];
          const num1 = parseNum(next1);
          const num2 = parseNum(next2);
          if (num1 !== null || num2 !== null) {
            result.usRecapRows.push({ name, month: next1 || '', ytd: next2 || '' });
            if (num1 !== null) i++;
            if (num2 !== null) i++;
          }
        }
        i++;
      }
      if (i < lines.length && lines[i].match(/^Yahoo Finance/)) {
        result.yahooDate = lines[i].replace(/^Yahoo Finance,\s*([^.]+)\..*/, '$1').trim();
        i++;
      }
      continue;
    }
    if (line.match(/^What Investors May Be Talking About in/)) {
      result.whatInvestorsTitle = line;
      i++;
      const body = [];
      while (i < lines.length && lines[i] !== 'World Markets' && !lines[i].match(/^World Market/)) {
        if (lines[i]) body.push(lines[i]);
        i++;
      }
      result.whatInvestorsBody = body.join('\n\n');
      continue;
    }
    if (line === 'World Markets') {
      i++;
      const narrative = [];
      while (i < lines.length && lines[i] !== 'Markets Recap' && !lines[i].match(/^Index$/)) {
        if (lines[i]) narrative.push(lines[i]);
        i++;
      }
      result.worldNarrative = narrative.join('\n\n');
      if (i < lines.length && (lines[i] === 'Markets Recap' || lines[i].match(/^Index$/))) {
        i++;
        while (i < lines.length && !lines[i].match(/^Yahoo Finance/)) {
          const l = lines[i];
          if (l && (l.match(/^\s*Hang Seng|KOSPI|Nikkei|Sensex|EGX|Bovespa|IPC|ASX|DAX|CAC|IBEX|FTSE|IT40/) || l.match(/^[A-Za-z0-9\s\-\.]+$/)) && l.length < 40) {
            const name = l.replace(/^\s+/, '');
            i++;
            const v1 = lines[i];
            const v2 = lines[i + 1];
            if (v1 !== undefined && v2 !== undefined) {
              result.worldRecapRows.push({ name, month: v1, ytd: v2 });
              i += 2;
            }
          }
          i++;
        }
        if (i < lines.length && lines[i].match(/^Yahoo Finance/)) {
          result.worldYahooDate = lines[i].replace(/^Yahoo Finance,\s*([^.]+)\..*/, '$1').trim();
        }
      }
      continue;
    }
    if (line === 'Indicators') {
      i++;
      if (lines[i] && lines[i].match(/Please note/)) {
        result.indicatorsNote = lines[i];
        i++;
      }
      const indSections = [];
      let current = null;
      while (i < lines.length && !lines[i].match(/^The Federal Reserve$/) && !lines[i].match(/^The Fed$/)) {
        const l = lines[i];
        if (l.match(/^[A-Z][a-z].*\([A-Z]+\)/) || (l.match(/^[A-Za-z\s]+:$/) && l.length < 60)) {
          if (current) indSections.push(current);
          current = { title: l.replace(/:$/, ''), body: '' };
        } else if (current && l) {
          current.body += (current.body ? '\n\n' : '') + l;
        }
        i++;
      }
      if (current) indSections.push(current);
      result.indicatorsSections = indSections;
      continue;
    }
    if (line === 'The Federal Reserve' || line === 'The Fed') {
      i++;
      const fed = [];
      while (i < lines.length && !lines[i].match(/^By the Numbers/)) {
        if (lines[i]) fed.push(lines[i]);
        i++;
      }
      result.fedParagraphs = fed;
      continue;
    }
    if (line.match(/^By the Numbers/)) {
      result.byTheNumbersTitle = line;
      i++;
      while (i < lines.length && !lines[i].match(/^---|^The content is developed/)) {
        const l = lines[i];
        if (l && (l.match(/^\$|^\d+%|^\d+\.\d|^\d{4}/) || l.match(/^[£€]/))) {
          const value = l;
          const ref = (value.match(/\d+$/) || [])[0];
          i++;
          const label = i < lines.length ? lines[i] : '';
          if (label) i++;
          result.byTheNumbersItems.push({ value, label, ref: ref || '' });
        } else {
          i++;
        }
      }
      continue;
    }
    if (line.match(/^Copyright\s+\d{4}/)) {
      seenCopyright = true;
      const m = line.match(/Copyright\s+(\d{4})/);
      if (m) result.copyrightYear = m[1];
      i++;
      continue;
    }
    // Skip "Sources" or "Sources:" section header so we don't consume it
    if (line.match(/^Sources\s*:?\s*$/i)) {
      i++;
      continue;
    }
    // Sources: all numbered lines (1. ..., 2. ..., etc.) after Copyright at end of RTF
    if (line.match(/^\d+\.\s+\S/) || line.match(/^\d+\.\s*$/)) {
      const footnotes = [];
      while (i < lines.length) {
        const ln = lines[i];
        if (ln.match(/^\d+\.\s+\S/)) {
          footnotes.push(ln);
          i++;
        } else if (ln.match(/^\d+\.\s*$/) && i + 1 < lines.length) {
          footnotes.push((ln.trim() + ' ' + (lines[i + 1] || '').trim()).trim());
          i += 2;
        } else if (!ln.trim()) {
          i++;
        } else {
          break;
        }
      }
      if (footnotes.length > 0 && seenCopyright) {
        result.footnotes = footnotes;
      }
      continue;
    }
    // Infer month/year from Email or first date in text
    if (result.email && result.email.match(/January|February|March|April|May|June|July|August|September|October|November|December/i)) {
      const months = { January: 'jan', February: 'feb', March: 'mar', April: 'apr', May: 'may', June: 'jun', July: 'jul', August: 'aug', September: 'sep', October: 'oct', November: 'nov', December: 'dec' };
      for (const [name, abbr] of Object.entries(months)) {
        if (result.email.includes(name)) {
          result.monthLabel = name;
          result.monthAbbrev = abbr;
          break;
        }
      }
    }
    const yearMatch = text.match(/\b(20\d{2})\b/);
    if (yearMatch) result.year = yearMatch[1];
    i++;
  }

  result.worldMonthLabel = result.monthLabel;
  return result;
}

function parseNum(s) {
  if (s == null) return null;
  const n = parseFloat(String(s).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? null : n;
}

module.exports = parseRtfContent;
