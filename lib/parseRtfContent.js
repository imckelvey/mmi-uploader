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
      const isRecapStart = (idx) => {
        const l = lines[idx];
        return l === 'S&P 500' && idx + 2 < lines.length && parseNum(lines[idx + 1]) !== null && parseNum(lines[idx + 2]) !== null;
      };
      while (i < lines.length && lines[i] !== 'Markets Recap' && !isRecapStart(i)) {
        const l = lines[i];
        if (!l) { i++; continue; }
        // Subheading: short line, title case, no period (skip S&P 500 / recap table start)
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
      // Parse US recap if we broke on "Markets Recap" or "S&P 500" block (September-style)
      if (i < lines.length && (lines[i] === 'Markets Recap' || isRecapStart(i))) {
        if (lines[i] === 'Markets Recap') i++;
        while (i < lines.length && !lines[i].match(/^Yahoo Finance/)) {
          const l = lines[i];
          if (l && (l.match(/^[A-Za-z&\/\s0-9\-]+$/) || l === 'S&P 500') && l.length > 2 && l.length < 55) {
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
              continue; // already advanced i, don't double-increment below
            }
          }
          i++;
        }
        if (i < lines.length && lines[i].match(/^Yahoo Finance/)) {
          result.yahooDate = lines[i].replace(/^Yahoo Finance,\s*([^.]+)\..*/, '$1').trim();
          i++;
        }
      }
      continue;
    }
    if (line === 'Markets Recap' && i < lines.length) {
      i++;
      while (i < lines.length && !lines[i].match(/^Yahoo Finance/)) {
        const l = lines[i];
        if (l && (l.match(/^[A-Za-z&\/\s0-9\-]+$/) || l === 'S&P 500') && l.length > 2 && l.length < 55) {
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
            continue;
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
      while (i < lines.length && lines[i] !== 'Markets Recap' && !lines[i].match(/^Index$/) && !lines[i].match(/^World Markets? Recap/i)) {
        if (lines[i]) narrative.push(lines[i]);
        i++;
      }
      result.worldNarrative = narrative.join('\n\n');
      if (i < lines.length && (lines[i] === 'Markets Recap' || lines[i].match(/^Index$/) || lines[i].match(/^World Markets? Recap/i))) {
        i++;
        const worldSectionStops = /^Indicators\s*:?\s*$|^The Fed\s*:?\s*$|^The Federal Reserve\s*:?\s*$|^By the Numbers|^Copyright\s+\d{4}|^---\s*$|^The content is developed|^Sources\s*:?\s*$/i;
        const isNumericLine = (s) => s != null && /^-?[\d.]+\s*%?\s*$/.test(String(s).trim());
        const isSectionHeader = (s) => s != null && (worldSectionStops.test(String(s).trim()) || /^(Emerging|Europe)(\s+Markets)?\s*$/i.test(String(s).trim()));
        const isIndexName = (s) => {
          if (!s || s.length > 45) return false;
          const t = String(s).trim();
          if (isNumericLine(t) || isSectionHeader(t)) return false;
          if (/^\d/.test(t)) return false;
          return /^[A-Za-z]/.test(t) && /[A-Za-z]/.test(t) && /^[A-Za-z0-9\s\-\.(),]+$/.test(t);
        };
        while (i < lines.length && !lines[i].match(/^Yahoo Finance/) && !worldSectionStops.test(lines[i] || '')) {
          const l = lines[i];
          if (!l) { i++; continue; }
          // One line per row: "Index Name  -2.76  3.90" or "Index Name\t-2.76\t3.90"
          const oneLineMatch = l.match(/^(.+?)\s+(-?[\d.]+\s*%?)\s+(-?[\d.]+\s*%?)\s*$/);
          if (oneLineMatch) {
            const name = oneLineMatch[1].trim();
            if (isIndexName(name)) {
              result.worldRecapRows.push({ name, month: oneLineMatch[2].trim(), ytd: oneLineMatch[3].trim() });
              i++;
              continue;
            }
          }
          const v1 = lines[i + 1];
          const v2 = lines[i + 2];
          if (v1 === undefined || v2 === undefined) { i++; continue; }
          const nameFirst = isIndexName(l) && isNumericLine(v1) && isNumericLine(v2);
          const monthYtdName = isNumericLine(l) && isNumericLine(v1) && isIndexName(v2);
          if (nameFirst) {
            result.worldRecapRows.push({ name: l.replace(/^\s+/, ''), month: v1, ytd: v2 });
            i += 3;
            continue;
          }
          if (monthYtdName) {
            result.worldRecapRows.push({ name: v2.replace(/^\s+/, ''), month: l, ytd: v1 });
            i += 3;
            continue;
          }
          i++;
        }
        if (i < lines.length && lines[i].match(/^Yahoo Finance/)) {
          result.worldYahooDate = lines[i].replace(/^Yahoo Finance,\s*([^.]+)\..*/, '$1').trim();
        }
      }
      continue;
    }
    if (line === 'Indicators' || /^Indicators\s*:?\s*$/.test(line)) {
      i++;
      if (lines[i] && lines[i].match(/Please note/)) {
        result.indicatorsNote = lines[i];
        i++;
      }
      const indSections = [];
      let current = null;
      while (i < lines.length && !lines[i].match(/^The Federal Reserve\s*:?\s*$/) && !lines[i].match(/^The Fed\s*:?\s*$/)) {
        const l = lines[i];
        // Section title: "Name (ABBR)", "Name:", "Title:", or title-case heading (e.g. Employment, Retail Sales, Industrial Production)
        const isSectionTitle = l.match(/^[A-Z][a-z].*\([A-Z]+\)/) ||
          (l.match(/^[A-Za-z\s]+:$/) && l.length < 60) ||
          (l.length > 2 && l.length < 60 && /^[A-Z][A-Za-z\s\-]+:\s*$/.test(l)) ||
          (l.length >= 2 && l.length < 55 && !l.endsWith('.') && /^[A-Z][a-z]*(\s+[A-Z][a-z]*)*$/.test(l.trim()));
        if (isSectionTitle) {
          if (current) indSections.push(current);
          current = { title: l.replace(/:$/, '').trim(), body: '' };
        } else if (current && l) {
          current.body += (current.body ? '\n\n' : '') + l;
        }
        i++;
      }
      if (current) indSections.push(current);
      result.indicatorsSections = indSections;
      continue;
    }
    if (line.match(/^The Federal Reserve\s*:?\s*$/) || line.match(/^The Fed\s*:?\s*$/)) {
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
