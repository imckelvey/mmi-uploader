/**
 * Apply parsed RTF content to the template HTML. Returns updated HTML string.
 */
function escapeHtml(s) {
  if (!s) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Superscript span used for footnote refs (match template: vertical-align top, 10px)
const SUP_STYLE = "font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; vertical-align: top; font-size: 10px !important; margin: 0; padding: 0;";

function wrapParagraph(text) {
  if (!text) return '';
  let body = String(text).trim();
  let ref = '';
  const m = body.match(/(\.?)\s*(\d+(,\d+)*)\s*$/);
  if (m) {
    ref = m[2];
    body = body.replace(/(\.?)\s*(\d+(,\d+)*)\s*$/, m[1] ? '.' : '').trim();
  }
  const safeBody = escapeHtml(body);
  const refSpan = ref ? `<span style="${SUP_STYLE}">${escapeHtml(ref)}</span>` : '';
  return `<p style="color: #666; font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; font-weight: normal; font-size: 20px !important; line-height: 1.8; margin: 0 0 15px; padding: 0;">${safeBody}${refSpan}</p>\n`;
}

function applyToHtml(html, parsed) {
  let out = html;

  // Section title: replace only the text inside the first <h2> that contains "U.S. Markets", and always close with </h2>
  out = out.replace(
    /(<h2[^>]*>)U\.S\. Markets<\/?h2>/,
    `$1${escapeHtml(parsed.sectionTitle)}</h2>`
  );

  // Quote block: quote text (no quotation marks), then attribution in blue + italic to match reference
  const quoteText = escapeHtml(parsed.quoteText);
  const quoteName = escapeHtml(parsed.quoteName);
  out = out.replace(
    /<p class="mmi-quote-text"[^>]*>[\s\S]*?<\/p>\s*<p class="mmi-quote-name"[^>]*>[\s\S]*?<\/p>/,
    `<p class="mmi-quote-text" style="font-weight: 400; color: #777; line-height: 1.8; text-align: left; margin: 0; font-size: 29px;">${quoteText}</p>\n<p class="mmi-quote-name" style="font-weight: 500; color: #2c7cb5; padding-bottom: 10px; height: 100%; font-size: 18px;"><i>${quoteName}</i></p>`
  );

  // Intro paragraphs: replace first two body paragraphs after section (before quote)
  const introBlock = (parsed.introParagraphs.slice(0, 2).map(wrapParagraph)).join('');
  out = out.replace(
    /(<td style="font-family: 'Rubik'[^>]*>\s*)(<p style="color: #666[^>]*>Stocks were mixed[\s\S]*?<\/p>\s*<p style="color: #666[^>]*>The Standard[\s\S]*?<\/p>)/,
    `$1${introBlock}`
  );

  // U.S. Market Recap header (November 2025 -> Month Year)
  const recapTitle = `U.S. Market Recap for ${parsed.monthLabel} ${parsed.year}`;
  out = out.replace(/U\.S\. Market Recap for November 2025/g, recapTitle);

  // Yahoo Finance date in US table footer
  out = out.replace(/Yahoo Finance, November,\s*30,\s*2025\./g, `Yahoo Finance, ${parsed.yahooDate || parsed.monthLabel + ' ' + parsed.year + '.'}.`);

  // What Investors section title
  out = out.replace(/What Investors May Be Talking About in December/g, escapeHtml(parsed.whatInvestorsTitle || 'What Investors May Be Talking About'));

  // World Market Recap header
  out = out.replace(/World Market Recap for November 2025/g, `World Market Recap for ${parsed.worldMonthLabel} ${parsed.year}`);

  // World table month column header (November -> Month) — use RegExp so pattern doesn't consume </ from </th>
  out = out.replace(new RegExp('>November \\(%\\)', 'g'), `>${parsed.worldMonthLabel} (%)`);

  // Yahoo world footer
  out = out.replace(/Yahoo Finance, November 30, 2025\./g, `Yahoo Finance, ${parsed.worldYahooDate || parsed.yahooDate || parsed.monthLabel + ' ' + parsed.year + '.'}.`);

  // CSS for mobile table label
  out = out.replace(/content: "NOVEMBER"/g, `content: "${(parsed.worldMonthLabel || 'NOVEMBER').substring(0, 3).toUpperCase()}"`);

  // By the Numbers title
  out = out.replace(/By the Numbers: Gift Wrapping/g, escapeHtml(parsed.byTheNumbersTitle || 'By the Numbers'));

  // Copyright
  out = out.replace(/Copyright 2025 FMG Suite\./g, `Copyright ${parsed.copyrightYear} FMG Suite.`);

  // US Recap table values: map by row name (S&P 500, Nasdaq, Russell 2000, 10-Year Treasury)
  const rowMap = {};
  (parsed.usRecapRows || []).forEach(r => {
    const key = (r.name || '').toLowerCase().replace(/\s*\/\s*tsx.*/i, '').trim();
    if (key.includes('s&p') || key === 's&amp;p 500') rowMap['sp500'] = r;
    else if (key.includes('nasdaq')) rowMap['nasdaq'] = r;
    else if (key.includes('russell')) rowMap['russell'] = r;
    else if (key.includes('10-year') || key.includes('treasury')) rowMap['treasury'] = r;
  });

  const arrowUp = '<img src="http://fmg-websites-custom.s3.amazonaws.com/Monthly_Market_Insights/January_2018/arrow_up.png" height="20" width="17" style="margin-right: 10px; padding: 0px !important;" alt="green up arrow" />';
  const arrowDown = '<img src="http://fmg-websites-custom.s3.amazonaws.com/Monthly_Market_Insights/January_2018/arrow_down.png" height="20" width="17" style="margin-right: 10px; padding: 0px !important;" alt="red down arrow" />';

  // US recap table: each column has <tr><td class="market_name">Name</td></tr><tr><td><img ... />NUMBER</td></tr>...
  const usOrder = ['sp500', 'nasdaq', 'russell', 'treasury'];
  const defaults = [{ month: '0.13', ytd: '16.45' }, { month: '-1.51', ytd: '21.00' }, { month: '0.85', ytd: '12.12' }, { month: '4.02', ytd: '-0.56' }];
  const replacements = usOrder.map((key, idx) => {
    const r = rowMap[key] || defaults[idx];
    return {
      month: String((r && r.month) || defaults[idx].month).trim(),
      ytd: String((r && r.ytd) || defaults[idx].ytd).trim(),
    };
  });
  const columns = [
    { name: 'S&#38;P 500', defMonth: '0.13', defYtd: '16.45' },
    { name: 'Nasdaq', defMonth: '-1.51', defYtd: '21.00' },
    { name: 'Russell 2000', defMonth: '0.85', defYtd: '12.12' },
    { name: '10-Year Treasury', defMonth: '4.02', defYtd: '-0.56' },
  ];
  columns.forEach((col, idx) => {
    const r = replacements[idx];
    const monthArrow = r.month.startsWith('-') ? arrowDown : arrowUp;
    const esc = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    out = out.replace(
      new RegExp(`(<td class="market_name"[^>]*>${col.name}<\\/td>\\s*<\\/tr>\\s*<tr>\\s*<td[^>]*>)(<img[^>]+\\/>)${esc(col.defMonth)}(<\\/td>)`),
      `$1${monthArrow}${r.month}$3`
    );
  });
  // YTD cells: replace existing arrow+number with our arrow+value (capture only opening <td>, not the img)
  out = out.replace(
    /(<td>&#8204;<\/td>\s*<\/tr>\s*<tr>\s*<td[^>]*>)<img[^>]+\/>16\.45(<\/td>)/,
    `$1${replacements[0].ytd.startsWith('-') ? arrowDown : arrowUp}${replacements[0].ytd}$2`
  );
  out = out.replace(
    /(<td>&#8204;<\/td>\s*<\/tr>\s*<tr>\s*<td[^>]*>)<img[^>]+\/>21\.00(<\/td>)/,
    `$1${replacements[1].ytd.startsWith('-') ? arrowDown : arrowUp}${replacements[1].ytd}$2`
  );
  out = out.replace(
    /(<td>&#8204;<\/td>\s*<\/tr>\s*<tr>\s*<td[^>]*>)<img[^>]+\/>12\.12(<\/td>)/,
    `$1${replacements[2].ytd.startsWith('-') ? arrowDown : arrowUp}${replacements[2].ytd}$2`
  );
  out = out.replace(
    /(<td>&#8204;<\/td>\s*<\/tr>\s*<tr>\s*<td[^>]*>)<img[^>]+\/>-0\.56(<\/td>)/,
    `$1${replacements[3].ytd.startsWith('-') ? arrowDown : arrowUp}${replacements[3].ytd}$2`
  );

  // Footnotes/sources: replace the entire block from first <p>1. ...</p> through last footnote (before </div><style>)
  if (parsed.footnotes && parsed.footnotes.length > 0) {
    const footnoteHtml = parsed.footnotes.map(f => `<p style="color: #2f4447 !important; font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; font-weight: normal; font-size: 12px !important; line-height: 1.8; margin: 0; padding: 0;">${escapeHtml(f)}</p>\n`).join('');
    out = out.replace(/<p style="color: #2f4447[^>]*>\s*1\.\s*\S[\s\S]*?(?=\s*<\/div>\s*<style>)/,
      footnoteHtml);
  }

  return out;
}

module.exports = applyToHtml;
