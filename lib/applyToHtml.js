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

  // Quote block: quote text styled like "You can always improve." (darker gray, semi-bold); attribution like "Shai Gilgeous-Alexander..." (blue, regular)
  const quoteText = escapeHtml(parsed.quoteText);
  const quoteNameHtml = escapeHtml(parsed.quoteName || '').replace(/\n/g, '<br />');
  out = out.replace(
    /<p class="mmi-quote-text"[^>]*>[\s\S]*?<\/p>\s*<p class="mmi-quote-name"[^>]*>[\s\S]*?<\/p>/,
    `<p class="mmi-quote-text" style="font-weight: 600; color: #4A4A4A; line-height: 1.8; text-align: left; margin: 0; font-size: 29px;">${quoteText}</p>\n<p class="mmi-quote-name" style="font-weight: 400; color: #2c7cb5; padding-bottom: 10px; height: 100%; font-size: 18px;">${quoteNameHtml}</p>`
  );

  // Intro paragraphs: replace first two body paragraphs after section (before quote)
  const introBlock = (parsed.introParagraphs.slice(0, 2).map(wrapParagraph)).join('');
  // Match first two <p style="color: #666...">...</p> in the U.S. Markets body td (any content so Sept/Dec/etc. all work)
  out = out.replace(
    /(<td style="font-family: 'Rubik'[^>]*>\s*)(<p style="color: #666[^>]*>[\s\S]*?<\/p>\s*<p style="color: #666[^>]*>[\s\S]*?<\/p>)(\s*<!-- START of Quote)/,
    `$1${introBlock}$3`
  );

  // U.S. subsections (h3 + paragraphs between quote and U.S. Market Recap table)
  const h3Style = "font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; line-height: 1.8; color: #35A2B1 !important; font-weight: 500; font-size: 26px !important; margin: 24px 0 5px 0; padding: 0;";
  const subsectionsBlock = (parsed.subsections || [])
    .map(sub => {
      const title = escapeHtml(sub.title || '');
      const paras = (sub.paragraphs || []).filter(p => p.length > 10).map(wrapParagraph).join('');
      return `<h3 style="${h3Style}">${title}</h3>\n${paras}`;
    })
    .join('\n');
  if (subsectionsBlock) {
    out = out.replace(
      /(<!-- END of Quote -->\s*)\n<h3 style="[^"]*"[^>]*>[\s\S]*?<table class="mmi_table-us"/,
      `$1\n${subsectionsBlock}\n\n<table class="mmi_table-us"`
    );
  }

  // U.S. Market Recap header (November 2025 -> Month Year)
  const recapTitle = `U.S. Market Recap for ${parsed.monthLabel} ${parsed.year}`;
  out = out.replace(/U\.S\. Market Recap for November 2025/g, recapTitle);

  // Yahoo Finance date in US table footer
  out = out.replace(/Yahoo Finance, November,\s*30,\s*2025\./g, `Yahoo Finance, ${parsed.yahooDate || parsed.monthLabel + ' ' + parsed.year + '.'}.`);

  // What Investors section title (replace any month so template works for any report)
  out = out.replace(/What Investors May Be Talking About in [^<]+/g, escapeHtml(parsed.whatInvestorsTitle || 'What Investors May Be Talking About'));

  // What Investors body: replace the entire cell content between What Investors h2 and World Markets h2
  if (parsed.whatInvestorsBody) {
    const whatInvestorsParas = parsed.whatInvestorsBody.split(/\n\n+/).filter(p => p.trim().length > 10).map(wrapParagraph).join('');
    if (whatInvestorsParas) {
      // Replace whatever is inside the body <td> (no dependency on <p> structure) up to World Markets h2
      out = out.replace(
        /(What Investors May Be Talking About in[^<]*<\/h2>\s*<\/td>\s*<\/tr>\s*<tr[^>]*>\s*<td[^>]*>\s*)([\s\S]*?)(\s*<\/td>\s*<\/tr>\s*<tr[^>]*>\s*<td[^>]*>\s*<h2[^>]*>World Markets<\/h2>)/,
        (m, before, cellContent, after) => before + whatInvestorsParas + after
      );
    }
  }

  // World Markets narrative: replace entire cell content between World Markets h2 and World Market Recap table
  if (parsed.worldNarrative) {
    const worldParas = parsed.worldNarrative.split(/\n\n+/).filter(p => p.trim().length > 10).map(wrapParagraph).join('');
    if (worldParas) {
      out = out.replace(
        /(<h2[^>]*>World Markets<\/h2>\s*<\/td>\s*<\/tr>\s*<tr[^>]*>\s*<td[^>]*>\s*)([\s\S]*?)(<table class="mmi_table" cellpadding="0" cellspacing="0")/,
        (m, before, cellContent, after) => before + worldParas + '\n\n' + after
      );
    }
  }

  // World Market Recap header
  out = out.replace(/World Market Recap for November 2025/g, `World Market Recap for ${parsed.worldMonthLabel} ${parsed.year}`);

  // World table month column header (November -> Month) — use RegExp so pattern doesn't consume </ from </th>
  out = out.replace(new RegExp('>November \\(%\\)', 'g'), `>${parsed.worldMonthLabel} (%)`);

  // Yahoo world footer
  out = out.replace(/Yahoo Finance, November 30, 2025\./g, `Yahoo Finance, ${parsed.worldYahooDate || parsed.yahooDate || parsed.monthLabel + ' ' + parsed.year + '.'}.`);

  // CSS for mobile table label
  out = out.replace(/content: "NOVEMBER"/g, `content: "${(parsed.worldMonthLabel || 'NOVEMBER').substring(0, 3).toUpperCase()}"`);

  // World Market Recap table body: split rows into Emerging Markets and Europe, insert Europe subheader
  const arrowUp = '<img src="http://fmg-websites-custom.s3.amazonaws.com/Monthly_Market_Insights/January_2018/arrow_up.png" height="20" width="17" style="margin-right: 10px; padding: 0px !important;" alt="green up arrow" />';
  const arrowDown = '<img src="http://fmg-websites-custom.s3.amazonaws.com/Monthly_Market_Insights/January_2018/arrow_down.png" height="20" width="17" style="margin-right: 10px; padding: 0px !important;" alt="red down arrow" />';
  const EUROPE_INDICES = /^(DAX|CAC\s*40|IBEX\s*35|FTSE\s*100|IT\s*40|IT40)/i;
  const isEuropeRow = (name) => EUROPE_INDICES.test(String(name || '').trim());
  if (parsed.worldRecapRows && parsed.worldRecapRows.length > 0) {
    const worldRowBgs = ['#fafafa', '#fff'];
    const emerging = parsed.worldRecapRows.filter((r) => !isEuropeRow(r.name));
    const europe = parsed.worldRecapRows.filter((r) => isEuropeRow(r.name));
    const makeRow = (row, idx, startIdx) => {
      const bg = worldRowBgs[(startIdx + idx) % 2];
      const monthVal = String(row.month || '').trim();
      const ytdVal = String(row.ytd || '').trim();
      const monthArrow = monthVal.startsWith('-') ? arrowDown : arrowUp;
      const ytdArrow = ytdVal.startsWith('-') ? arrowDown : arrowUp;
      return `<tbody style="width: 100%; font-weight: 400; font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif';">
<tr>
<td class="mmi-category pushedin" style="background-color: ${bg}; padding: 20px;">${escapeHtml(row.name || '')}</td>
<td class="info " id="month" style="text-align: center; background-color: ${bg};">${monthArrow}${escapeHtml(monthVal)}</td>
<td class="info " id="year" style="text-align: center; background-color: ${bg};">${ytdArrow}${escapeHtml(ytdVal)}</td>
</tr>
</tbody>`;
    };
    const emergingBlock = emerging.map((row, idx) => makeRow(row, idx, 0)).join('\n');
    const europeHeaderRow = `<tbody class="mmi-body-header " width="100%" style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif';">
<tr style="background-color: #d2d2d2; color: #4e4e4e;">
<th class="mmi-top-sub" style="font-weight: 500; color: #4e4e4e; font-size: 20px; padding: 20px; text-transform: uppercase; text-align: left; padding-left: 6% !important;" width="33.333%" align="center" valign="center">Europe</th>
<th class="mmi-body-empty europe_header"><b>&#160;</b></th>
<th class="mmi-body-empty europe_header"><b>&#160;</b></th>
</tr>
</tbody>`;
    const europeBlock = europe.map((row, idx) => makeRow(row, idx, emerging.length)).join('\n');
    const worldDataBlock = [emergingBlock, europe.length ? europeHeaderRow : '', europeBlock].filter(Boolean).join('\n');
    out = out.replace(
      /(<\/tbody>\s*)((?:<tbody[^>]*>[\s\S]*?<\/tbody>\s*)+)(<tfoot class="mmi-top-header")/,
      (m, before, oldData, after) => {
        const preceding = out.substring(Math.max(0, out.indexOf(m) - 600), out.indexOf(m));
        if (preceding.includes('Emerging Markets') && preceding.includes('Year-to-Date (%)')) return before + worldDataBlock + '\n' + after;
        return m;
      }
    );
  }

  // Indicators: replace entire block (note + sections) between Indicators h2 and The Fed h2.
  // Regex allows any content in first cell (e.g. <p>Please note...</p>) between </h2> and the row that contains The Fed.
  const noteHtml = parsed.indicatorsNote
    ? `<p style="color: #666; font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; font-weight: normal; font-size: 20px !important; line-height: 1.8; margin: 0 0 24px; padding: 0; font-style: italic;">${escapeHtml(parsed.indicatorsNote)}</p>\n</td></tr><tr style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; margin: 0; padding: 0;"><td style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; margin: 0; padding: 0;">\n`
    : `</td></tr><tr style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; margin: 0; padding: 0;"><td style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; margin: 0; padding: 0;">\n`;
  let sectionsHtml = '';
  (parsed.indicatorsSections || []).forEach(sec => {
    sectionsHtml += `<h3 style="${h3Style}">${escapeHtml(sec.title || '')}</h3>\n`;
    (sec.body || '').split(/\n\n+/).filter(p => p.trim()).forEach(p => { sectionsHtml += wrapParagraph(p); });
  });
  const indicatorsBlockContent = noteHtml + sectionsHtml;
  out = out.replace(
    /(<h2[^>]*>Indicators<\/h2>\s*)([\s\S]*?)(<\/td>\s*<\/tr>\s*<tr[^>]*>\s*<td[^>]*>\s*<h2[^>]*>The Fed<\/h2>)/,
    (m, before, cellContent, after) => before + indicatorsBlockContent + after
  );

  // The Fed: replace entire cell content between The Fed h2 and By the Numbers (always; use empty when no RTF content)
  const fedBlock = (parsed.fedParagraphs && parsed.fedParagraphs.length > 0)
    ? parsed.fedParagraphs.filter(p => p.trim().length > 10).map(wrapParagraph).join('')
    : '';
  out = out.replace(
    /(<h2[^>]*>The Fed<\/h2>\s*<\/td>\s*<\/tr>\s*<tr[^>]*>\s*<td[^>]*>\s*)([\s\S]*?)(<h2 class="by-the-nums-title")/,
    (m, before, cellContent, after) => before + fedBlock + '\n\n' + after
  );

  // By the Numbers title
  out = out.replace(/By the Numbers: Gift Wrapping/g, escapeHtml(parsed.byTheNumbersTitle || 'By the Numbers'));

  // By the Numbers: build cards from parsed items (RTF or OCR from screenshot) using template structure
  const btnColors = [
    { bg: '#93d6df', fg: '#555' },
    { bg: '#6b71b2', fg: '#fff' },
    { bg: '#2D7CB6', fg: '#fff' },
    { bg: '#fff', fg: '#555' },
    { bg: '#93d6df', fg: '#555' },
    { bg: '#6b71b2', fg: '#fff' },
  ];
  const btnSupStyle = "font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif'; vertical-align: top; font-size: 10px; margin: 0; padding: 0;";
  let byTheNumbersCardsHtml = '';
  if (parsed.byTheNumbersItems && parsed.byTheNumbersItems.length > 0) {
    const items = parsed.byTheNumbersItems;
    const card = (item, c) => {
      const refSpan = (item.ref && String(item.ref).trim()) ? `<span style="${btnSupStyle}">${escapeHtml(item.ref)}</span>` : '';
      return `<div class="col-md-6" style="background-color: ${c.bg}; color: ${c.fg} !important; width: 50%; height: auto; text-align: center; display: inline-block !important; padding: 0 8%; margin: 0;">
<h3 style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif'; line-height: 1.1; font-size: 32px; margin: 0; font-weight: 400; padding: 100px 0 15px 0; text-align: center; color: ${c.fg} !important;">${escapeHtml(item.value || '')}${refSpan}</h3>
<p style="margin: 0; padding: 0 0 100px 0; color: ${c.fg} !important; font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important;">${escapeHtml(item.label || '')}</p>
</div>`;
    };
    const fullCard = (item, c) => {
      const refSpan = (item.ref && String(item.ref).trim()) ? `<span style="${btnSupStyle}">${escapeHtml(item.ref)}</span>` : '';
      return `<div class="col-md-12" style="background-color: ${c.bg}; color: ${c.fg} !important; width: 100%; height: auto; text-align: center; display: inline-block !important; padding: 0 8%; margin: 0;">
<h3 style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif'; line-height: 1.1; font-size: 32px; margin: 0; font-weight: 400; padding: 100px 0 15px 0; text-align: center; color: ${c.fg} !important;">${escapeHtml(item.value || '')}${refSpan}</h3>
<p style="margin: 0; padding: 0 0 100px 0; color: ${c.fg} !important; font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important;">${escapeHtml(item.label || '')}</p>
</div>`;
    };
    let idx = 0;
    const rows = [];
    while (idx < items.length) {
      const c1 = btnColors[idx % btnColors.length];
      const c2 = idx + 1 < items.length ? btnColors[(idx + 1) % btnColors.length] : null;
      if (rows.length % 2 === 1 && c2) {
        rows.push(`<div style="width: 100%; overflow: hidden; display: flex;"><!-- ===== Two Column ===== -->\n${card(items[idx], c1)}\n${card(items[idx + 1], c2)}\n</div>`);
        idx += 2;
      } else {
        rows.push(`<div style="width: 100%; overflow: hidden; display: flex;"><!-- ===== One Column ===== -->\n${fullCard(items[idx], c1)}\n</div>`);
        idx += 1;
      }
    }
    byTheNumbersCardsHtml = `<div class="" style="font-family: 'Rubik', 'Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif' !important; max-width: 800px; display: block; margin: 0 auto;">\n${rows.join('\n')}\n</div>`;
  }
  out = out.replace(
    /(<div class="by-the-numbers-container">)\s*([\s\S]*?)(<hr\s*\/>\s*<!-- ===== End BTN ===== -->)/,
    (m, open, inner, end) => open + '\n' + (byTheNumbersCardsHtml || '') + '\n' + end
  );

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
