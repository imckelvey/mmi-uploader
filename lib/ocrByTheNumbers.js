const fs = require('fs');
const path = require('path');
const os = require('os');
const { createWorker } = require('tesseract.js');

/**
 * Run OCR on a By the Numbers screenshot buffer and parse into value/label items.
 * @param {Buffer} imageBuffer - PNG/JPEG buffer from RTFD
 * @returns {Promise<Array<{ value: string, label: string, ref: string }>>}
 */
async function ocrByTheNumbers(imageBuffer) {
  const ext = imageBuffer[0] === 0x89 && imageBuffer[1] === 0x50 ? '.png' : '.jpg';
  const tmpPath = path.join(os.tmpdir(), `mmi-btn-${Date.now()}${ext}`);
  try {
    fs.writeFileSync(tmpPath, imageBuffer);
    const worker = await createWorker('eng', 1, { logger: () => {} });
    try {
      const { data: { text } } = await worker.recognize(tmpPath);
      await worker.terminate();
      return parseOcrText(text || '');
    } finally {
      await worker.terminate().catch(() => {});
    }
  } finally {
    try { fs.unlinkSync(tmpPath); } catch (_) {}
  }
}

/**
 * Parse OCR output into value/label pairs. Values are typically: $X, 25%, 2024, "5.5 million", etc.
 */
function parseOcrText(text) {
  const items = [];
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const valueLike = /^[\$£€]|^\d+%|^\d{4}\b|^\d+\.?\d*\s*(million|billion|percent|%)?/i;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!valueLike.test(line) || line.length > 80) continue;
    const label = i + 1 < lines.length ? lines[i + 1] : '';
    const ref = (line.match(/\d+$/) || [])[0] || '';
    const value = line.replace(/\s*\d+$/, '').trim() || line;
    items.push({ value, label: label.trim(), ref });
    i++;
  }
  return items;
}

module.exports = { ocrByTheNumbers, parseOcrText };
