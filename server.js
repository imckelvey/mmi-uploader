const path = require('path');
const fs = require('fs');
const express = require('express');
const multer = require('multer');
const getRtfFromUpload = require('./lib/getRtfFromUpload');
const rtfToText = require('./lib/rtfToText');
const parseRtfContent = require('./lib/parseRtfContent');
const applyToHtml = require('./lib/applyToHtml');
const { ocrByTheNumbers } = require('./lib/ocrByTheNumbers');

const app = express();
const PORT = process.env.PORT || 3000;
const templatePath = path.join(__dirname, 'og code', 'dec_mmi_2025_fmg.html');
const outputDir = path.join(__dirname, 'generated reports');

const MONTHS = 'january|february|march|april|may|june|july|august|september|october|november|december'.split('|');
const ABBREV = 'jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec'.split('|');

function outputFilename(rtfName, custom) {
  if (custom && custom.trim()) return custom.trim().endsWith('.html') ? custom.trim() : custom.trim() + '.html';
  const base = path.basename(rtfName).replace(/\.(rtf|rtfd|zip)$/i, '');
  const m = base.match(/(\w+)\s+(\d{4})\s*MMI/i) || base.match(/(\w+)\s+(\d{4})/i);
  if (m) {
    const i = MONTHS.indexOf((m[1] || '').toLowerCase());
    return `${i >= 0 ? ABBREV[i] : (m[1] || '').slice(0, 3)}_mmi_${m[2] || ''}_fmg.html`;
  }
  return (base.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_-]/g, '').toLowerCase() || 'mmi') + '_fmg.html';
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (_, file, cb) => {
    const name = (file.originalname || '').toLowerCase();
    const ok = name.endsWith('.rtf') || name.endsWith('.rtfd') || name.endsWith('.zip');
    cb(ok ? null : new Error('Only .rtf, .rtfd, or .zip files are allowed'), ok);
  },
});

app.use(express.json());
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});
app.use(express.static(path.join(__dirname, 'public')));

// So you can confirm deployed app matches local (same code, template, Node)
app.get('/version', (req, res) => {
  const pkg = require('./package.json');
  const templateExists = fs.existsSync(templatePath);
  res.json({
    version: pkg.version,
    node: process.version,
    template: templateExists ? 'ok' : 'missing',
    templatePath: templatePath,
  });
});

app.post('/process', upload.single('rtf'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });

  const filename = outputFilename(req.file.originalname, req.body && req.body.filename);
  const outputPath = path.join(outputDir, filename);

  let template;
  try { template = fs.readFileSync(templatePath, 'utf8'); }
  catch (e) { return res.status(500).json({ success: false, error: 'Template not found: dec_mmi_2025_fmg.html' }); }

  let uploadResult;
  try { uploadResult = getRtfFromUpload(req.file.buffer); }
  catch (e) { return res.status(400).json({ success: false, error: e.message || 'Could not extract RTF from upload' }); }

  const { rtfBuffer, byTheNumbersImage } = uploadResult;
  const isZip = req.file.buffer[0] === 0x50 && req.file.buffer[1] === 0x4b;
  console.log('[MMI] Upload: size=' + req.file.buffer.length + ' isZip=' + isZip + ' rtfLen=' + rtfBuffer.length);

  try {
    const text = rtfToText(rtfBuffer);
    console.log('[MMI] RTF->text: len=' + (text ? text.length : 0) + ' preview=' + (text ? text.slice(0, 120).replace(/\n/g, ' ') : '') + '...');
    const data = parseRtfContent(text);
    console.log('[MMI] Parse: sectionTitle=' + (data.sectionTitle || '') + ' introParas=' + (data.introParagraphs && data.introParagraphs.length) + ' subsections=' + (data.subsections && data.subsections.length));
    if (byTheNumbersImage && byTheNumbersImage.buffer) {
      try {
        const ocrItems = await ocrByTheNumbers(byTheNumbersImage.buffer);
        if (ocrItems.length > 0) data.byTheNumbersItems = ocrItems;
      } catch (ocrErr) {
        console.warn('By the Numbers OCR failed:', ocrErr.message);
      }
    }
    const content = applyToHtml(template, data);
    try {
      fs.mkdirSync(outputDir, { recursive: true });
      fs.writeFileSync(outputPath, content, 'utf8');
    } catch (writeErr) {
      console.warn('Could not save report to disk (e.g. read-only filesystem on host):', writeErr.message);
    }
    res.json({ success: true, filename, content });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message || 'Processing failed' });
  }
});

app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    return res.status(400).json({ success: false, error: err.code === 'LIMIT_FILE_SIZE' ? 'File too large' : err.message });
  }
  if (err) {
    return res.status(400).json({ success: false, error: err.message || 'Upload failed' });
  }
  next();
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`MMI uploader running at http://localhost:${PORT}`);
  console.log(`Open that URL in your browser to use the uploader.`);
});
