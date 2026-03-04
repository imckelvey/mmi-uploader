const path = require('path');
const fs = require('fs');
const express = require('express');
const multer = require('multer');
const rtfToText = require('./lib/rtfToText');
const parseRtfContent = require('./lib/parseRtfContent');
const applyToHtml = require('./lib/applyToHtml');

const app = express();
const PORT = process.env.PORT || 3000;
const templatePath = path.join(__dirname, 'og code', 'dec_mmi_2025_fmg.html');
const outputDir = path.join(__dirname, 'generated reports');

const MONTHS = 'january|february|march|april|may|june|july|august|september|october|november|december'.split('|');
const ABBREV = 'jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec'.split('|');

function outputFilename(rtfName, custom) {
  if (custom && custom.trim()) return custom.trim().endsWith('.html') ? custom.trim() : custom.trim() + '.html';
  const base = path.basename(rtfName, '.rtf').replace(/\.rtf$/i, '');
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
  fileFilter: (_, file, cb) => file.originalname.toLowerCase().endsWith('.rtf') ? cb(null, true) : cb(new Error('Only .rtf files are allowed')),
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

app.post('/process', upload.single('rtf'), (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, error: 'No .rtf file uploaded' });

  const filename = outputFilename(req.file.originalname, req.body && req.body.filename);
  const outputPath = path.join(outputDir, filename);

  let template;
  try { template = fs.readFileSync(templatePath, 'utf8'); }
  catch (e) { return res.status(500).json({ success: false, error: 'Template not found: dec_mmi_2025_fmg.html' }); }

  try {
    const text = rtfToText(req.file.buffer);
    const data = parseRtfContent(text);
    const content = applyToHtml(template, data);
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(outputPath, content, 'utf8');
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
