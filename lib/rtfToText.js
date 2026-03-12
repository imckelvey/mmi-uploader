const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

/**
 * Convert RTF buffer to plain text.
 * On macOS, uses textutil for best results. Otherwise uses a basic RTF strip.
 */
function rtfToText(rtfBuffer) {
  if (os.platform() === 'darwin') {
    try {
      const tmpRtf = path.join(os.tmpdir(), `mmi-upload-${Date.now()}.rtf`);
      const tmpTxt = path.join(os.tmpdir(), `mmi-upload-${Date.now()}.txt`);
      fs.writeFileSync(tmpRtf, rtfBuffer);
      execSync(`textutil -convert txt -output "${tmpTxt}" "${tmpRtf}"`, { stdio: 'pipe' });
      const text = fs.readFileSync(tmpTxt, 'utf8');
      try { fs.unlinkSync(tmpRtf); } catch (_) {}
      try { fs.unlinkSync(tmpTxt); } catch (_) {}
      return text;
    } catch (e) {
      // Fall back to strip if textutil fails
    }
  }
  return stripRtf(rtfBuffer.toString('binary'));
}

/**
 * Basic RTF to text: remove control words and groups, keep text.
 * Handles \'XX hex, \uNNNN? Unicode, \par/\line, and plain ASCII.
 * Used on Linux/Render where textutil is not available.
 */
function stripRtf(binary) {
  let text = '';
  let i = 0;
  const len = binary.length;
  let skipGroup = 0;

  while (i < len) {
    const ch = binary[i];
    if (ch === '\\') {
      i++;
      if (i >= len) break;
      const next = binary[i];
      if (next === '{' || next === '}' || next === '\\') {
        text += next;
        i++;
      } else if (next === "'" && i + 2 <= len) {
        const hex = binary.slice(i + 1, i + 3);
        const code = parseInt(hex, 16);
        if (!isNaN(code)) text += String.fromCharCode(code);
        i += 3;
      } else {
        let word = '';
        while (i < len && /[a-zA-Z]/.test(binary[i])) {
          word += binary[i];
          i++;
        }
        let num = '';
        if (i < len && (binary[i] === '-' || /[0-9]/.test(binary[i]))) {
          if (binary[i] === '-') { num += '-'; i++; }
          while (i < len && /[0-9]/.test(binary[i])) { num += binary[i]; i++; }
        }
        if (i < len && binary[i] === ' ') i++;
        // Unicode: \uNNNN? (optional fallback char after ?)
        if (word === 'u' && num !== '') {
          const code = parseInt(num, 10);
          if (!isNaN(code)) {
            const unicode = code < 0 ? code + 65536 : code;
            if (unicode <= 0xFFFF) text += String.fromCharCode(unicode);
            else text += String.fromCodePoint(unicode);
          }
          if (i < len && binary[i] === '?') i++;
        } else if (word === 'par' || word === 'line') {
          text += '\n';
        }
        // skip other control words
      }
      continue;
    }
    if (ch === '{') {
      skipGroup++;
      i++;
      continue;
    }
    if (ch === '}') {
      if (skipGroup > 0) skipGroup--;
      i++;
      continue;
    }
    if (skipGroup === 0 && (ch >= ' ' && ch <= '~' || ch === '\n' || ch === '\r' || ch === '\t')) {
      if (ch !== '\r') text += ch;
    }
    i++;
  }
  return text.replace(/\n{3,}/g, '\n\n').trim();
}

module.exports = rtfToText;
