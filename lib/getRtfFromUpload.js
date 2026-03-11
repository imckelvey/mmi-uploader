const AdmZip = require('adm-zip');

const IMAGE_EXT = /\.(png|jpg|jpeg|gif|webp)$/i;

/**
 * Return RTF buffer and optional By the Numbers screenshot from an uploaded file.
 * - If the buffer is plain RTF, return { rtfBuffer, byTheNumbersImage: null }.
 * - If the buffer is a zip (RTFD package or .zip), find TXT.rtf and the first image (for By the Numbers screenshot).
 * @param {Buffer} buffer - Uploaded file content
 * @returns {{ rtfBuffer: Buffer, byTheNumbersImage: { buffer: Buffer, mimeType: string } | null }}
 * @throws {Error} - If zip but no .rtf entry found, or zip is invalid
 */
function getRtfFromUpload(buffer) {
  if (!buffer || buffer.length < 4) throw new Error('Empty or invalid file');
  const isZip = buffer[0] === 0x50 && buffer[1] === 0x4b;
  if (!isZip) return { rtfBuffer: buffer, byTheNumbersImage: null };
  const zip = new AdmZip(buffer);
  const entries = zip.getEntries();
  const rtfEntry = entries.find((e) => !e.isDirectory && (e.entryName === 'TXT.rtf' || e.entryName.endsWith('.rtf')));
  if (!rtfEntry) throw new Error('No RTF file found inside the uploaded package (expected TXT.rtf or another .rtf file)');
  const rtfBuffer = rtfEntry.getData();
  const imageEntry = entries.find((e) => !e.isDirectory && IMAGE_EXT.test(e.entryName));
  let byTheNumbersImage = null;
  if (imageEntry) {
    const ext = (imageEntry.entryName.match(IMAGE_EXT) || [])[1] || 'png';
    const mimeType = ext.toLowerCase() === 'jpg' || ext.toLowerCase() === 'jpeg' ? 'image/jpeg' : ext.toLowerCase() === 'png' ? 'image/png' : ext.toLowerCase() === 'gif' ? 'image/gif' : ext.toLowerCase() === 'webp' ? 'image/webp' : 'image/png';
    byTheNumbersImage = { buffer: imageEntry.getData(), mimeType };
  }
  return { rtfBuffer, byTheNumbersImage };
}
module.exports = getRtfFromUpload;
