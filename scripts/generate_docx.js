/**
 * Docx Master — Node.js + JSZip fallback generator
 * Usage: node generate_docx.js <output.docx> <content.json>
 *
 * Auto-installs JSZip if missing, then generates a .docx from JSON.
 * Works on any platform with Node.js ≥ 12.
 */

const fs = require('fs');
const path = require('path');
const cp = require('child_process');

// ── Auto-install JSZip if missing ────────────────────────────────
let JSZip;
try {
  JSZip = require('jszip');
} catch (e) {
  console.log('[docx-master] jszip not found, installing...');
  try {
    cp.execSync('npm install jszip --no-save', {
      cwd: __dirname,
      stdio: 'pipe',
      timeout: 30000,
    });
    JSZip = require('jszip');
    console.log('[docx-master] jszip installed OK');
  } catch (e2) {
    console.error('[docx-master] ERROR: cannot install jszip.');
    console.error('  Install manually: npm install jszip');
    process.exit(1);
  }
}

// ── CLI ───────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (args.length < 2) {
  console.error('Usage: node generate_docx.js <output.docx> <content.json>');
  process.exit(1);
}

const outPath = args[0];
const content = JSON.parse(fs.readFileSync(args[1], 'utf-8'));

// ── XML helpers ───────────────────────────────────────────────────

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function pNode(text, opts = {}) {
  const bold = opts.bold ? '<w:rPr><w:b/></w:rPr>' : '';
  const sz = opts.sz ? `<w:rPr><w:sz w:val="${opts.sz}"/></w:rPr>` : '';
  const italic = opts.italic ? '<w:rPr><w:i/></w:rPr>' : '';
  const color = opts.color ? `<w:rPr><w:color w:val="${opts.color}"/></w:rPr>` : '';
  const align = opts.align ? `<w:pPr><w:jc w:val="${opts.align}"/></w:pPr>` : '';
  const indent = opts.indent ? `<w:pPr><w:ind w:left="${opts.indent}" w:hanging="360"/></w:pPr>` : '';
  const rPr = bold || sz || italic || color;
  return `<w:p>${align}${indent}<w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}

function headingNode(text, level) {
  const sizes = { 1: 36, 2: 28, 3: 24, 4: 22 };
  const styleIds = { 1: '2', 2: '3', 3: '4', 4: '5' };
  const sz = sizes[level] || 22;
  const sid = styleIds[level] || '5';
  return `<w:p><w:pPr><w:pStyle w:val="${sid}"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr><w:t>${esc(text)}</w:t></w:r></w:p>`;
}

function drawingXml(id, rid, cx, cy) {
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="114300" distR="114300">
    <wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="img${id}"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr><pic:cNvPr id="${id}" name="img${id}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline></w:drawing></w:r>`;
}

// ── Table helpers ─────────────────────────────────────────────────

function tableXml(headers, rows, colWidths) {
  const nCols = headers.length;
  const widths = colWidths || headers.map(() => Math.floor(8304 / nCols));
  let grid = '<w:tblGrid>';
  widths.forEach(w => { grid += `<w:gridCol w:w="${w}"/>`; });
  grid += '</w:tblGrid>';

  let h = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>${grid}`;
  h += trXml(headers, true);
  rows.forEach(r => { h += trXml(r, false); });
  h += '</w:tbl>';
  return h;
}

function trXml(cells, isHeader) {
  let r = '<w:tr>';
  cells.forEach(c => {
    r += '<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr>';
    r += '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>';
    r += isHeader
      ? `<w:r><w:rPr><w:b/></w:rPr><w:t>${esc(c)}</w:t></w:r>`
      : `<w:r><w:t>${esc(c)}</w:t></w:r>`;
    r += '</w:p></w:tc>';
  });
  r += '</w:tr>';
  return r;
}

// ── Build document body ───────────────────────────────────────────

let docBody = '';

if (content.title) {
  docBody += `<w:p><w:pPr><w:pStyle w:val="2"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>${esc(content.title)}</w:t></w:r></w:p>`;
}

let imageIdx = 0;
const images = [];   // { path, rid }

(content.sections || []).forEach(sec => {
  if (sec.heading) {
    docBody += headingNode(sec.heading, sec.level || 1);
  }
  (sec.content || []).forEach(item => {
    switch (item.type) {
      case 'paragraph':
        docBody += pNode(item.text, item);
        break;
      case 'heading':
        docBody += headingNode(item.text, item.level || 2);
        break;
      case 'bullet':
        docBody += pNode(`• ${item.text}`, { indent: '720' });
        break;
      case 'table':
        docBody += tableXml(item.headers, item.rows, item.colWidths);
        docBody += pNode('', {});
        break;
      case 'image':
        imageIdx++;
        const rid = `rIdImg${imageIdx}`;
        const cx = item.width || 5400000;
        const cy = item.height || 3000000;
        docBody += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${drawingXml(100 + imageIdx, rid, cx, cy)}</w:p>`;
        if (item.caption) {
          docBody += pNode(item.caption, { align: 'center', sz: '20', color: '666666' });
        }
        images.push({ path: item.path, rid });
        break;
    }
  });
});

// ── Supporting XML parts ──────────────────────────────────────────

const RELS_ROOT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="2"><w:name w:val="Heading 1"/><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体"/><w:b/><w:sz w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="3"><w:name w:val="Heading 2"/><w:pPr><w:spacing w:before="200" w:after="100"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体"/><w:b/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="4"><w:name w:val="Heading 3"/><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体"/><w:b/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="5"><w:name w:val="Heading 4"/><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体"/><w:b/></w:rPr></w:style>
</w:styles>`;

let docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
images.forEach(img => {
  docRels += `\n  <Relationship Id="${img.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${img.rid.replace('rIdImg', '')}.png"/>`;
});
docRels += '\n</Relationships>';

// ── Package & output ──────────────────────────────────────────────

async function main() {
  const zip = new JSZip();

  // mimetype MUST be first, uncompressed
  zip.file('mimetype', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', { compression: 'STORE' });

  zip.file('_rels/.rels', RELS_ROOT);
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('word/document.xml', `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${docBody}</w:body></w:document>`);
  zip.file('word/styles.xml', STYLES);
  zip.file('word/_rels/document.xml.rels', docRels);

  // Embed images
  let imgNum = 0;
  for (const img of images) {
    imgNum++;
    try {
      const buf = fs.readFileSync(img.path);
      const ext = path.extname(img.path).slice(1) || 'png';
      zip.file(`word/media/image${imgNum}.${ext}`, buf);
    } catch (e) {
      console.error(`[docx-master] WARNING: cannot read image: ${img.path}`);
    }
  }

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, buf);
  console.log(`OK: ${outPath} (${buf.length} bytes, ${images.length} images)`);
}

main().catch(e => { console.error(e); process.exit(1); });
