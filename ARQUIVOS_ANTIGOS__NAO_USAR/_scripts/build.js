const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  PageNumber, PageBreak, AlignmentType, PageOrientation,
} = require('docx');

const dir = __dirname;
const MARK = '9C4221';
const QUOTE = '1F3864';
const CUT = '808080';

// mode: 'ensaio' | 'palco' | 'limpo'
//   ensaio -> todas as direções, com explicações
//   palco  -> só as marcadas com {!...}, sem a explicação após " — "
//   limpo  -> nenhuma direção
function markText(body, mode) {
  const essential = body.startsWith('!');
  if (essential) body = body.slice(1);
  const ensaioOnly = body.startsWith('*');
  if (ensaioOnly) body = body.slice(1);
  if (mode === 'limpo') return null;
  if (mode === 'palco' && (!essential || ensaioOnly)) return null;
  if (mode === 'palco') body = body.split(' — ')[0];
  return `{${body}}`;
}

function runs(text, base, mode) {
  const out = [];
  text.split(/(\{[^}]*\}|\*\*[^*]+\*\*)/).filter(s => s.length).forEach(seg => {
    if (seg.startsWith('{')) {
      const m = markText(seg.slice(1, -1), mode);
      if (m) out.push(new TextRun({ text: m, bold: true, color: MARK, size: 19, font: 'Calibri' }));
    } else if (seg.startsWith('**')) {
      out.push(new TextRun(Object.assign({}, base, { text: seg.slice(2, -2), bold: true, font: 'Calibri' })));
    } else {
      out.push(new TextRun(Object.assign({}, base, { text: seg, font: 'Calibri' })));
    }
  });
  return out;
}

const isCue = t => /:$/.test(t.replace(/\{[^}]*\}/g, '').trim()) && t.replace(/\*\*|\{[^}]*\}/g, '').length <= 60;

function paragraphs(lines, o) {
  const { mode = 'ensaio', showCuts = false, dropCuts = false, size = 24, spacing = 140 } = o;
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    let t = lines[i].trim();
    if (!t) continue;
    if (t === '===') { out.push(new Paragraph({ children: [new PageBreak()] })); continue; }

    if (t.startsWith('%')) {            // ponte: só existe na versão cortada
      if (!dropCuts) continue;
      t = t.slice(1).trim();
    }
    let glue = false;
    if (t.startsWith('|')) { glue = true; t = t.slice(1).trim(); }
    const cuttable = t.startsWith('~');
    if (cuttable) { if (dropCuts) continue; t = t.slice(1).trim(); }

    if (t.startsWith('{') && t.endsWith('}')) {
      const m = markText(t.slice(1, -1), mode);
      if (!m) continue;
      void glue;
      out.push(new Paragraph({
        spacing: { before: 60, after: 80 }, keepNext: true, keepLines: true,
        children: [new TextRun({ text: m, bold: true, color: MARK, size: 19, font: 'Calibri' })],
      }));
      continue;
    }

    const isQuote = t.startsWith('>');
    if (isQuote) t = t.slice(1).trim();
    const base = { size, color: cuttable && showCuts ? CUT : (isQuote ? QUOTE : '000000') };
    if (isQuote) base.italics = true;

    out.push(new Paragraph({
      spacing: { after: spacing }, keepLines: true,
      keepNext: glue || (!isQuote && isCue(t)),
      indent: isQuote ? { left: 400 } : undefined,
      children: runs(t, base, mode),
    }));
  }
  return out;
}

function build(srcFile, outFile, label, o = {}) {
  const lines = fs.readFileSync(path.join(dir, srcFile), 'utf8').split('\n');
  const children = [
    new Paragraph({
      spacing: { after: o.landscape ? 200 : 340 }, keepNext: true,
      children: [new TextRun({ text: (o.title || lines[0]).trim(), bold: true, size: o.landscape ? 24 : 30, font: 'Calibri' })],
    }),
    ...paragraphs(lines, o),
  ];

  const props = { titlePage: true };
  if (o.landscape) {
    props.page = { size: { orientation: PageOrientation.LANDSCAPE }, margin: { top: 560, bottom: 560, left: 620, right: 620 } };
    props.column = { count: 2, space: 400 };
  }

  const doc = new Document({
    sections: [{
      properties: props,
      headers: {
        first: new Header({ children: [new Paragraph('')] }),
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: label, color: '808080', size: 18, font: 'Calibri' })] })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ children: [`${label} | página `, PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES],
            color: '808080', size: 18, font: 'Calibri' })] })] }),
      },
      children,
    }],
  });
  return Packer.toBuffer(doc).then(b => fs.writeFileSync(path.join(dir, outFile), b));
}

module.exports = { build };
