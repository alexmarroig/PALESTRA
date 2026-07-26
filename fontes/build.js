const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  PageNumber, PageBreak, AlignmentType,
} = require('docx');

const dir = __dirname;
const MARK = '9C4221';   // direção de fala
const QUOTE = '1F3864';  // relato de participante
const CUT = '767171';    // trecho cortável

// ---------------------------------------------------------------------------
// Convenções do arquivo-fonte
//   linha 1        título do bloco
//   {DIREÇÃO}      linha só de marcação        -> keepNext (cola na fala seguinte)
//   {*DIREÇÃO}     marcação só no ENSAIO
//   {X — explica}  a explicação cai na versão de PALCO
//   >texto         relato literal de participante
//   **texto**      negrito
//   ===            quebra de página estratégica
//   ~texto         trecho cortável (versão de 40 min)
//   texto:         deixa curta -> keepNext (nunca se separa do que vem depois)
// ---------------------------------------------------------------------------

const isCue = t => /:$/.test(t) && t.replace(/\*\*/g, '').length <= 60;

function runs(text, base, stage) {
  const out = [];
  text.split(/(\{[^}]*\}|\*\*[^*]+\*\*)/).filter(s => s.length).forEach(seg => {
    if (seg.startsWith('{')) {
      let body = seg.slice(1, -1);
      if (body.startsWith('*')) { if (stage) return; body = body.slice(1); }
      if (stage) body = body.split(' — ')[0];
      out.push(new TextRun({ text: `{${body}}`, bold: true, color: MARK, size: 19, font: 'Calibri' }));
    } else if (seg.startsWith('**')) {
      out.push(new TextRun(Object.assign({}, base, { text: seg.slice(2, -2), bold: true, font: 'Calibri' })));
    } else {
      out.push(new TextRun(Object.assign({}, base, { text: seg, font: 'Calibri' })));
    }
  });
  return out;
}

function paragraphs(lines, { stage = false, showCuts = false } = {}) {
  const paras = [];
  for (let i = 1; i < lines.length; i++) {
    let t = lines[i].trim();
    if (!t) continue;

    if (t === '===') {
      paras.push(new Paragraph({ children: [new PageBreak()] }));
      continue;
    }

    const cuttable = t.startsWith('~');
    if (cuttable) t = t.slice(1).trim();

    // Marcação isolada: gruda no que vem depois, nunca quebra no meio.
    if (t.startsWith('{') && t.endsWith('}')) {
      let body = t.slice(1, -1);
      if (body.startsWith('*')) { if (stage) continue; body = body.slice(1); }
      if (stage) body = body.split(' — ')[0];
      paras.push(new Paragraph({
        spacing: { before: 80, after: 100 },
        keepNext: true, keepLines: true,
        children: [new TextRun({ text: `{${body}}`, bold: true, color: MARK, size: 19, font: 'Calibri' })],
      }));
      continue;
    }

    const base = { size: 24 };
    if (cuttable && showCuts) base.color = CUT;

    if (t.startsWith('>')) {
      paras.push(new Paragraph({
        spacing: { before: 80, after: 160 },
        indent: { left: 480 },
        keepLines: true,               // a citação não se parte entre páginas
        children: runs(t.slice(1).trim(),
          Object.assign({ italics: true, color: cuttable && showCuts ? CUT : QUOTE }, { size: 24 }), stage),
      }));
      continue;
    }

    paras.push(new Paragraph({
      spacing: { after: 160 },
      keepLines: true,
      keepNext: isCue(t),             // só deixas curtas prendem o parágrafo seguinte
      children: runs(t, base, stage),
    }));
  }
  return paras;
}

function build(srcFile, outFile, headerLabel, opts = {}) {
  const lines = fs.readFileSync(path.join(dir, srcFile), 'utf8').split('\n');
  const title = (opts.title || lines[0]).trim();

  const children = [
    new Paragraph({
      spacing: { after: 360 },
      keepNext: true,
      children: [new TextRun({ text: title, bold: true, size: 30, font: 'Calibri' })],
    }),
    ...paragraphs(lines, opts),
  ];

  const doc = new Document({
    sections: [{
      properties: { titlePage: true },
      headers: {
        first: new Header({ children: [new Paragraph('')] }),
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: headerLabel, color: '808080', size: 18, font: 'Calibri' })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              children: [`${headerLabel} | página `, PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES],
              color: '808080', size: 18, font: 'Calibri',
            })],
          })],
        }),
      },
      children,
    }],
  });
  return Packer.toBuffer(doc).then(b => fs.writeFileSync(path.join(dir, outFile), b));
}

module.exports = { build };

if (require.main === module) {
  const jobs = [
    ['bloco3.txt', 'BLOCO_3__PARTITURA_ENSAIO.docx', 'Bloco 3 — Achados da pesquisa', {}],
    ['bloco3.txt', 'BLOCO_3__PARTITURA_PALCO.docx', 'Bloco 3 — Achados da pesquisa', { stage: true }],
    ['bloco4.txt', 'BLOCO_4__PARTITURA_ENSAIO.docx', 'Bloco 4 — Conclusão', {}],
    ['bloco4.txt', 'BLOCO_4__PARTITURA_PALCO.docx', 'Bloco 4 — Conclusão', { stage: true }],
  ];
  Promise.all(jobs.map(j => build(...j))).then(() => console.log('ok'));
}
