const fs = require('fs');
const path = require('path');
const { build } = require('./build.js');

const dir = __dirname;
const BLOCKS = [
  ['abertura.txt', 'Abertura'],
  ['bloco1.txt', 'Bloco 1 — Ampliação da escuta'],
  ['bloco2.txt', 'Bloco 2 — Função do comportamento'],
  ['bloco3.txt', 'Bloco 3 — Achados da pesquisa'],
  ['bloco4.txt', 'Bloco 4 — Conclusão'],
];

// Partitura geral: todos os blocos num arquivo, com quebra entre eles.
const geral = [];
BLOCKS.forEach(([f], i) => {
  const lines = fs.readFileSync(path.join(dir, f), 'utf8').split('\n').filter(l => l.trim());
  if (i) geral.push('===');
  geral.push(`**${lines[0]}**`);
  geral.push(...lines.slice(1));
});
fs.writeFileSync(path.join(dir, 'geral.txt'), 'PARTITURA GERAL DA PALESTRA — 25 DE JULHO DE 2026\n' + geral.join('\n'));

const jobs = [];
BLOCKS.forEach(([src, label]) => {
  const stem = src.replace('.txt', '').toUpperCase();
  jobs.push([src, `${stem}__ENSAIO.docx`, label, {}]);
  jobs.push([src, `${stem}__PALCO.docx`, label, { stage: true }]);
});
jobs.push(['geral.txt', 'PALESTRA__PARTITURA_GERAL.docx', 'Partitura geral', {}]);
jobs.push(['geral.txt', 'PALESTRA__VERSAO_40MIN.docx', 'Versão de segurança — 40 min', { showCuts: true }]);
jobs.push(['mapa.txt', 'PALESTRA__MAPA_DE_RECUPERACAO.docx', 'Mapa de recuperação', { stage: true }]);

Promise.all(jobs.map(j => build(...j))).then(() => console.log('gerados:', jobs.length));
