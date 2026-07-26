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

// Partitura geral: concatena os blocos. Remove marcações soltas no fim de cada
// bloco (eram elas que caíam sozinhas numa página depois da quebra).
const geral = [];
BLOCKS.forEach(([f], i) => {
  let lines = fs.readFileSync(path.join(dir, f), 'utf8').split('\n').map(l => l.trim()).filter(Boolean);
  while (lines.length && /^\{.*\}$/.test(lines[lines.length - 1])) lines.pop();
  if (i) geral.push('===');
  geral.push(`**${lines[0]}**`);
  geral.push(...lines.slice(1));
});
const head = 'PARTITURA GERAL DA PALESTRA — 25 DE JULHO DE 2026\n';
fs.writeFileSync(path.join(dir, 'geral.txt'), head + geral.join('\n'));

const legenda = 'TRECHOS EM CINZA NÃO DEVEM SER FALADOS NA VERSÃO DE 40 MINUTOS';
fs.writeFileSync(path.join(dir, 'geral40.txt'),
  'PALESTRA — VERSÃO DE 40 MINUTOS (CORTES MARCADOS)\n' + `**${legenda}**\n` + geral.join('\n'));

const jobs = [];
BLOCKS.forEach(([src, label]) => {
  const stem = src.replace('.txt', '').toUpperCase();
  jobs.push([src, `${stem}__ENSAIO.docx`, label, { mode: 'ensaio' }]);
  jobs.push([src, `${stem}__PALCO.docx`, label, { mode: 'palco' }]);
});
jobs.push(['bloco3.txt', 'BLOCO3__TEXTO_FINAL_LIMPO.docx', 'Bloco 3 — texto final', { mode: 'limpo' }]);
jobs.push(['bloco4.txt', 'BLOCO4__TEXTO_FINAL_LIMPO.docx', 'Bloco 4 — texto final', { mode: 'limpo' }]);
jobs.push(['geral.txt', 'PALESTRA__PARTITURA_GERAL__ENSAIO.docx', 'Partitura geral — ensaio', { mode: 'ensaio' }]);
jobs.push(['geral40.txt', 'PALESTRA__40MIN__CORTES_MARCADOS.docx', 'Versão 40 min — cortes marcados', { mode: 'ensaio', showCuts: true }]);
jobs.push(['geral.txt', 'PALESTRA__VERSAO_40MIN__LIMPA.docx', 'Versão 40 min — limpa',
           { mode: 'limpo', dropCuts: true, title: 'PALESTRA — VERSÃO DE 40 MINUTOS (TEXTO JÁ CORTADO)' }]);
jobs.push(['mapa.txt', 'PALESTRA__MAPA_DE_RECUPERACAO.docx', 'Mapa de recuperação',
           { mode: 'palco', landscape: true, size: 18, spacing: 70 }]);

Promise.all(jobs.map(j => build(...j))).then(() => console.log('gerados:', jobs.length));
