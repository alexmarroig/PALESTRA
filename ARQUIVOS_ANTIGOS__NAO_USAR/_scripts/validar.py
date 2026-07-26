"""Validação textual: comparação EXATA (pontuação, maiúsculas, acentos, números
de lista, ordem e divisão de parágrafos). Só remove marcações cênicas, cabeçalho,
rodapé e o título do bloco."""
import zipfile, re, html, sys, difflib

UP = '/root/.claude/uploads/165b0804-c9de-557d-912f-f20e921197e2/'
ORIG = {
    'ABERTURA': UP + '8eae8446-ABERTURA_25_DE_JULHO_DE_2026.docx',
    'BLOCO 1':  UP + '69c0d315-BLOCO_1__Amplica__o_da_Escuta__25_07.docx',
    'BLOCO 2':  UP + 'd63db454-BLOCO_2__FUN__O_DO_COMPORTAMENTO__25_07.docx',
}
ABERTURA_FIX = [
    ('Boa tarde pessoal, tudo bem?', 'Boa tarde, pessoal, tudo bem?'),
    ('Então vamos lá..', 'Então vamos lá.'),
    ('Instituto brasileiros do comportamento',
     'Instituto Brasiliense de Análise do Comportamento — IBAC'),
]
TITULOS = ('BLOCO 1', 'BLOCO 2', 'BLOCO 3', 'BLOCO 4', 'ABERTURA', 'TEMPO Na Abertura',
           'PARTITURA GERAL', 'PALESTRA —')

def paras(fp, numbered=True):
    """Texto falado, parágrafo a parágrafo, com o número da lista quando houver."""
    xml = zipfile.ZipFile(fp).read('word/document.xml').decode('utf-8')
    body = xml.split('<w:body>')[1]
    out, counters = [], {}
    for p in re.findall(r'<w:p[ >].*?</w:p>|<w:p/>', body, re.S):
        t = ''.join(html.unescape(x) for x in re.findall(r'<w:t[^>]*>(.*?)</w:t>', p))
        if not t.strip(): continue
        if t.strip().startswith('{') and t.strip().endswith('}'): continue   # marcação
        if any(t.strip().startswith(x) for x in TITULOS): continue
        m = re.search(r'<w:numId w:val="(\d+)"', p)
        if m and numbered:
            counters[m.group(1)] = counters.get(m.group(1), 0) + 1
            t = f'[{counters[m.group(1)]}] ' + t
        out.append(t)
    return out

def fatia(lista, ini, fim=None):
    """Recorta os parágrafos de um bloco dentro de um documento combinado."""
    i = next(k for k, x in enumerate(lista) if x.strip().startswith(ini))
    if fim is None: return lista[i:]
    j = next(k for k, x in enumerate(lista) if k > i and x.strip().startswith(fim))
    return lista[i:j]

def esperado(bloco):
    if bloco in ORIG:
        ps = paras(ORIG[bloco])
        if bloco == 'ABERTURA':
            novo = []
            for x in ps:
                for a, b in ABERTURA_FIX: x = x.replace(a, b)
                novo.append(x)
            ps = novo
        return ps
    return paras(f'{bloco.replace(" ", "")}__TEXTO_FINAL_LIMPO.docx')

MARCOS = {  # primeira frase de cada bloco, para recortar nos documentos combinados
    'ABERTURA': 'Boa tarde',
    'BLOCO 1': 'Para compreender uma relação',
    'BLOCO 2': 'Agora, a questão deixa de ser',
    'BLOCO 3': 'Enquanto eu desenvolvia essa pesquisa',
    'BLOCO 4': 'Para concluir, eu quero voltar',
}
ORDEM = ['ABERTURA', 'BLOCO 1', 'BLOCO 2', 'BLOCO 3', 'BLOCO 4']
STEM = {'ABERTURA': 'ABERTURA', 'BLOCO 1': 'BLOCO1', 'BLOCO 2': 'BLOCO2',
        'BLOCO 3': 'BLOCO3', 'BLOCO 4': 'BLOCO4'}
COMBI = ['PALESTRA__TEXTO_FINAL_LIMPO_COMPLETO.docx',
         'PALESTRA__PARTITURA_GERAL__ENSAIO.docx',
         'PALESTRA__PARTITURA_GERAL__PALCO.docx']

falhas = 0
for b in ORDEM:
    ref = esperado(b)
    versoes = {f'{STEM[b]}__{s}.docx': paras(f'{STEM[b]}__{s}.docx')
               for s in ('TEXTO_FINAL_LIMPO', 'ENSAIO', 'PALCO')}
    for c in COMBI:
        todos = paras(c)
        prox = ORDEM[ORDEM.index(b) + 1] if b != ORDEM[-1] else None
        versoes[c] = fatia(todos, MARCOS[b], MARCOS[prox] if prox else None)
    ruins = []
    for nome, ps in versoes.items():
        if ps != ref:
            ruins.append(nome)
            if falhas < 3:
                d = [l for l in difflib.unified_diff(ref, ps, 'ESPERADO', nome, lineterm='', n=0)
                     if not l.startswith(('---', '+++', '@@'))]
                print(f'\n  DIVERGÊNCIA em {nome}:')
                for l in d[:8]: print('    ', l[:140])
    status = 'PASS' if not ruins else f'FAIL ({", ".join(ruins)})'
    if ruins: falhas += 1
    print(f'{b}: {status}   [{len(ref)} parágrafos, {sum(len(x) for x in ref)} caracteres]')

sys.exit(1 if falhas else 0)
