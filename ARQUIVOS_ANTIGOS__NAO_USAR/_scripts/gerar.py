import re, os, shutil
from docxbuild import (original_paragraphs, apply_replacements, set_props, para_mark,
                       para_text, write_docx, PAGEBREAK, IS_CUE)
import marcas

OUT = os.path.dirname(os.path.abspath(__file__))

ABERTURA_FIX = [
    ('Boa tarde pessoal, tudo bem?', 'Boa tarde, pessoal, tudo bem?'),
    ('Então vamos lá..', 'Então vamos lá.'),
    ('Instituto brasileiros do comportamento',
     'Instituto Brasiliense de Análise do Comportamento — IBAC'),
]

# --------------------------------------------------------------------------- #
# blocos originais (Abertura, 1, 2): XML preservado, marcações inseridas
# --------------------------------------------------------------------------- #
def blocos_originais(key, marks, mode, reps=()):
    paras = original_paragraphs(key, drop=('BLOCO ', 'ABERTURA', 'TEMPO Na Abertura'))
    out, usados = [], set()
    for txt, pxml in paras:
        for i, (anchor, mark, ess) in enumerate(marks):
            if i in usados or not txt.strip().startswith(anchor): continue
            usados.add(i)
            if mode == 'ensaio':
                out.append(para_mark(mark))
            elif mode == 'palco' and ess:
                out.append(para_mark(mark.split(' — ')[0].rstrip('}') + '}'))
            break
        pxml = apply_replacements(pxml, reps)
        out.append(set_props(pxml, keep_next=IS_CUE(txt)))
    faltando = [m[0] for i, m in enumerate(marks) if i not in usados]
    if faltando:
        raise SystemExit(f'ÂNCORA NÃO ENCONTRADA em {key}: {faltando}')
    return out

TITULO = {'abertura': 'ABERTURA', 'bloco1': 'BLOCO 1 — AMPLIAÇÃO DA ESCUTA',
          'bloco2': 'BLOCO 2 — FUNÇÃO DO COMPORTAMENTO'}

# --------------------------------------------------------------------------- #
# blocos 3 e 4: gerados das fontes .txt congeladas
# --------------------------------------------------------------------------- #
def blocos_txt(src, mode):
    out = []
    for line in open(os.path.join(OUT, src), encoding='utf-8').read().split('\n')[1:]:
        t = line.strip()
        if not t or t == '===': continue
        if t.startswith('%'): continue
        glue = t.startswith('|')
        if glue: t = t[1:].strip()
        if t.startswith('~'): t = t[1:].strip()
        if t.startswith('{') and t.endswith('}'):
            body = t[1:-1]
            ess = body.startswith('!')
            if ess: body = body[1:]
            only_ens = body.startswith('*')
            if only_ens: body = body[1:]
            if mode == 'limpo': continue
            if mode == 'palco' and (not ess or only_ens): continue
            if mode == 'palco': body = body.split(' — ')[0]
            out.append(para_mark('{' + body + '}'))
            continue
        quote = t.startswith('>')
        if quote: t = t[1:].strip()
        out.append(para_text(t, quote=quote, keep_next=glue or (not quote and IS_CUE(t)),
                             indent=400 if quote else 0))
    return out

TITULO_TXT = {'bloco3.txt': 'BLOCO 3 — ACHADOS DA PESQUISA',
              'bloco4.txt': 'BLOCO 4 — CONCLUSÃO: O QUE PRECISA ENTRAR NO CUIDADO'}

def corpo(key, mode):
    """Parágrafos de um bloco, já com título."""
    if key in TITULO:
        ps = blocos_originais(key, getattr(marcas, key.upper()), mode,
                              reps=ABERTURA_FIX if key == 'abertura' else ())
        titulo = TITULO[key]
    else:
        ps = blocos_txt(key, mode)
        titulo = TITULO_TXT[key]
    return [para_text(titulo, bold_title=True, keep_next=True, spacing=340)] + ps

LABEL = {'abertura': 'Abertura', 'bloco1': 'Bloco 1 — Ampliação da escuta',
         'bloco2': 'Bloco 2 — Função do comportamento',
         'bloco3.txt': 'Bloco 3 — Achados da pesquisa', 'bloco4.txt': 'Bloco 4 — Conclusão'}
STEM = {'abertura': 'ABERTURA', 'bloco1': 'BLOCO1', 'bloco2': 'BLOCO2',
        'bloco3.txt': 'BLOCO3', 'bloco4.txt': 'BLOCO4'}
BLOCOS = ['abertura', 'bloco1', 'bloco2', 'bloco3.txt', 'bloco4.txt']
BASE = 'bloco2'          # container: leva numbering.xml das listas

for k in BLOCOS:
    for mode, suf in [('limpo', 'TEXTO_FINAL_LIMPO'), ('ensaio', 'ENSAIO'), ('palco', 'PALCO')]:
        write_docx(os.path.join(OUT, f'{STEM[k]}__{suf}.docx'), BASE, corpo(k, mode), LABEL[k])

for mode, nome, label in [('limpo', 'PALESTRA__TEXTO_FINAL_LIMPO_COMPLETO.docx', 'Palestra — texto final'),
                          ('ensaio', 'PALESTRA__PARTITURA_GERAL__ENSAIO.docx', 'Partitura geral — ensaio'),
                          ('palco', 'PALESTRA__PARTITURA_GERAL__PALCO.docx', 'Partitura geral — palco')]:
    todo = []
    for i, k in enumerate(BLOCOS):
        if i: todo.append(PAGEBREAK)
        ps = corpo(k, mode)
        while ps and '9C4221' in ps[-1]:      # nenhuma marcação solta no fim do bloco
            ps.pop()
        todo += ps
    write_docx(os.path.join(OUT, nome), BASE, todo, label)

# mapa de recuperação
mapa = [l.strip() for l in open(os.path.join(OUT, 'mapa.txt'), encoding='utf-8') if l.strip()]
ps = [para_text(mapa[0], bold_title=True, keep_next=True, spacing=200)]
for t in mapa[1:]:
    if t.startswith('{'):
        ps.append(para_mark(t))
    else:
        ps.append(para_text(t, size=18, spacing=70, keep_next=IS_CUE(t)))
write_docx(os.path.join(OUT, 'PALESTRA__MAPA_DE_RECUPERACAO.docx'), BASE, ps,
           'Mapa de recuperação', landscape=True)

print('gerados:', len(BLOCOS) * 3 + 4)
