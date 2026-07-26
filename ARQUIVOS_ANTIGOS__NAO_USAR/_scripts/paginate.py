"""Simulador de paginação: lê o .docx e estima onde o Word quebra as páginas.

Não substitui uma renderização real — é uma aproximação calibrada para Calibri
em A4, útil para detectar páginas vazias, marcações órfãs e citações separadas
da sua deixa. Larguras de caractere medidas empiricamente para Calibri.
"""
import zipfile, re, html, sys, glob

TWIP_IN = 1440
A4_W, A4_H = 11906, 16838
MARGIN = 1440          # padrão do docx-js
HDR_FTR = 900          # cabeçalho + rodapé + folga

def parse(fp):
    z = zipfile.ZipFile(fp)
    xml = z.read('word/document.xml').decode('utf-8')
    body = xml.split('<w:body>')[1]
    landscape = 'w:orient="landscape"' in xml
    cols = 2 if re.search(r'<w:cols[^>]*w:num="2"', xml) else 1
    m = re.search(r'<w:pgMar[^>]*w:left="(\d+)"[^>]*w:right="(\d+)"', xml)
    left, right = (int(m.group(1)), int(m.group(2))) if m else (MARGIN, MARGIN)
    m = re.search(r'<w:pgMar[^>]*w:top="(\d+)"[^>]*w:bottom="(\d+)"', xml)
    top, bottom = (int(m.group(1)), int(m.group(2))) if m else (MARGIN, MARGIN)
    pw, ph = (A4_H, A4_W) if landscape else (A4_W, A4_H)
    text_w = (pw - left - right - (400 if cols == 2 else 0)) / cols
    text_h = (ph - top - bottom - HDR_FTR) * cols

    paras = []
    for p in re.findall(r'<w:p[ >].*?</w:p>|<w:p/>', body, re.S):
        txt = ''.join(html.unescape(t) for t in re.findall(r'<w:t[^>]*>(.*?)</w:t>', p))
        sz = int((re.search(r'<w:sz w:val="(\d+)"', p) or [0, '24'])[1])
        after = int((re.search(r'w:after="(\d+)"', p) or [0, '0'])[1])
        before = int((re.search(r'w:before="(\d+)"', p) or [0, '0'])[1])
        ind = int((re.search(r'<w:ind[^>]*w:left="(\d+)"', p) or [0, '0'])[1])
        paras.append({
            'text': txt,
            'break': '<w:br w:type="page"/>' in p,
            'keepNext': '<w:keepNext/>' in p,
            'keepLines': '<w:keepLines/>' in p,
            'sz': sz, 'after': after, 'before': before, 'indent': ind,
        })
    return paras, text_w, text_h

def height(p, text_w):
    pt = p['sz'] / 2
    # Calibri: largura média ~0.48em; caixa-alta e negrito puxam para cima
    char_w = pt * 20 * 0.475
    usable = text_w - p['indent']
    per_line = max(10, int(usable / char_w))
    n = max(1, -(-len(p['text']) // per_line)) if p['text'] else 1
    line_h = pt * 20 * 1.18
    return n * line_h + p['after'] + p['before'], n

def layout(fp):
    paras, tw, th = parse(fp)
    pages, cur, used = [], [], 0.0
    i = 0
    while i < len(paras):
        p = paras[i]
        if p['break']:
            pages.append((cur, used)); cur, used = [], 0.0; i += 1; continue
        # grupo keepNext: o parágrafo e tudo que ele prende
        group, j = [], i
        while j < len(paras) and paras[j]['keepNext'] and not paras[j]['break']:
            group.append(paras[j]); j += 1
        if j < len(paras) and not paras[j]['break']:
            group.append(paras[j]); j += 1
        if not group:
            group, j = [p], i + 1
        gh = sum(height(g, tw)[0] for g in group)
        if used + gh > th and used > 0:
            pages.append((cur, used)); cur, used = [], 0.0
        cur.extend(group); used += gh; i = j
    if cur: pages.append((cur, used))
    return pages, th

def report(fp):
    pages, th = layout(fp)
    print(f'\n=== {fp}  —  {len(pages)} páginas ===')
    problems = []
    for n, (ps, used) in enumerate(pages, 1):
        fill = used / th * 100
        spoken = [p for p in ps if p['text'] and not p['text'].startswith('{')]
        flag = ''
        if fill < 45 and n != len(pages):
            flag = '  <<< POUCO APROVEITADA'
            problems.append(f'página {n}: {fill:.0f}% de ocupação')
        if not spoken:
            flag = '  <<< SEM TEXTO FALADO'
            problems.append(f'página {n}: só marcações')
        print(f'  p.{n}: {fill:5.1f}% | {len(ps):2d} parágrafos | inicia: {(ps[0]["text"] or "(vazio)")[:52]}{flag}')
    return len(pages), problems

if __name__ == '__main__':
    files = sys.argv[1:] or sorted(glob.glob('*.docx'))
    allp = {}
    for f in files:
        n, probs = report(f)
        allp[f] = (n, probs)
    print('\n\n########## RESUMO ##########')
    for f, (n, probs) in allp.items():
        print(f'{f:<46} {n:>2} pág.  {"OK" if not probs else "; ".join(probs)}')
