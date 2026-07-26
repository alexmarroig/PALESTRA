"""Gera os quatro documentos finais dos Blocos 3 e 4.

O texto falado é extraído run a run das fontes oficiais e reescrito verbatim —
nada é redigitado. Só a formatação (fonte, cor, negrito, espaçamento) é nova.
"""
import zipfile, re, html, os

AQUI = os.path.dirname(os.path.abspath(__file__))
UP = '/root/.claude/uploads/165b0804-c9de-557d-912f-f20e921197e2/'
REPO = '/home/user/PALESTRA/'

FONTE_B3 = REPO + 'BLOCO3__TEXTO_FINAL_LIMPO.docx'
FONTE_B4 = UP + 'f0f45174-BLOCO_4___CONCLUS_O___O_QUE_PRECISA_ENTRAR_NO_CUIDADO__25_07.docx'

# Versão 2.0: única diferença em relação ao arquivo de 25/07
B4_V20 = [('Quais prejuízos esse comportamento está produzindo?',
           'Quais custos esse comportamento está produzindo?')]

PRETO, MARCA, CITACAO = '000000', '9C4221', '1F3864'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

# --------------------------------------------------------------------------- #
def ler(fp, reps=()):
    """Parágrafos como lista de runs (texto, negrito). Texto nunca é redigitado."""
    xml = zipfile.ZipFile(fp).read('word/document.xml').decode('utf-8')
    body = xml.split('<w:body>')[1]
    out = []
    for p in re.findall(r'<w:p[ >].*?</w:p>', body, re.S):
        runs = []
        for r in re.findall(r'<w:r[ >].*?</w:r>', p, re.S):
            t = ''.join(html.unescape(x) for x in re.findall(r'<w:t[^>]*>(.*?)</w:t>', r, re.S))
            if not t: continue
            runs.append([t, bool(re.search(r'<w:b/>|<w:b w:val="(?:1|true)"/>', r))])
        if not runs or not ''.join(r[0] for r in runs).strip(): continue
        out.append(runs)
    # substituições da Versão 2.0, aplicadas sobre o parágrafo inteiro
    for alvo, novo in reps:
        for runs in out:
            full = ''.join(r[0] for r in runs)
            if alvo in full and len(runs) == 1:
                runs[0][0] = full.replace(alvo, novo)
    return out

texto = lambda runs: ''.join(r[0] for r in runs)
E_MARCA_ANTIGA = lambda runs: texto(runs).strip().startswith('[') and texto(runs).strip().endswith(']')
E_CITACAO = lambda runs: texto(runs).strip().startswith(('“', '"'))
E_DEIXA = lambda runs: texto(runs).strip().endswith(':') and len(texto(runs).strip()) <= 70

# --------------------------------------------------------------------------- #
def p_texto(runs, *, titulo=False, citacao=False, negrito=False, keep_next=False):
    ppr = '<w:pPr>'
    if keep_next: ppr += '<w:keepNext/>'
    ppr += '<w:keepLines/>'
    if citacao: ppr += '<w:ind w:left="454" w:right="284"/>'
    ppr += f'<w:spacing w:after="{240 if titulo else 160}" w:line="300" w:lineRule="auto"/></w:pPr>'
    xml = ''
    for t, b in runs:
        rpr = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        if b or titulo or negrito: rpr += '<w:b/>'
        if citacao: rpr += '<w:i/>'
        rpr += f'<w:color w:val="{CITACAO if citacao else PRETO}"/>'
        rpr += f'<w:sz w:val="{32 if titulo else 24}"/>'
        xml += f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>'
    return f'<w:p>{ppr}{xml}</w:p>'

def p_marca(t):
    return ('<w:p><w:pPr><w:keepNext/><w:keepLines/>'
            '<w:spacing w:before="80" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/>'
            f'<w:color w:val="{MARCA}"/><w:sz w:val="19"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(t)}</w:t></w:r></w:p>')

PAGEBREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# --------------------------------------------------------------------------- #
def montar(paras, marcas=(), negritos=(), grupo_final=None):
    corpo, usadas, no_grupo = [], set(), False
    vivos = [r for r in paras if not E_MARCA_ANTIGA(r)]
    ultimo = len(vivos) - 1
    for i, runs in enumerate(vivos):
        t = texto(runs).strip()
        if grupo_final and t.startswith(grupo_final):
            no_grupo = True
        for k, (anchor, m) in enumerate(marcas):
            if k in usadas or not t.startswith(anchor): continue
            usadas.add(k); corpo.append(p_marca(m)); break
        corpo.append(p_texto(
            runs, titulo=(i == 0), citacao=E_CITACAO(runs),
            negrito=any(t.startswith(n) for n in negritos),
            keep_next=(i == 0) or E_DEIXA(runs) or (no_grupo and i < ultimo)))
    faltando = [a for k, (a, _) in enumerate(marcas) if k not in usadas]
    if faltando: raise SystemExit(f'ÂNCORA NÃO ENCONTRADA: {faltando}')
    return corpo

HDR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:hdr {NS}><w:p/></w:hdr>')
FTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:ftr {NS}><w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
       '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr>'
       '<w:t xml:space="preserve">{L} — </w:t></w:r>'
       '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
       '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
       '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr>'
       '<w:t xml:space="preserve">/</w:t></w:r>'
       '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>'
       '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')

SECT = ('<w:sectPr><w:footerReference w:type="default" r:id="rIdFtr"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
        'w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')

DOC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:document {NS}><w:body>{{B}}{SECT}</w:body></w:document>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        '</Relationships>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '</Relationships>')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')

def escrever(nome, corpo, label):
    with zipfile.ZipFile(os.path.join(AQUI, nome), 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/document.xml', DOC.replace('{B}', ''.join(corpo)))
        z.writestr('word/_rels/document.xml.rels', RELS)
        z.writestr('word/footer1.xml', FTR.replace('{L}', esc(label)))

# --------------------------------------------------------------------------- #
import entonacao as E

NEG3 = ['O tempo de tela mostra', 'Quando existe sofrimento ou prejuízo']
NEG4 = ['O que costuma acontecer antes?', 'O que esse comportamento oferece ou ajuda a afastar?',
        'Quais custos esse comportamento está produzindo?',
        'Se esse comportamento for reduzido ou interrompido',
        'O risco define a prioridade', 'Mas essas palavras julgavam',
        'Porque o comportamento que vemos pode ser',
        'Porque o cuidado não termina', 'Ele se sustenta quando a vida']

b3 = ler(FONTE_B3)
b4 = ler(FONTE_B4, B4_V20)

escrever('BLOCO3__FINAL_LIMPO.docx',     montar(b3, (), NEG3), 'Bloco 3')
escrever('BLOCO4__FINAL_LIMPO.docx',     montar(b4, (), NEG4), 'Bloco 4')
escrever('BLOCO3__COM_ENTONACAO.docx',   montar(b3, E.BLOCO3, NEG3), 'Bloco 3')
escrever('BLOCO4__COM_ENTONACAO.docx',
         montar(b4, E.BLOCO4, NEG4, grupo_final='Se vocês saírem daqui com uma ideia'), 'Bloco 4')
print('4 documentos gerados')
