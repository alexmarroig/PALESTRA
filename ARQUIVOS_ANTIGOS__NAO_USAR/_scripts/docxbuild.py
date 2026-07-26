"""Monta os .docx da palestra preservando os originais caractere por caractere.

Abertura, Bloco 1 e Bloco 2 NÃO são regerados a partir de texto: os parágrafos
originais são reaproveitados como XML, o que preserva palavras, pontuação,
espaçamento e a numeração automática das listas. As marcações cênicas entram
como parágrafos novos, inseridos entre os originais.

Blocos 3 e 4 são gerados a partir das fontes .txt congeladas.
"""
import zipfile, re, html, copy, os

UP = '/root/.claude/uploads/165b0804-c9de-557d-912f-f20e921197e2/'
ORIG = {
    'abertura': UP + '8eae8446-ABERTURA_25_DE_JULHO_DE_2026.docx',
    'bloco1':   UP + '69c0d315-BLOCO_1__Amplica__o_da_Escuta__25_07.docx',
    'bloco2':   UP + 'd63db454-BLOCO_2__FUN__O_DO_COMPORTAMENTO__25_07.docx',
}
MARK, QUOTE = '9C4221', '1F3864'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))

# --------------------------------------------------------------------------- #
# parágrafos
# --------------------------------------------------------------------------- #
def para_mark(text):
    return ('<w:p><w:pPr><w:keepNext/><w:keepLines/>'
            '<w:spacing w:before="60" w:after="80"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/>'
            f'<w:color w:val="{MARK}"/><w:sz w:val="19"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def para_text(text, *, bold_title=False, quote=False, keep_next=False,
              size=24, spacing=140, indent=0):
    ppr = '<w:pPr>'
    if keep_next: ppr += '<w:keepNext/>'
    ppr += '<w:keepLines/>'
    if indent: ppr += f'<w:ind w:left="{indent}"/>'
    ppr += f'<w:spacing w:after="{spacing}"/></w:pPr>'
    runs = ''
    for seg in re.split(r'(\*\*[^*]+\*\*)', text):
        if not seg: continue
        b = seg.startswith('**')
        body = seg[2:-2] if b else seg
        rpr = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        if b or bold_title: rpr += '<w:b/>'
        if quote: rpr += '<w:i/>'
        rpr += f'<w:color w:val="{QUOTE if quote else "000000"}"/>'
        rpr += f'<w:sz w:val="{30 if bold_title else size}"/>'
        runs += f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(body)}</w:t></w:r>'
    return f'<w:p>{ppr}{runs}</w:p>'

PAGEBREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# --------------------------------------------------------------------------- #
# leitura dos originais
# --------------------------------------------------------------------------- #
def original_paragraphs(key, drop=()):
    """Devolve os <w:p> originais, intactos, sem os parágrafos administrativos."""
    xml = zipfile.ZipFile(ORIG[key]).read('word/document.xml').decode('utf-8')
    body = xml.split('<w:body>')[1].rsplit('</w:body>')[0]
    out = []
    for p in re.findall(r'<w:p[ >].*?</w:p>|<w:p/>', body, re.S):
        t = ''.join(html.unescape(x) for x in re.findall(r'<w:t[^>]*>(.*?)</w:t>', p))
        if not t.strip(): continue
        if any(t.strip().startswith(d) for d in drop): continue
        out.append((t, p))
    return out

def apply_replacements(pxml, reps):
    """Substitui texto que pode estar espalhado por vários <w:r>.

    Mapeia cada caractere do parágrafo ao run que o contém, localiza o trecho e
    reescreve apenas os runs atingidos — a formatação dos demais fica intacta.
    """
    if not reps: return pxml
    for alvo, novo in reps:
        runs = list(re.finditer(r'(<w:t[^>]*>)(.*?)(</w:t>)', pxml, re.S))
        if not runs: break
        textos = [html.unescape(m.group(2)) for m in runs]
        full = ''.join(textos)
        i = full.find(alvo)
        if i < 0: continue
        j = i + len(alvo)
        # posição inicial de cada run no texto completo
        starts, acc = [], 0
        for t in textos:
            starts.append(acc); acc += len(t)
        partes = []
        colocado = False
        for k, t in enumerate(textos):
            a, b = starts[k], starts[k] + len(t)
            if b <= i or a >= j:            # run fora do trecho
                partes.append(t); continue
            pre  = t[:max(0, i - a)]
            post = t[max(0, j - a):] if j > a else t
            if not colocado:
                partes.append(pre + novo + (post if j <= b else ''))
                colocado = True
            else:
                partes.append(post if j <= b else '')
        out, pos = [], 0
        for k, m in enumerate(runs):
            out.append(pxml[pos:m.start()])
            out.append(m.group(1) + esc(partes[k]) + m.group(3))
            pos = m.end()
        out.append(pxml[pos:])
        pxml = ''.join(out)
    return pxml

def set_props(pxml, keep_next):
    """Injeta keepLines (e keepNext quando pedido) sem tocar no resto."""
    props = ('<w:keepNext/>' if keep_next else '') + '<w:keepLines/>'
    if '<w:pPr>' in pxml:
        return pxml.replace('<w:pPr>', '<w:pPr>' + props, 1)
    return re.sub(r'^(<w:p[^>]*>)', r'\1<w:pPr>' + props + '</w:pPr>', pxml, count=1)

IS_CUE = lambda t: t.strip().endswith(':') and len(t.strip()) <= 60

# --------------------------------------------------------------------------- #
# escrita
# --------------------------------------------------------------------------- #
HDR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:hdr {NS}><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr>'
       '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="808080"/>'
       '<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{L}</w:t></w:r></w:p></w:hdr>')
FTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:ftr {NS}><w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
       '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr>'
       '<w:t xml:space="preserve">{L} | página </w:t></w:r>'
       '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
       '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
       '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr>'
       '<w:t xml:space="preserve"> de </w:t></w:r>'
       '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>'
       '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')
EMPTY_HDR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:hdr {NS}><w:p/></w:hdr>')

def write_docx(out_path, base_key, paragraphs, label, landscape=False):
    """Usa um original como container (leva styles.xml e numbering.xml junto)."""
    src = zipfile.ZipFile(ORIG[base_key])
    entries = {n: src.read(n) for n in src.namelist()}
    doc = entries['word/document.xml'].decode('utf-8')

    sect = re.search(r'<w:sectPr.*?</w:sectPr>', doc, re.S)
    sect_xml = sect.group(0) if sect else '<w:sectPr/>'
    # cabeçalho/rodapé + primeira página sem cabeçalho
    sect_inner = re.sub(r'<w:headerReference[^>]*/>|<w:footerReference[^>]*/>|<w:titlePg/>', '', sect_xml)
    sect_inner = sect_inner.replace('<w:sectPr>',
        '<w:sectPr><w:headerReference w:type="first" r:id="rIdHdrF"/>'
        '<w:headerReference w:type="default" r:id="rIdHdr"/>'
        '<w:footerReference w:type="default" r:id="rIdFtr"/><w:titlePg/>', 1)
    if landscape:
        sect_inner = re.sub(r'<w:pgSz[^/]*/>', '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>', sect_inner)
        sect_inner = re.sub(r'<w:pgMar[^/]*/>',
            '<w:pgMar w:top="560" w:right="620" w:bottom="560" w:left="620" w:header="280" w:footer="280"/>', sect_inner)
        sect_inner = sect_inner.replace('<w:cols', '<w:colsX', 1)
        sect_inner = sect_inner.replace('<w:titlePg/>', '<w:titlePg/><w:cols w:num="2" w:space="400"/>', 1)
        sect_inner = re.sub(r'<w:colsX[^/]*/>', '', sect_inner)

    body = ''.join(paragraphs) + sect_inner
    head = doc.split('<w:body>')[0]
    entries['word/document.xml'] = (head + '<w:body>' + body + '</w:body></w:document>').encode('utf-8')

    entries['word/header1.xml'] = EMPTY_HDR.encode('utf-8')
    entries['word/header2.xml'] = HDR.replace('{L}', esc(label)).encode('utf-8')
    entries['word/footer1.xml'] = FTR.replace('{L}', esc(label)).encode('utf-8')

    rels = entries['word/_rels/document.xml.rels'].decode('utf-8')
    add = ('<Relationship Id="rIdHdrF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
           '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>'
           '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>')
    entries['word/_rels/document.xml.rels'] = rels.replace('</Relationships>', add + '</Relationships>').encode('utf-8')

    ct = entries['[Content_Types].xml'].decode('utf-8')
    for n, t in [('/word/header1.xml', 'header+xml'), ('/word/header2.xml', 'header+xml'),
                 ('/word/footer1.xml', 'footer+xml')]:
        if n not in ct:
            ct = ct.replace('</Types>',
                f'<Override PartName="{n}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{t}"/></Types>')
    entries['[Content_Types].xml'] = ct.encode('utf-8')

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in entries.items():
            z.writestr(n, d)
