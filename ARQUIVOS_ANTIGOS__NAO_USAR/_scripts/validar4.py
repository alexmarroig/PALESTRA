"""Comparação caractere por caractere. Nada é normalizado: espaços, pontuação,
acentos, caixa, travessões, reticências, aspas e ordem contam."""
import zipfile, re, html, sys, difflib

UP = '/root/.claude/uploads/165b0804-c9de-557d-912f-f20e921197e2/'
FONTE_B3 = '/home/user/PALESTRA/BLOCO3__TEXTO_FINAL_LIMPO.docx'
FONTE_B4 = UP + 'f0f45174-BLOCO_4___CONCLUS_O___O_QUE_PRECISA_ENTRAR_NO_CUIDADO__25_07.docx'
V20 = [('Quais prejuízos esse comportamento está produzindo?',
        'Quais custos esse comportamento está produzindo?')]

def paras(fp, tirar_marcas_novas=False, tirar_marcas_antigas=False, reps=()):
    xml = zipfile.ZipFile(fp).read('word/document.xml').decode('utf-8')
    body = xml.split('<w:body>')[1]
    out = []
    for p in re.findall(r'<w:p[ >].*?</w:p>', body, re.S):
        t = ''.join(html.unescape(x) for x in re.findall(r'<w:t[^>]*>(.*?)</w:t>', p))
        if not t.strip(): continue
        s = t.strip()
        if tirar_marcas_novas and s.startswith('{') and s.endswith('}'): continue
        if tirar_marcas_antigas and s.startswith('[') and s.endswith(']'): continue
        for a, b in reps: t = t.replace(a, b)
        out.append(t)
    return out

def cmp(rotulo, esperado, obtido):
    ok = esperado == obtido
    print(f'{rotulo}: {"PASS" if ok else "FAIL"}')
    if not ok:
        d = [l for l in difflib.unified_diff(esperado, obtido, 'ESPERADO', 'OBTIDO',
                                             lineterm='', n=0)
             if not l.startswith(('---', '+++', '@@'))]
        for l in d[:12]: print('   ', l[:150])
    return ok

falhou = False

# 1 e 2 — LIMPO x ENTONAÇÃO
for b in ('3', '4'):
    limpo = paras(f'BLOCO{b}__FINAL_LIMPO.docx')
    ento  = paras(f'BLOCO{b}__COM_ENTONACAO.docx', tirar_marcas_novas=True)
    falhou |= not cmp(f'BLOCO {b} — LIMPO x ENTONAÇÃO', limpo, ento)

# 3 — Bloco 3 fonte x final
fonte3 = paras(FONTE_B3)
final3 = paras('BLOCO3__FINAL_LIMPO.docx')
falhou |= not cmp('BLOCO 3 — FONTE x FINAL', fonte3, final3)

# 4 — Bloco 4 Versão 2.0 x final (da fonte tiram-se só as indicações cênicas)
fonte4 = paras(FONTE_B4, tirar_marcas_antigas=True, reps=V20)
final4 = paras('BLOCO4__FINAL_LIMPO.docx')
falhou |= not cmp('BLOCO 4 — VERSÃO 2.0 x FINAL', fonte4, final4)

# trechos obrigatórios
TRECHOS = [
 "Ele passava horas rolando o feed, já não gostava do que via e percebia os prejuízos. Mesmo assim, quando fechava o aplicativo, em poucos segundos estava ali novamente.",
 "Hoje, o uso problemático de redes sociais ainda não possui uma categoria diagnóstica específica formalmente reconhecida.",
 "O que esse comportamento oferece ou ajuda a afastar?",
 "Quando uma resposta produz consequências imediatas como essas — seja acrescentando algo desejável, seja afastando algo aversivo — aumenta a probabilidade de que ela seja repetida em situações semelhantes.",
 "É justamente aqui que a vinheta de Rafael ajuda a tornar essa ideia concreta.",
 "a experiência de solidão dentro do lugar em que ele esperava encontrar vínculo.",
 "Quais custos esse comportamento está produzindo?",
 "O risco define a prioridade, mas não esgota a formulação.",
 "Se esse comportamento for reduzido ou interrompido, o que ficará sem resposta?",
 "Novas respostas precisam ser praticadas, fortalecidas e se tornar disponíveis justamente quando a pessoa mais precisa delas.",
 "Às vezes, o psiquiatra encontrará uma parte do caso, o psicólogo encontrará outra, e o diálogo responsável entre eles tornará visível aquilo que, separadamente, poderia permanecer fora do cuidado.",
 "diante de um comportamento aditivo, não observem apenas quanto a pessoa usa ou como fazê-la parar.",
 "Porque o comportamento que vemos pode ser apenas a ponta do iceberg.",
 "Porque o cuidado não termina quando o comportamento para.",
 "Ele se sustenta quando a vida se torna mais ampla, mais valiosa e volta a oferecer outras respostas.",
]
inteiro = ' '.join(final4)
faltam = [t for t in TRECHOS if t not in inteiro]
print(f'\nBLOCO 4 — 15 trechos obrigatórios: {"PASS" if not faltam else "FAIL"} ({15-len(faltam)}/15)')
for t in faltam: print('   falta:', t[:90])
falhou |= bool(faltam)

# nenhuma marcação sobrando nos LIMPO
for b in ('3', '4'):
    ps = paras(f'BLOCO{b}__FINAL_LIMPO.docx')
    sobra = [p for p in ps if p.strip().startswith(('{', '['))]
    print(f'BLOCO {b} — LIMPO sem marcações: {"PASS" if not sobra else "FAIL " + str(sobra[:3])}')
    falhou |= bool(sobra)

sys.exit(1 if falhou else 0)
