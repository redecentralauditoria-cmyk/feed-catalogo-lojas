"""
TABELA POR PRINCÍPIO ATIVO (pedido do Marcel 23/10/2026) - igual ao botão "Filtrar Princípio" do FarmaPRO:
para cada princípio ativo + dose + forma, os medicamentos com estoque da rede separados em genérico, referência e similar,
com o preço Fidelidade do dia. Serve para a I.A. sempre oferecer o genérico junto (instrução 6).

De onde vem cada coisa:
- princípio ativo e tipo (Genérico / Novo=referência / Similar...): lista pública da CMED/Anvisa (pasta cmed\\), cruzada pelo
  CÓDIGO DE BARRAS do cadastro (EAN 1/2/3 da CMED). O PREÇO da Anvisa NÃO é usado.
- preço: o mesmo das listas (P.M.C. x (1 - Promoção) ou tablóide do dia - campanhas.py).
Trava: se a dose da descrição do cadastro não aparece na apresentação da CMED, o item fica de fora (código de barras errado).
Genérico sem código na CMED entra pelo nome (o nome do genérico já é a substância). Referência/similar sem código fica de fora.
Gera tabela_principio_ativo_DD-MM-AAAA.pdf para Arquivos e links da I.A.
"""
import openpyxl, glob, os, re, sys, unicodedata
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from campanhas import precos_campanha
CAMPANHA = precos_campanha()


def preco_venda(pmc, promo):
    """Preço real de venda (Marcel 21/10/2026): P.M.C. x (1 - coluna Promoção do cadastro)."""
    p = float(pmc or 0)
    try: d = float(promo or 0)
    except (TypeError, ValueError): d = 0
    if 0 < d < 1: p = p * (1 - d)
    return round(p, 2)


def preco_final(cod, pmc, promo):
    p = preco_venda(pmc, promo)
    try: fator = CAMPANHA.get(int(str(cod).strip()))
    except ValueError: fator = None
    c = round(float(pmc or 0) * fator, 2) if fator else 0
    return c if c > 0 and (p <= 0 or c < p) else p


PASTA = r"C:\Users\marcel.pereira\CLAUDE\AUTOMATIZ. ATENDIM.  WHATSAPP  LOJAS"
DIARIO = os.path.join(PASTA, "ATUALIZAÇÃO DIARIA - CATALOGO PROD")
ESTOQUE_MINIMO = 1
norm = lambda s: "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn").upper().strip()
RE_DOSE = re.compile(r"(\d+(?:[.,]\d+)?(?:\s*/\s*\d+(?:[.,]\d+)?)*)\s*(MCG|MG|G|UI|ML|%)\b")
FORMAS = [(r"\bSL\b", "sublingual"), (r"\bODT\b", "orodispersível"), (r"\b(CR|XR|LP|XL|RETARD|AP)\b.*\bC(PR|PS)\b|\bC(PR|PS)\b.*\b(CR|XR|LP|XL|RETARD)\b", "liberação prolongada"),
          (r"\b(CPR|COMP|DRG|DRAG|CPR REV)\b|\d+CPR|\d+DRG", "comprimido"), (r"\b(CPS|CAPS|CAP)\b|\d+CPS|\d+CAP", "cápsula"),
          (r"\bGTS\b|GOTAS", "gotas"), (r"\bXPE\b|XAROPE", "xarope"), (r"\bSUSP\b", "suspensão"), (r"\bSOL\b|SOLUC", "solução"),
          (r"\bPOM\b|POMADA", "pomada"), (r"\bCREME\b|\bCR\b", "creme"), (r"\bGEL\b", "gel"), (r"\bAMP\b|\bINJ\b|\bFA\b", "injetável"),
          (r"SPRAY|AEROSOL|\bJATO", "spray"), (r"\bSUP\b|SUPOS", "supositório"), (r"\bENV\b|SACHE|\bPO\b", "pó/sachê"), (r"COLIRIO", "colírio")]
TIPO = {"GENERICO": "Genérico", "NOVO": "Referência", "BIOLOGICO": "Referência", "SIMILAR": "Similar"}


def dose_de(desc):
    m = RE_DOSE.search(desc)
    return (re.sub(r"\s+", "", m.group(1)).replace(",", ".") + m.group(2)) if m else ""


def forma_de(desc):
    for rx, nome in FORMAS:
        if re.search(rx, desc): return nome
    return ""


def nums(s):
    """números de uma dose/apresentação, normalizados (5.000 -> 5000; 2,5 -> 2.5), também em x1000 e /1000
    para casar 1G = 1000MG e 75MCG = 0,075MG."""
    s = re.sub(r"(\d)\.(\d{3})(?!\d)", r"\1\2", s)
    out = set()
    for x in re.findall(r"\d+(?:[.,]\d+)?", s):
        v = float(x.replace(",", "."))
        out |= {round(v, 4), round(v * 1000, 4), round(v / 1000, 4)}
    out.discard(0.0)
    return out


def recente(p):
    a = glob.glob(os.path.join(DIARIO, p))
    if not a: sys.exit(f"ERRO: falta {p}")
    return max(a, key=os.path.getmtime)


# ---- CMED: EAN -> (substância, tipo, apresentação)
arq_cmed = max(glob.glob(os.path.join(AQUI, "cmed", "lista_pmc_*.xlsx")), key=os.path.getmtime)
cmed = {}
for i, r in enumerate(openpyxl.load_workbook(arq_cmed, read_only=True)["Lista PMC"].iter_rows(values_only=True)):
    if i == 0:
        cab = [norm(c) for c in r]
        ci = {k: next(j for j, c in enumerate(cab) if c.startswith(k)) for k in ("SUBSTANCIA", "EAN 1", "EAN 2", "EAN 3", "APRESENTACAO", "TIPO DE PRODUTO")}
        continue
    sub, tipo, apr = norm(r[ci["SUBSTANCIA"]] or ""), norm(r[ci["TIPO DE PRODUTO"]] or ""), norm(r[ci["APRESENTACAO"]] or "")
    for k in ("EAN 1", "EAN 2", "EAN 3"):
        e = re.sub(r"\D", "", str(r[ci[k]] or ""))
        if len(e) >= 8: cmed.setdefault(e.lstrip("0"), (sub, TIPO.get(tipo, tipo.title()), apr))

# ---- estoque e cadastro
arq_est, arq_cad = recente("Estoque das Filiais em *.XLSX"), recente("CADASTRO DE PRODUTOS ATUALIZADO*.xlsx")
tem = set()
for r in openpyxl.load_workbook(arq_est, read_only=True, data_only=True).worksheets[0].iter_rows(min_row=2, values_only=True):
    try:
        if r[0] and float(r[3] or 0) >= ESTOQUE_MINIMO: tem.add(str(r[0]).strip().lstrip("0"))
    except (TypeError, ValueError): pass

itens, sem_cmed_gen, cont = [], [], {"ean": 0, "nome_generico": 0, "dose_diferente": 0, "sem_codigo_ref_sim": 0}
for r in openpyxl.load_workbook(arq_cad, read_only=True, data_only=True)["CADASTRO"].iter_rows(min_row=2, values_only=True):
    ean, cod, desc, grupo, pmc, lsit, promo = r[0], r[1], norm(re.sub(r"\s+", " ", str(r[2] or ""))), norm(r[4] or ""), r[8], r[14], r[12]
    if not cod or str(lsit).strip().lower() != "true" or str(cod).strip().lstrip("0") not in tem: continue
    if grupo.startswith(("HPC", "CONSUMO")) or grupo in ("MANIPULADOS", "LB VACINAS"): continue
    try: preco = preco_final(cod, pmc, promo)
    except (TypeError, ValueError): continue
    if preco <= 0: continue
    dose, forma = dose_de(desc), forma_de(desc)
    if not dose: continue
    c = cmed.get(re.sub(r"\D", "", str(ean or "")).lstrip("0"))
    if c:
        sub, tipo, apr = c
        if not (nums(dose) & nums(apr)):
            cont["dose_diferente"] += 1; continue
        cont["ean"] += 1
        if grupo.startswith("GENER"): tipo = "Genérico"
        itens.append([sub, dose, forma, tipo, desc, preco])
    elif grupo.startswith("GENER"):
        sem_cmed_gen.append([desc.split(dose.rstrip("MCGUIL%").split("/")[0][:1] or "#")[0], dose, forma, desc, preco])
    else:
        cont["sem_codigo_ref_sim"] += 1

# genérico sem código: usa a substância que outro genérico de mesmo nome-base já recebeu pela CMED
base_sub = {}
for sub, dose, forma, tipo, desc, preco in itens:
    m = re.match(r"([A-Z +\-]+?)\s+\d", desc)
    if m and tipo == "Genérico": base_sub.setdefault(m.group(1).strip(), sub)
for base, dose, forma, desc, preco in sem_cmed_gen:
    b = base.strip()
    itens.append([base_sub.get(b, b), dose, forma, "Genérico", desc, preco]); cont["nome_generico"] += 1

# ---- agrupa por princípio ativo + dose + forma
grupos = {}
for sub, dose, forma, tipo, desc, preco in itens:
    grupos.setdefault((sub, dose, forma), []).append((tipo, desc, preco))
# só grupos com 2+ itens (é onde existe escolha entre marca e genérico)
grupos = {k: v for k, v in grupos.items() if len(v) >= 2}

hoje = date.today()
st = ParagraphStyle("c", fontName="Helvetica", fontSize=7.5, leading=9.5)
sth = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=8, leading=10, spaceBefore=4)
stt = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, leading=17, spaceAfter=6)
fmt = lambda p: f"{p:.2f}".replace(".", ",")
aviso = ("<b>Tabela por princípio ativo - %s.</b> Igual ao 'Filtrar Princípio' do sistema: cada bloco reúne os medicamentos com o MESMO "
         "princípio ativo, dosagem e forma, separados em Genérico, Referência e Similar, com o preço do cadastro Fidelidade de hoje. "
         "Use para: (1) pedido pela substância = informar SOMENTE o genérico de menor preço (o primeiro da linha Genérico); "
         "(2) pedido pela marca = informar a marca pedida e também o genérico de menor preço do MESMO bloco. Quantidades diferentes "
         "(ex.: 30 ou 60 comprimidos) aparecem no mesmo bloco: informe a mesma quantidade que o cliente pediu. Nunca ofereça item de "
         "outro bloco (outro princípio ativo, outra dose ou outra forma). Nunca oriente uso, dose ou troca." % f"{hoje:%d/%m/%Y}")
S = [Paragraph("Rede Central Farma - Tabela por princípio ativo (genérico x referência x similar)", stt), Paragraph(aviso, st), Spacer(1, 4)]
ordem = {"Genérico": 0, "Referência": 1, "Similar": 2}
for (sub, dose, forma) in sorted(grupos, key=lambda k: (k[0], k[1], k[2])):
    v = grupos[(sub, dose, forma)]
    S.append(Paragraph(f"{sub} {dose} {forma}".strip().replace("&", "&amp;"), sth))
    por = {}
    for tipo, desc, preco in v: por.setdefault(tipo, []).append((preco, desc))
    for tipo in sorted(por, key=lambda t: ordem.get(t, 9)):
        lst = "; ".join(f"{d} R$ {fmt(p)}" for p, d in sorted(por[tipo]))
        S.append(Paragraph(f"<b>{tipo}:</b> {lst}".replace("&", "&amp;").replace("&amp;lt;", "&lt;"), st))

arquivo = f"tabela_principio_ativo_{hoje:%d-%m-%Y}.pdf"
saida = os.path.join(DIARIO, arquivo)
SimpleDocTemplate(saida, pagesize=A4, leftMargin=1.2 * cm, rightMargin=1.2 * cm, topMargin=1.2 * cm, bottomMargin=1.2 * cm,
                  title="Tabela por princípio ativo").build(S)
n = sum(len(v) for v in grupos.values())
print(f"{arquivo}: {len(grupos)} blocos, {n} itens | cmed {os.path.basename(arq_cmed)} | cadastro {os.path.basename(arq_cad)} | {cont}")
