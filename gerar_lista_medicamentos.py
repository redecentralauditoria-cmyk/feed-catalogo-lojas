"""
Gera a LISTA SEMANAL de precos de medicamentos de venda livre (MIP) + vitaminas/suplementos
para carregar em Meta Business Agent > Conhecimento > Arquivos e links (o catalogo do WhatsApp
nao pode ter medicamento - Politica Comercial da Meta).

REGRAS (combinadas com o Marcel em 03/10/2026):
- So entram os grupos de venda livre (GRUPOS_MIP). Nunca psicotropico, controlado, antibiotico
  ou tarja vermelha/preta (esses grupos simplesmente nao estao na lista).
- Preco = P.M.C. do CADASTRO (ja com desconto Fidelidade). Ranking = quantidade do fechamento
  mensal mais recente x P.M.C. Precisa ter estoque em alguma loja.
- Toda semana: apagar a lista anterior em Arquivos e links e carregar a nova.
"""
import openpyxl, glob, os, re, sys, unicodedata
from datetime import date, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

PASTA = r"C:\Users\marcel.pereira\CLAUDE\AUTOMATIZ. ATENDIM.  WHATSAPP  LOJAS"
DIARIO = os.path.join(PASTA, "ATUALIZAÇÃO DIARIA - CATALOGO PROD")  # cadastro e estoque do dia
TOP_N = 300

GRUPOS_MIP = {
    "LB O.T.C", "LB CONTINUOS O.T.C", "GENER O.T.C", "SIMIL O.T.C",   # medicamentos de venda livre
    "LB LIBERADOS GERAL", "LB LIBERADOS CONTINUOS", "SIMIL LIBERADOS",  # vitaminas, suplementos, colirio lubrificante etc.
}
# dentro dos grupos acima, apresentacoes que exigem receita
EXCLUIR = [r"\b(IBUPRO\w*|BUPROVIL|IBUFRAN|ALIVIUM)\b.*\b(600|800)\s?MG", r"PAREGORICO"]
# Medicamentos de MARCA PROPRIA (aba MEDICAMENTO do checklist): nao podem ir no catalogo do WhatsApp, entram SEMPRE nesta lista
# (decisao Marcel 10/10/2026): Lufree x2, Lactulona x2, Gastroclean, Gelomag aerosol
FORCAR = ["107026", "107025", "116240", "116239", "107903", "110417"]
IGNORAR_VENDEDOR = {"devolucao mercadorias", "remanejos entre lojas"}


def norm(s):
    s = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").strip().lower()


def mais_recente(padrao):
    achados = glob.glob(os.path.join(DIARIO, padrao))
    if not achados:
        sys.exit(f"ERRO: nenhum arquivo '{padrao}' em {PASTA}")
    return max(achados, key=os.path.getmtime)


# estoque
arq_est = mais_recente("Estoque das Filiais em *.XLSX")
tem_estoque = set()
for r in openpyxl.load_workbook(arq_est, read_only=True, data_only=True).worksheets[0].iter_rows(min_row=2, values_only=True):
    try:
        if r[0] and float(r[3] or 0) > 0:
            tem_estoque.add(str(r[0]).strip().lstrip("0"))
    except (TypeError, ValueError):
        pass

# nomes bonitos (banco Instabuy), quando houver
nomes = {}
for r in openpyxl.load_workbook(os.path.join(PASTA, "BANCO_IMAGENS_INSTABUY.xlsx"), read_only=True, data_only=True).worksheets[0].iter_rows(min_row=2, values_only=True):
    if r[1] and r[4]:
        nomes[str(r[1]).strip()] = str(r[4]).strip()

# cadastro
arq_cad = mais_recente("CADASTRO DE PRODUTOS ATUALIZADO*.xlsx")
itens = {}
for r in openpyxl.load_workbook(arq_cad, read_only=True, data_only=True)["CADASTRO"].iter_rows(min_row=2, values_only=True):
    cod, desc, grupo, pmc, lsit = r[1], r[2], (r[4] or "").strip(), r[8], r[14]
    if not cod or grupo not in GRUPOS_MIP or str(lsit).strip().lower() != "true":
        continue
    d = str(desc or "").upper()
    if any(re.search(p, d) for p in EXCLUIR):
        continue
    try:
        preco = float(pmc or 0)
    except (TypeError, ValueError):
        continue
    k = str(cod).strip()
    if preco <= 0 or k.lstrip("0") not in tem_estoque:
        continue
    itens[k.lstrip("0")] = (nomes.get(k) or d.title(), preco)

# vendas (ranking)
arqs_ven = sorted(a for a in glob.glob(os.path.join(PASTA, "VENDAS GERAL", "20??-?? Vendas geral*.xlsx")) if "editada" not in norm(a))
ws = openpyxl.load_workbook(arqs_ven[-1], read_only=True, data_only=True)
ws = ws["Plan1"] if "Plan1" in ws.sheetnames else ws.worksheets[0]
linhas = ws.iter_rows(values_only=True)
cab = [norm(c or "") for c in next(linhas)]
i_fil, i_prod, i_qtd = cab.index("filial"), cab.index("produto"), cab.index("quantidade")
i_vend = cab.index("nome vendedor") if "nome vendedor" in cab else None
qtd = {}
for r in linhas:
    cod, fil = r[i_prod], str(r[i_fil] or "").strip()
    if not cod or not fil or fil.lstrip("0") == "":
        continue
    if i_vend is not None and norm(r[i_vend] or "") in IGNORAR_VENDEDOR:
        continue
    try:
        q = float(r[i_qtd] or 0)
    except (TypeError, ValueError):
        continue
    k = str(cod).strip().lstrip("0")
    qtd[k] = qtd.get(k, 0) + q

top = sorted((k for k in itens if qtd.get(k, 0) > 0 and k not in FORCAR), key=lambda k: qtd[k] * itens[k][1], reverse=True)[:TOP_N - len(FORCAR)]
faltou = [c for c in FORCAR if c not in itens]
top += [c for c in FORCAR if c in itens]
if faltou:
    print("AVISO: marca propria sem preco/estoque/cadastro, ficou fora da lista:", faltou)
top.sort(key=lambda k: norm(itens[k][0]))

# PDF
hoje = date.today()
saida = os.path.join(DIARIO, f"precos_medicamentos_venda_livre_{hoje:%d-%m-%Y}.pdf")   # lista DIARIA desde 10/10/2026

st = ParagraphStyle("c", fontName="Helvetica", fontSize=9, leading=12)
stt = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, leading=17, spaceAfter=6)
S = [Paragraph("Rede Central Farma - Lista de preços de medicamentos de venda livre e vitaminas", stt),
     Paragraph(f"<b>Preços de {hoje:%d/%m/%Y}.</b> Esta lista SUBSTITUI qualquer lista de preços anterior. "
               "Preços do cadastro Fidelidade (já com desconto). Use apenas para informar preço e disponibilidade quando o "
               "cliente pedir o produto pelo nome - nunca para orientar uso, dose, indicação ou substituição (isso é sempre "
               "com o farmacêutico). Não é oferta nem promoção. Produto que não estiver nesta lista nem no catálogo: "
               "confirmar o valor com a loja.", st),
     Spacer(1, 8)]
dados = [["Produto", "Preço (R$)"]] + [[Paragraph(itens[k][0], st), f"{itens[k][1]:.2f}".replace(".", ",")] for k in top]
t = Table(dados, colWidths=[14.5 * cm, 2.5 * cm], repeatRows=1)
t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                       ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9),
                       ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D5D8DD")), ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                       ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F3F5")])]))
S.append(t)
SimpleDocTemplate(saida, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.5 * cm, bottomMargin=1.5 * cm,
                  title="Precos medicamentos venda livre - Rede Central Farma").build(S)

print(f"cadastro: {os.path.basename(arq_cad)} | estoque: {os.path.basename(arq_est)} | vendas: {os.path.basename(arqs_ven[-1])}")
print(f"elegiveis: {len(itens)} | na lista: {len(top)}")
print(f"arquivo: {saida}")
