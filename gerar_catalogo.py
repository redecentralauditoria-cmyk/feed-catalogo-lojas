"""
Gera o catalogo.csv do catalogo da Meta a partir das planilhas do FarmaPRO.

REGRAS DE SEGURANCA (nao alterar sem falar com o Marcel):
- So entram produtos cujo "Nome Grupo" comeca com HPC (higiene pessoal e cosmeticos), mais os
  grupos LB de nao-medicamento em GRUPOS_LB_LIBERADOS (leites/nutricao e dermocosmeticos,
  autorizados pelo Marcel em 02/10/2026). Os demais grupos LB / GENER / SIMIL sao MEDICAMENTO
  pelo cadastro oficial e NUNCA entram.
  Esse filtro e a trava real: o medicamento nao existe no arquivo, entao a IA nao
  tem como informar preco dele pelo catalogo.
- Preco = coluna P.M.C. do CADASTRO (ja com desconto Fidelidade).
"""
import openpyxl, csv, glob, os, re, sys, unicodedata
from datetime import datetime

# Todas as fontes vem da pasta dedicada deste projeto (o Marcel cola os arquivos aqui manualmente,
# nunca direto da rede) - pasta com espacos duplos de proposito, e o nome real da pasta no disco.
AUTOMATIZ_DIR = r"C:\Users\marcel.pereira\CLAUDE\AUTOMATIZ. ATENDIM.  WHATSAPP  LOJAS"
DIARIO_DIR = os.path.join(AUTOMATIZ_DIR, "ATUALIZAÇÃO DIARIA - CATALOGO PROD")  # o Marcel salva aqui o cadastro e o estoque do dia
CADASTRO_DIR = DIARIO_DIR
ESTOQUE_DIR = DIARIO_DIR
BANCO_FOTOS = os.path.join(AUTOMATIZ_DIR, "BANCO_IMAGENS_INSTABUY.xlsx")
SAIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalogo.csv")
LINK_LOJA = "https://loja.redecentralfarma.com.br/"

GRUPOS_LB_LIBERADOS = {
    "LB LIBERADOS NUTRICAO",             # leites infantis, Ensure, Nutren, Sustagen...
    "LB LIBERADOS DERMO",                # dermocosmeticos (CeraVe, La Roche, Avene...)
    "LB EXCESSO ESTOQUE HPC/LIBERADOS",  # HPC/liberados em queima de estoque (NAN sem lactose etc.)
}

# itens de medicamento/OTC que caem nos grupos liberados acima, ou classificados como tratamento na loja online
EXCLUIR_NOME = ["violeta genciana", "clotrimix", "targifor", "zincopro", "valda", "bye bye fever", "canfora 1un"]

# Marca propria (decisao do Marcel 07/10/2026): todo item do checklist de marca propria entra no catalogo do WhatsApp
# (vaga garantida, mesmo fora dos mais vendidos e mesmo em grupo SIMIL LIBERADOS); o resto das 500 vagas e
# completado pelos mais vendidos. Fica de fora a aba MEDICAMENTO do checklist (politica do WhatsApp nao permite remedio no catalogo).
CHECKLIST_MP = r"C:\Users\marcel.pereira\Desktop\AUDITORIAS\01-CLAUDE\VERIFICAR NAS LOJAS\01- CHECKLIST - ULTIMO\Checklist_MarcaPropria_RedeCentral_2026_25.xlsx"
MP_LIGADO = True    # ligado em 10/10/2026 com OK do Marcel (regra: >=5 un, sem Natusplant, sem aba MEDICAMENTO, MP sem foto entra com logo)
MP_CODIGOS = set()
MP_NOMES = {}          # codigo -> nome completo do checklist (usado quando o item entra sem foto)
MP_MIN_UN = 5          # regra em discussao 07-10/10/2026: so tem vaga garantida o item MP que vendeu >= 5 un no fechamento
FOTO_ERRADA = {"115550", "116282", "61956"}  # banco Instabuy tem foto/titulo de OUTRO produto (Eudora, Integralmedica, Milnutri) - nao usar
# item MP sem foto entra no catalogo com esta imagem provisoria ate o MKT subir a foto (a Meta exige image_link)
FOTO_PROVISORIA = "https://raw.githubusercontent.com/redecentralauditoria-cmyk/feed-catalogo-lojas/main/sem_foto.png"
if os.environ.get("MP_SIMULAR") == "1":
    MP_LIGADO = True   # so para simulacao em pasta temporaria; nunca publicar assim sem OK do Marcel
if MP_LIGADO and os.path.exists(CHECKLIST_MP):
    _wb = openpyxl.load_workbook(CHECKLIST_MP, data_only=True, read_only=True)
    for _ws in _wb.worksheets:
        if "Instru" in _ws.title or "Geral" in _ws.title or _ws.title.strip().upper() == "MEDICAMENTO":
            continue
        for _r in _ws.iter_rows(values_only=True):
            if _r and len(_r) > 3 and _r[1] and str(_r[1]).strip().isdigit():
                if "natus" in unicodedata.normalize("NFKD", str(_r[3]).lower()):   # Natusplant fora (decisao 07/10)
                    continue
                MP_CODIGOS.add(str(_r[1]).strip().lstrip("0"))
                MP_NOMES[str(_r[1]).strip().lstrip("0")] = re.sub(r"\s+", " ", str(_r[3]).replace(" ", " ")).strip().title()
elif MP_LIGADO:
    print(f"  AVISO: checklist de marca propria nao encontrado ({CHECKLIST_MP}) - catalogo sai so pelos mais vendidos")


def mais_recente(pasta, padrao):
    achados = glob.glob(os.path.join(pasta, padrao))
    if not achados:
        sys.exit(f"ERRO: nenhum arquivo '{padrao}' em {pasta}")
    return max(achados, key=os.path.getmtime)


def norm(s):
    s = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").strip().lower()


print("Gerando catalogo...\n")

# ---------- 1. fotos publicas (banco Instabuy) ----------
wb = openpyxl.load_workbook(BANCO_FOTOS, data_only=True, read_only=True)
ws = wb[wb.sheetnames[0]]
fotos = {}
for r in ws.iter_rows(min_row=2, values_only=True):
    cod, nome, marca, img = r[1], r[4], r[5], r[7]
    if img and img != "Sem imagem" and cod:
        fotos[str(cod).strip()] = (nome, marca, img)
print(f"  fotos publicas disponiveis: {len(fotos)}")

# ---------- 2. estoque (so entra o que existe em alguma loja) ----------
arq_est = mais_recente(ESTOQUE_DIR, "Estoque das Filiais em *.XLSX")
wb = openpyxl.load_workbook(arq_est, data_only=True, read_only=True)
ws = wb[wb.sheetnames[0]]
tem_estoque = set()
for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True)):
    cod = r[0]
    try:
        total = float(r[3] or 0)
    except (TypeError, ValueError):
        total = 0
    if cod and total > 0:
        tem_estoque.add(str(cod).strip().lstrip("0"))
print(f"  estoque: {os.path.basename(arq_est)} -> {len(tem_estoque)} com saldo")

# ---------- 3. cadastro (preco + grupo) ----------
arq_cad = mais_recente(CADASTRO_DIR, "CADASTRO DE PRODUTOS ATUALIZADO*.xlsx")
wb = openpyxl.load_workbook(arq_cad, data_only=True, read_only=True)
ws = wb["CADASTRO"]
print(f"  cadastro: {os.path.basename(arq_cad)}")

linhas, medicamento, sem_foto, sem_estoque, excluidos = [], 0, 0, 0, []
fotos_provisorias = []
for r in ws.iter_rows(min_row=2, values_only=True):
    ean, cod, desc, lab, grupo, _cf, _cm, _uc, pmc, _fr, _m1, _m2, _pr, _dm, lsit = r[:15]
    if not cod:
        continue
    grupo = (grupo or "").strip()
    if not (grupo.startswith("HPC") or grupo in GRUPOS_LB_LIBERADOS
            or str(cod).strip().lstrip("0") in MP_CODIGOS):                 # <-- trava de medicamento (marca propria do checklist passa)
        medicamento += 1
        continue
    if str(lsit).strip().lower() != "true":
        continue
    try:
        preco = float(pmc or 0)
    except (TypeError, ValueError):
        preco = 0
    if preco <= 0:
        continue

    k = str(cod).strip()
    if k.lstrip("0") not in tem_estoque:
        sem_estoque += 1
        continue
    k0 = k.lstrip("0")
    if k in FOTO_ERRADA or k not in fotos:
        if k0 in MP_CODIGOS:      # marca propria entra mesmo sem foto (decisao Marcel 10/10): imagem provisoria + nome do checklist
            nome_ib, marca, img = MP_NOMES.get(k0) or desc, "Rede Central", FOTO_PROVISORIA
            fotos_provisorias.append(k0)
        else:
            sem_foto += 1
            continue
    else:
        nome_ib, marca, img = fotos[k]
    titulo = (nome_ib or desc or "").strip()
    if any(x in norm(titulo) or x in norm(desc or "") for x in EXCLUIR_NOME):
        excluidos.append(titulo)
        continue
    # segmentos fora do catalogo (decisao do Marcel 06/10/2026): agulhas, seringas e testes (palavra inteira no nome)
    if k0 not in MP_CODIGOS and re.search(r"\b(agulhas?|seringas?|testes?)\b", norm(titulo)):
        excluidos.append(titulo)
        continue

    linhas.append({
        "id": k,
        "title": titulo[:200],
        "description": titulo[:9999],
        "availability": "in stock",
        "condition": "new",
        "price": f"{preco:.2f} BRL",
        "link": LINK_LOJA,
        "image_link": img,
        "brand": (marca or "").strip(),
    })

# ---------- 4. gravar ----------
# Ordem alfabetica (por titulo) so afeta a ORDEM DE EXIBICAO no catalogo do WhatsApp/loja;
# quais produtos entram continua decidido pelos filtros acima (e, no top 490, pelo faturamento).
cols = ["id", "title", "description", "availability", "condition",
        "price", "link", "image_link", "brand"]
linhas.sort(key=lambda l: norm(l["title"]))
with open(SAIDA, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    w.writerows(linhas)

# ---------- 5. top 490 mais vendidos (limite de 500 produtos do WhatsApp Business app) ----------
# Ranking por faturamento (quantidade do fechamento mensal mais recente x P.M.C. do cadastro). Ficam de fora o CD (filial 000) e as
# linhas de movimentacao de estoque (devolucao / remanejo), que nao sao venda ao cliente.
VENDAS_DIR = os.path.join(AUTOMATIZ_DIR, "VENDAS GERAL")
SAIDA_TOP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalogo_top500.csv")
LIMITE_WHATSAPP = 500
IGNORAR_VENDEDOR = {"devolucao mercadorias", "remanejos entre lojas"}

# Pega o fechamento do mes mais recente; se existir a versao "editada colunas" (CD ja removido pelo Marcel), ela tem prioridade.
arqs_ven = glob.glob(os.path.join(VENDAS_DIR, "20??-?? Vendas geral*.xlsx"))
if not arqs_ven:
    sys.exit(f"ERRO: nenhum arquivo 'AAAA-MM Vendas geral*.xlsx' em {VENDAS_DIR}")
arq_ven = max(arqs_ven, key=lambda a: (os.path.basename(a)[:7], "editada" in norm(os.path.basename(a))))
wb = openpyxl.load_workbook(arq_ven, data_only=True, read_only=True)
ws = wb["Plan1"] if "Plan1" in wb.sheetnames else wb[wb.sheetnames[0]]
linhas_ven = ws.iter_rows(values_only=True)
cab = [norm(c or "") for c in next(linhas_ven)]
i_fil, i_prod, i_qtd = cab.index("filial"), cab.index("produto"), cab.index("quantidade")
i_bruto = cab.index("valor bruto")  # regra do Marcel (06/10/2026): faturamento = Quantidade x Valor Bruto
i_vend = cab.index("nome vendedor") if "nome vendedor" in cab else None

vendido, qtd_un = {}, {}
for r in linhas_ven:
    cod, fil = r[i_prod], str(r[i_fil] or "").strip()
    if not cod or not fil or fil.lstrip("0") == "":
        continue
    if i_vend is not None and norm(r[i_vend] or "") in IGNORAR_VENDEDOR:
        continue
    try:
        q = float(r[i_qtd] or 0)
        v = float(r[i_bruto] or 0)
    except (TypeError, ValueError):
        continue
    k = str(cod).strip().lstrip("0")
    vendido[k] = vendido.get(k, 0) + q * v
    qtd_un[k] = qtd_un.get(k, 0) + q

def faturamento(l):
    return vendido.get(l["id"].lstrip("0"), 0)

eh_mp = lambda l: l["id"].lstrip("0") in MP_CODIGOS and qtd_un.get(l["id"].lstrip("0"), 0) >= MP_MIN_UN
top_mp = sorted((l for l in linhas if eh_mp(l)), key=faturamento, reverse=True)[:LIMITE_WHATSAPP]
top_outros = sorted((l for l in linhas if not eh_mp(l) and faturamento(l) > 0), key=faturamento, reverse=True)
top = top_mp + top_outros[:LIMITE_WHATSAPP - len(top_mp)]
print(f"  marca propria sem foto (imagem provisoria): {len([l for l in top if l['image_link'] == FOTO_PROVISORIA])}")
print(f"  marca propria no catalogo WhatsApp: {len(top_mp)} | mais vendidos (demais): {len(top) - len(top_mp)}")
top.sort(key=lambda l: norm(l["title"]))  # exibicao alfabetica; entrada no top 500 já foi decidida acima por faturamento
with open(SAIDA_TOP, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    w.writerows(top)

print(f"""
  descartados:
    medicamento (LB/GENER/SIMIL): {medicamento}
    sem estoque em nenhuma loja:  {sem_estoque}
    sem foto publica:             {sem_foto}
    excluidos por nome:           {len(excluidos)} {excluidos if excluidos else ''}

  CATALOGO GERADO: {len(linhas)} produtos
  arquivo: {SAIDA}

  CATALOGO WHATSAPP (mais vendidos): {len(top)} produtos
  ranking de vendas: {os.path.basename(arq_ven)}
  arquivo: {SAIDA_TOP}
""")
