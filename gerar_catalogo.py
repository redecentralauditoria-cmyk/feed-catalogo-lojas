"""
Gera o catalogo.csv do catalogo da Meta a partir das planilhas do FarmaPRO.

REGRAS DE SEGURANCA (nao alterar sem falar com o Marcel):
- So entram produtos cujo "Nome Grupo" comeca com HPC (higiene pessoal e cosmeticos).
  Os grupos LB / GENER / SIMIL sao MEDICAMENTO pelo cadastro oficial e NUNCA entram.
  Esse filtro e a trava real: o medicamento nao existe no arquivo, entao a IA nao
  tem como informar preco dele pelo catalogo.
- Preco = coluna P.M.C. do CADASTRO (ja com desconto Fidelidade).
"""
import openpyxl, csv, glob, os, re, sys, unicodedata
from datetime import datetime

BASE = r"C:\Users\marcel.pereira\Desktop\AUDITORIAS\01-CLAUDE"
CADASTRO_DIR = os.path.join(BASE, "ESTOQUE - VENDAS LOJAS", "VENDAS")
ESTOQUE_DIR = os.path.join(BASE, "ESTOQUE - VENDAS LOJAS", "ESTOQUE")
BANCO_FOTOS = os.path.join(BASE, "IMAGENS-CATALOGO", "BANCO_IMAGENS_INSTABUY.xlsx")
SAIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalogo.csv")
LINK_LOJA = "https://loja.redecentralfarma.com.br/"

EXCLUIR_NOME = ["violeta genciana"]  # classificado como "para infeccoes" na loja online


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
for r in ws.iter_rows(min_row=2, values_only=True):
    ean, cod, desc, lab, grupo, _cf, _cm, _uc, pmc, _fr, _m1, _m2, _pr, _dm, lsit = r[:15]
    if not cod:
        continue
    grupo = (grupo or "").strip()
    if not grupo.startswith("HPC"):      # <-- trava de medicamento
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
    if k not in fotos:
        sem_foto += 1
        continue

    nome_ib, marca, img = fotos[k]
    titulo = (nome_ib or desc or "").strip()
    if any(x in norm(titulo) for x in EXCLUIR_NOME):
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
cols = ["id", "title", "description", "availability", "condition",
        "price", "link", "image_link", "brand"]
with open(SAIDA, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    w.writerows(linhas)

print(f"""
  descartados:
    medicamento (LB/GENER/SIMIL): {medicamento}
    sem estoque em nenhuma loja:  {sem_estoque}
    sem foto publica:             {sem_foto}
    excluidos por nome:           {len(excluidos)} {excluidos if excluidos else ''}

  CATALOGO GERADO: {len(linhas)} produtos
  arquivo: {SAIDA}
""")
