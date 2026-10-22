# -*- coding: utf-8 -*-
"""Preços de campanha (tablóides/ofertas do FarmaPRO) em vigor numa data.

O Marcel salva os PDFs dos tablóides em ...\\ATUALIZAÇÃO DIARIA - CATALOGO PROD\\TABELAS DESCONTOS CAMPANHAS.
Cada PDF traz "VALIDADE dd/mm/aaaa A dd/mm/aaaa" e linhas
"CÓDIGO DESCRIÇÃO  PREÇO VISTA  DESCONTO%  PREÇO PROMOÇÃO".
No sistema o preço do tablóide VALE NO LUGAR da coluna Promoção do cadastro
(ex.: Tadalafila 065602: cadastro 71,99% = R$ 17,99; tablóide 86% = R$ 8,99).
"""
import glob, os, re, datetime
import fitz  # PyMuPDF

PASTA = os.path.join(r"C:\Users\marcel.pereira\CLAUDE\AUTOMATIZ. ATENDIM.  WHATSAPP  LOJAS",
                     "ATUALIZAÇÃO DIARIA - CATALOGO PROD", "TABELAS DESCONTOS CAMPANHAS")

RE_VAL = re.compile(r"VALIDADE\s+(\d{2})/(\d{2})/(\d{4})\s+A\s+(\d{2})/(\d{2})/(\d{4})")
RE_LIN = re.compile(r"^\s*(\d{4,7})\s+(.+?)\s+([\d.]+,\d{2})\s+([\d.]+,\d+)%\s+([\d.]+,\d{2})\s*$")


def _num(s):
    return float(s.replace(".", "").replace(",", "."))


def ler_pdf(arq):
    """-> (inicio, fim, {codigo: (descricao, preco_promocao, preco_vista)})"""
    txt = "\n".join(p.get_text() for p in fitz.open(arq))
    m = RE_VAL.search(txt)
    if not m:
        return None, None, {}
    d1, m1, a1, d2, m2, a2 = map(int, m.groups())
    itens = {}
    # o PyMuPDF pode quebrar a linha em pedaços: junta tudo e reparte por código
    for lin in txt.splitlines():
        r = RE_LIN.match(lin)
        if r:
            itens[int(r.group(1))] = (r.group(2).strip(), _num(r.group(5)), _num(r.group(3)))
    if not itens:
        plano = re.sub(r"\s+", " ", txt)
        for r in re.finditer(r"(\d{6})\s+(.+?)\s+([\d.]+,\d{2})\s+([\d.]+,\d+)%\s+([\d.]+,\d{2})", plano):
            itens[int(r.group(1))] = (r.group(2).strip(), _num(r.group(5)), _num(r.group(3)))
    return datetime.date(a1, m1, d1), datetime.date(a2, m2, d2), itens


def precos_campanha(data=None, verbose=True):
    """{codigo: fator} das campanhas válidas na data (padrão: hoje); fator = preço promoção / preço vista.
    Usar fator x P.M.C. do cadastro: itens "(F)" têm P.M.C. por fração (cartela/unidade) e o tablóide
    traz o preço da caixa inteira (Nevralgex: P.M.C. 8,436 x 10 frações = 84,36 no tablóide).
    Se 2 campanhas tiverem o mesmo item, fica o MENOR fator (é o que o caixa cobra)."""
    data = data or datetime.date.today()
    precos = {}
    for arq in sorted(glob.glob(os.path.join(PASTA, "*.pdf"))):
        ini, fim, itens = ler_pdf(arq)
        nome = os.path.basename(arq)
        if ini is None:
            if verbose: print(f"  campanha SEM VALIDADE (ignorada): {nome}")
            continue
        vale = ini <= data <= fim
        if verbose:
            print(f"  campanha {ini:%d/%m}-{fim:%d/%m} {'EM VIGOR' if vale else 'fora da data'}: {nome} ({len(itens)} itens)")
        if vale:
            for cod, (_, p, v) in itens.items():
                if v > 0:
                    precos[cod] = min(p / v, precos.get(cod, p / v))
    return precos


if __name__ == "__main__":
    import sys
    dt = datetime.datetime.strptime(sys.argv[1], "%d/%m/%Y").date() if len(sys.argv) > 1 else None
    for arq in sorted(glob.glob(os.path.join(PASTA, "*.pdf"))):
        ini, fim, itens = ler_pdf(arq)
        print(os.path.basename(arq), ini, fim, len(itens))
        for c, (d, p, v) in itens.items():
            print(f"   {c:06d} {d[:40]:40s} {v:8.2f} {p:8.2f}")
    print(precos_campanha(dt))
