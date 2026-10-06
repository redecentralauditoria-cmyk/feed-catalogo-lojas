# Gera o Excel (com filtro) dos 500 do catalogo do WhatsApp: ranking por faturamento = Quantidade x Valor Bruto (fechamento mais recente).
import csv, os, glob, unicodedata, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
A = r"C:\Users\marcel.pereira\CLAUDE\AUTOMATIZ. ATENDIM.  WHATSAPP  LOJAS"
D = A + r"\ATUALIZAÇÃO DIARIA - CATALOGO PROD"
n = lambda s: unicodedata.normalize('NFKD', s.lower()).encode('ascii', 'ignore').decode()
f = max(glob.glob(D + r"\CADASTRO*.xlsx"), key=os.path.getmtime)
g = {str(r[1]).strip().lstrip('0'): (r[4] or '').strip()
     for r in openpyxl.load_workbook(f, read_only=True, data_only=True)["CADASTRO"].iter_rows(min_row=2, values_only=True) if r[1]}
arqs = glob.glob(A + r"\VENDAS GERAL\20??-?? Vendas geral*.xlsx")
vf = max(arqs, key=lambda a: (os.path.basename(a)[:7], "editada" in n(os.path.basename(a))))
it = openpyxl.load_workbook(vf, read_only=True, data_only=True)["Plan1"].iter_rows(values_only=True)
cab = [str(c or '').lower() for c in next(it)]
iv = [i for i, c in enumerate(cab) if 'vendedor' in c and 'nome' in c][0]
ig, ip, iq, ib = cab.index('filial'), cab.index('produto'), cab.index('quantidade'), cab.index('valor bruto')
fat, qt = {}, {}
for r in it:
    if not r[ip] or not r[ig] or str(r[ig]).strip().lstrip('0') == '': continue
    if n(str(r[iv] or '')) in ('devolucao mercadorias', 'remanejos entre lojas'): continue
    k = str(r[ip]).strip().lstrip('0'); q = float(r[iq] or 0)
    fat[k] = fat.get(k, 0) + q * float(r[ib] or 0); qt[k] = qt.get(k, 0) + q
top = list(csv.DictReader(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'catalogo_top500.csv'), encoding='utf-8-sig')))
rows = sorted(([r['id'], r['title'], r['brand'], g.get(r['id'].lstrip('0'), ''), float(r['price'].split()[0]),
                qt.get(r['id'].lstrip('0'), 0), round(fat.get(r['id'].lstrip('0'), 0), 2)] for r in top), key=lambda x: -x[6])
wb = openpyxl.Workbook(); s = wb.active; s.title = 'Top 500'
s.append(['Posição', 'Código', 'Produto', 'Marca', 'Grupo FarmaPRO', 'Preço (P.M.C.)', 'Qtde vendida (mês)', 'Faturamento (Qtde x Valor Bruto)'])
for i, r in enumerate(rows, 1): s.append([i] + r)
for c in s[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78'); c.alignment = Alignment(wrap_text=True, vertical='center')
for i, wd in enumerate([9, 10, 58, 26, 36, 14, 14, 22], 1): s.column_dimensions[get_column_letter(i)].width = wd
s.row_dimensions[1].height = 45
for r in s.iter_rows(min_row=2): r[5].number_format = 'R$ #,##0.00'; r[7].number_format = 'R$ #,##0.00'; r[6].number_format = '#,##0'
s.freeze_panes = 'C2'; s.auto_filter.ref = s.dimensions
out = A + r"\catalogo_whatsapp_500_produtos_06-10-2026.xlsx"; wb.save(out); print(out, len(rows), os.path.basename(vf))
