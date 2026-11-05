# Gera o PDF "marcas lideres x marca propria" (Arquivos e links da I.A.): quando o cliente pede uma marca LIDER do segmento,
# a I.A. oferece tambem a nossa marca. Base: Checklist de Marca Propria (coluna "Lideres de Mercado em SP e REDE").
# So entram itens que estao no catalogo (catalogo_top500.csv) ou na lista de medicamentos (FORCAR) - assim a I.A. sempre tem preco.
import csv, os, re, collections, unicodedata, openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from datetime import date
AQUI = os.path.dirname(os.path.abspath(__file__))
A = r"C:\Users\marcel.pereira\CLAUDE\AUTOMATIZ. ATENDIM.  WHATSAPP  LOJAS"
D = A + r"\ATUALIZAÇÃO DIARIA - CATALOGO PROD"
import glob as _g, re as _re
def checklist_mais_novo():
    """Checklist de Marca Propria mais novo (maior versao no nome: ..._RedeCentral_2026_NN.xlsx) na pasta ULTIMO."""
    pasta = r"C:\Users\marcel.pereira\Desktop\AUDITORIAS\01-CLAUDE\VERIFICAR NAS LOJAS\01- CHECKLIST - ULTIMO"
    a = [x for x in _g.glob(pasta + "\\Checklist_MarcaPropria_RedeCentral_2026_*.xlsx") if not _re.search(r"BACKUP|~\$", x)]
    return max(a, key=lambda x: int(_re.search(r"_(\d+)\.xlsx$", x).group(1))) if a else pasta + "\\Checklist_MarcaPropria_RedeCentral_2026_NAO_ENCONTRADO.xlsx"
CK = checklist_mais_novo()
MEDS = {"107026", "107025", "116240", "116239", "110417"}      # na lista de medicamentos (Gastroclean fora: principio ativo nao confirmado)
no_cat = {r['id'].lstrip('0') for r in csv.DictReader(open(os.path.join(AQUI, 'catalogo_top500.csv'), encoding='utf-8-sig'))}
# --- REVISAO DOS LIDERES (10/10/2026, pedido do Marcel) -------------------------------------------------------------
# A coluna de lideres do checklist as vezes vem de OUTRO segmento. Aqui: (1) tira marca propria e termo generico da lista de
# lideres (cliente nao "pede" Algodao ou Bem Care como lider); (2) corrige pelo codigo os itens cujo segmento esta errado.
n_ = lambda x: re.sub(r'[^a-z0-9 ]', '', unicodedata.normalize('NFKD', str(x).lower()).encode('ascii', 'ignore').decode()).strip()
PROPRIAS = {n_(x) for x in ["Bem Care", "Baby Care", "Gelavit", "Hidraday", "Melune", "Dermunn", "Illus", "Unipantol", "Bemvita", "Politabs", "Unifabra",
            "Polimag", "Farlife", "Puradent", "Keep", "Amis", "Vitaclinical", "Hair Heal", "Natusplant", "Laxivita", "Lufree", "Gelomag",
            "Hairheal", "Amis Tech", "Dermun", "Vita Clinical", "Bemfix", "Smalto", "Gelavit"]}
GENERICOS = {n_(x) for x in ["Balanca Digital", "Joelheira", "Munhequeira", "Tipoia", "Tornozeleira", "Algodao", "Soro Fisiologico", "Sol Fisiologica",
             "Agua Oxigenada", "Agua Boricada", "Agua Micelar", "Complexo B", "Vitamina B12", "Vitamina C", "Colageno Tipo Ii", "Creatina", "Cafeina",
             "Melatonina", "Magnesio", "Magnesio+Treonina", "Whey 100% Pure", "100% Whey", "Diclofen.Dietil.Mg", "Simeticona Mg", "Calcio Mdk",
             "Coletor Universal", "Smells", "Kerasys", "Pampers do oleo"]}
# codigo -> lideres corretos do segmento do item ('' = sem lider confiavel, nao oferecer por lider)
OVERRIDE = {
    # 04/11/2026: o checklist v29 trouxe os lideres corretos por segmento (formato "SP: (1)... · REDE: (1)..."); as correcoes manuais
    # de 10/10 foram aposentadas, EXCETO estas 2 em que a v29 ainda traz segmento errado (avisar o Marcel para corrigir no checklist):
    "115435": "Johnsons Baby · Joao E Maria",   # locao hidratante infantil vinha com Cerave/Dove/Nivea (adulto)
    "115434": "Johnsons Baby · Joao E Maria",   # oleo infantil vinha com Pampers (fralda) e a propria Baby Care
}
def limpa_lideres(lid):
    """Aceita o formato novo (v27+: 'SP: (1) A (2) B · REDE: (1) C') e o antigo ('(1) A (SP) · (2) B (REDE)'); junta SP e REDE sem repetir."""
    lid = re.sub(r'\[.*?\]', '', lid)   # observacao '[ajuste ...]' gravada no checklist
    lid = re.sub(r'\((SP|REDE|SP\s*/\s*REDE)\)', '', lid)
    nomes = []
    for x in re.findall(r'\(\d\)\s*([^()\u00b7]+)', lid):
        x = x.strip(' ·').strip()
        if x and x.lower() not in [n.lower() for n in nomes]:
            nomes.append(x)
    ok = [x for x in nomes if n_(x) not in PROPRIAS and n_(x) not in GENERICOS]
    return ' · '.join(ok), nomes
revisao = []   # (cod, categoria, subcategoria, item, lider original, lider aplicado, motivo)
grupos = collections.OrderedDict()
for w in openpyxl.load_workbook(CK, read_only=True, data_only=True).worksheets:
    if 'Instru' in w.title or 'Geral' in w.title: continue
    cat = re.sub(r'^\W+', '', w.title).strip(); sub = cat
    for r in w.iter_rows(values_only=True):
        if not r: continue
        a0 = str(r[0] or '')
        if a0.strip().startswith('■') and 'prod.' in str(r[5] or ''):
            sub = a0.replace('■', '').strip(); continue
        if len(r) > 6 and r[1] and str(r[1]).strip().isdigit():
            cod = str(r[1]).strip().lstrip('0'); lid = str(r[6] or '').strip()
            if 'natus' in str(r[3]).lower() or not lid: continue
            if cod not in no_cat and cod not in MEDS: continue
            orig = lid
            lid, brutos = limpa_lideres(lid)
            if cod in OVERRIDE:
                novo = OVERRIDE[cod]
                revisao.append((cod, cat, sub, str(r[3]).strip(), re.sub(r'\s+', ' ', orig), novo or '(sem lider - nao oferecer por lider)', 'segmento errado na planilha - corrigido'))
                lid = novo
            elif lid != ' · '.join(brutos):
                revisao.append((cod, cat, sub, str(r[3]).strip(), re.sub(r'\s+', ' ', orig), lid or '(sem lider)', 'tirada marca propria / termo generico'))
            if not lid: continue
            nome = re.sub(r'\s+', ' ', str(r[3]).replace('\xa0', ' ')).strip().title()
            grupos.setdefault((cat, sub, lid), []).append(nome + (' (preço na lista de medicamentos)' if cod in MEDS else ''))
st = ParagraphStyle('c', fontName='Helvetica', fontSize=8.5, leading=11)
sb = ParagraphStyle('b', fontName='Helvetica-Bold', fontSize=9, leading=12, spaceBefore=5)
stt = ParagraphStyle('t', fontName='Helvetica-Bold', fontSize=13, leading=17, spaceAfter=6)
S = [Paragraph('Rede Central Farma - Marcas líderes x nossa marca (indicação)', stt),
     Paragraph(f'Atualizada em {date.today():%d/%m/%Y}. Quando o cliente pedir (ou perguntar o preço de) uma das MARCAS LÍDERES de cada segmento abaixo, informe o preço do que ele pediu e OFEREÇA TAMBÉM, em uma frase curta, os produtos da nossa marca do mesmo segmento (use só os da mesma linha, mesma forma/tamanho quando houver). '
               'Preço sempre do catálogo (ou da lista de medicamentos, quando indicado). Nunca invente produto nem preço; se o item da nossa marca não estiver no catálogo/lista do dia, não ofereça. Nunca oriente dose, uso ou indicação.', st), Spacer(1, 6)]
atual = None
for (cat, sub, lid), nomes in grupos.items():
    if cat != atual: S.append(Paragraph(cat.upper(), sb)); atual = cat
    S.append(Paragraph(f'<b>{sub}</b> - se pedir: {lid}', st))
    S.append(Paragraph('&nbsp;&nbsp;&rarr; nossa marca: ' + '; '.join(nomes), st)); S.append(Spacer(1, 3))
saida = os.path.join(D, 'marcas_lideres_x_marca_propria.pdf')
SimpleDocTemplate(saida, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=1.5*cm, bottomMargin=1.5*cm, title='Marcas lideres x marca propria').build(S)
print(saida, os.path.getsize(saida)//1024, 'KB |', len(grupos), 'grupos |', sum(len(v) for v in grupos.values()), 'itens')

# Excel de conferencia dos lideres (para o Marcel revisar o que foi trocado)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
wb = Workbook(); ws = wb.active; ws.title = 'Lideres corrigidos'
ws.append(['Código', 'Categoria', 'Subcategoria', 'Item de marca própria', 'Líderes que estavam na planilha', 'Líderes agora usados', 'Motivo'])
for r in revisao: ws.append(list(r))
for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78')
for i, wd in enumerate([10, 14, 24, 52, 62, 52, 40], 1): ws.column_dimensions[chr(64 + i)].width = wd
ws2 = wb.create_sheet('Como ficou (por segmento)')
ws2.append(['Categoria', 'Subcategoria', 'Se o cliente pedir (líderes)', 'Oferecer (nossa marca)'])
for (cat, sub, lid), nomes in grupos.items(): ws2.append([cat, sub, lid, '; '.join(nomes)])
for c in ws2[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78')
for i, wd in enumerate([14, 28, 60, 110], 1): ws2.column_dimensions[chr(64 + i)].width = wd
x = os.path.join(A, 'revisao_lideres_marca_propria_10-10-2026.xlsx'); wb.save(x); print(x, len(revisao), 'itens ajustados')
