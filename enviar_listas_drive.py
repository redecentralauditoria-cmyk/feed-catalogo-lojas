"""
Envia as 5 listas de preço do dia (+ marcas líderes) para a pasta do Google Drive IA_WhatsApp_Precos,
com NOME FIXO (sem data), sobrescrevendo o mesmo arquivo (mesmo ID no Drive) — o agente de IA das lojas
conectado à pasta sincroniza sozinho (a Meta sincroniza a cada ~12h). Usa o rclone (remote 'gdrive_ia',
conta redecentralauditoria@gmail.com, autorizado em 21/10/2026). Rodar DEPOIS dos 5 geradores de lista.
"""
import glob, os, subprocess, sys
from datetime import date

RCLONE = glob.glob(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WinGet\Packages\Rclone.Rclone*\rclone-*\rclone.exe"))
DIARIO = r"C:\Users\marcel.pereira\CLAUDE\AUTOMATIZ. ATENDIM.  WHATSAPP  LOJAS\ATUALIZAÇÃO DIARIA - CATALOGO PROD"
DESTINO = "gdrive_ia:IA_WhatsApp_Precos"
hoje = f"{date.today():%d-%m-%Y}"
ARQS = {f"precos_medicamentos_venda_livre_{hoje}.pdf": "precos_medicamentos_venda_livre.pdf",
        f"precos_genericos_{hoje}.pdf": "precos_genericos.pdf",
        f"precos_demais_medicamentos_{hoje}.pdf": "precos_demais_medicamentos.pdf",
        f"precos_demais_produtos_{hoje}.pdf": "precos_demais_produtos.pdf",
        f"precos_controlados_antibioticos_{hoje}.pdf": "precos_controlados_antibioticos.pdf",
        "marcas_lideres_x_marca_propria.pdf": "marcas_lideres_x_marca_propria.pdf"}
if not RCLONE: sys.exit("ERRO: rclone não encontrado")
falta = [a for a in ARQS if not os.path.exists(os.path.join(DIARIO, a))]
if falta: sys.exit("ERRO: faltam os arquivos de hoje: " + ", ".join(falta))
for local, remoto in ARQS.items():
    r = subprocess.run([RCLONE[0], "copyto", os.path.join(DIARIO, local), f"{DESTINO}/{remoto}"], capture_output=True, text=True)
    erro = "\n".join(l for l in r.stderr.splitlines() if "NOTICE" not in l)
    print(("OK   " if r.returncode == 0 else "ERRO ") + remoto + ("" if r.returncode == 0 else " -> " + erro))
