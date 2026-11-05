# -*- coding: utf-8 -*-
"""Gera (e opcionalmente envia ao Drive) as listas de preço como se HOJE fosse outra data. Uso: python gerar_para_data.py 06/11/2026 [--enviar]
Criado em 05/11/2026 para deixar o DIA D pronto na véspera: usa o cadastro/estoque mais recentes, mas as campanhas válidas na data pedida."""
import datetime, runpy, sys, os
alvo = datetime.datetime.strptime(sys.argv[1], "%d/%m/%Y").date()
class _D(datetime.date):
    @classmethod
    def today(cls):
        return cls(alvo.year, alvo.month, alvo.day)
datetime.date = _D
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.getcwd())
etapas = ["gerar_lista_medicamentos.py", "gerar_lista_genericos.py", "gerar_lista_demais.py", "gerar_lista_controlados.py", "gerar_tabela_principio_ativo.py"]
if "--enviar" in sys.argv:
    etapas.append("enviar_listas_drive.py")
for e in etapas:
    print("==", e, flush=True)
    runpy.run_path(e, run_name="__main__")
