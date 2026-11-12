# -*- coding: utf-8 -*-
"""
REPROCESSAR FALHAS - Natela Analytics

Le falhas.log, identifica CCIs com erro real (incluindo BoletoIndisponivel)
e limpa os marcadores deles na planilha para que o baixar_iptu.py reprocesse
com uma DATA diferente (ex: 29/05/2026 em vez de 29/06/2026).

Util quando alguns boletos so estao disponiveis ate uma data limite.

Uso:
  python reprocessar_falhas.py
      -> mostra os CCIs falhados, pergunta se limpa marcadores
  python reprocessar_falhas.py --incluir-indisponivel
      -> tambem inclui os marcados como BoletoIndisponivel
      (recomendado quando a data muda)
"""
import argparse
import re
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl

NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]
COL_CCI       = 10
COL_OBS_IPTU  = 12
ARQ_LOG       = "falhas.log"

# Pastas para procurar PDFs ja existentes (de rodadas anteriores).
# Se um CCI ja tem PDF em qualquer uma delas, NAO reprocessamos.
PASTAS_CRUZAR = [
    "IPTUs_Baixados",                  # rodada original (29/05/2026)
    "IPTUs_Baixados_29_06_2026",       # rodada atual (29/06/2026)
    "IPTUs_Baixados_29_05_2026_RETRY", # eventual retry anterior
]

# Tipos de erro no log
PADRAO_LINHA        = re.compile(r"CCI\s*=\s*(\d+)\s*\t(.+)$")
PADRAO_SEM_DEBITO   = re.compile(r"SemDebito\s*:", re.I)
PADRAO_INDISPONIVEL = re.compile(r"BoletoIndisponivel\s*:", re.I)
# Tudo que NAO eh SemDebito/Indisponivel sao erros transientes


def achar_planilha():
    for nome in NOMES_PLANILHA:
        p = Path(nome)
        if p.exists():
            return p
    for f in Path(".").glob("*.xlsx"):
        if "iptu" in f.name.lower():
            return f
    return None


def mapear_ccis_ja_baixados(pastas):
    """Varre as pastas e devolve {cci: caminho_do_pdf} dos ja baixados."""
    mapa = {}
    for nome in pastas:
        p = Path(nome)
        if not p.exists():
            continue
        for pdf in p.rglob("*.pdf"):
            m = re.search(r"CCI(\d+)[_.]", pdf.name, re.I)
            if m:
                cci = m.group(1)
                # se ja tem, mantem o primeiro (preferencia da ordem da lista PASTAS_CRUZAR)
                if cci not in mapa:
                    mapa[cci] = pdf
    return mapa


def classificar_log(arq_log):
    """
    Retorna 3 sets:
      cci_sem_debito (esperado, nao reprocessar)
      cci_indisponivel (talvez funcione com outra data)
      cci_erros (erros transientes, sempre reprocessar)
    """
    sem_debito = set()
    indisponivel = set()
    erros = set()
    if not Path(arq_log).exists():
        return sem_debito, indisponivel, erros
    with open(arq_log, "r", encoding="utf-8", errors="ignore") as f:
        for linha in f:
            m = PADRAO_LINHA.search(linha)
            if not m:
                continue
            cci, motivo = m.group(1), m.group(2)
            if PADRAO_SEM_DEBITO.search(motivo):
                sem_debito.add(cci)
            elif PADRAO_INDISPONIVEL.search(motivo):
                indisponivel.add(cci)
            else:
                erros.add(cci)
    # Se foi sem_debito alguma vez, prevalece (nao reprocessa)
    indisponivel -= sem_debito
    erros -= sem_debito
    erros -= indisponivel   # erros so se NAO for ja indisponivel
    return sem_debito, indisponivel, erros


def main():
    parser = argparse.ArgumentParser(description="Reprocessa CCIs com erro")
    parser.add_argument("--incluir-indisponivel", action="store_true",
                        help="Tambem reprocessa CCIs marcados como BoletoIndisponivel "
                             "(use quando a data limite do boleto mudou)")
    args = parser.parse_args()

    sem_debito, indisponivel, erros = classificar_log(ARQ_LOG)
    print(f"Lendo {ARQ_LOG}:")
    print(f"  SemDebito         : {len(sem_debito)}  (nao reprocessar)")
    print(f"  BoletoIndisponivel: {len(indisponivel)}")
    print(f"  Outros erros      : {len(erros)}")

    a_reprocessar = set(erros)
    if args.incluir_indisponivel:
        a_reprocessar |= indisponivel
        print(f"\n--incluir-indisponivel: vou reprocessar TAMBEM os indisponiveis")

    if not a_reprocessar:
        print("\nNenhum CCI a reprocessar.")
        return

    print(f"\nTotal candidatos a reprocessar: {len(a_reprocessar)} CCIs")

    # CRUZAR com PDFs ja existentes em outras pastas (29/05 antiga, etc.)
    print(f"\nCruzando com pastas:")
    for p in PASTAS_CRUZAR:
        existe = "OK" if Path(p).exists() else "nao existe"
        print(f"  - {p:<40} [{existe}]")
    mapa_ja_baixados = mapear_ccis_ja_baixados(PASTAS_CRUZAR)
    print(f"  Total de CCIs ja com PDF em disco: {len(mapa_ja_baixados)}")

    ja_tem_pdf  = a_reprocessar & set(mapa_ja_baixados.keys())
    falta_baixar = a_reprocessar - ja_tem_pdf

    print(f"\nResultado do cruzamento:")
    print(f"  Ja tem PDF em outra pasta (pular)  : {len(ja_tem_pdf)}")
    print(f"  Realmente falta baixar             : {len(falta_baixar)}")

    if ja_tem_pdf:
        print(f"\nAmostra dos JA EXISTENTES (primeiros 10):")
        for cci in sorted(ja_tem_pdf)[:10]:
            print(f"  CCI {cci}  -> {mapa_ja_baixados[cci]}")

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    print(f"\nAbrindo planilha: {planilha}")
    wb = None
    for tentativa in range(8):
        try:
            wb = openpyxl.load_workbook(planilha)
            break
        except Exception:
            time.sleep(2)
    if wb is None:
        print("[ERRO] Nao consegui ler. FECHE o baixar_iptu.py se estiver rodando.")
        sys.exit(1)
    ws = wb.active

    # Mapeia linhas dos CCIs candidatos
    linha_por_cci = {}
    for row in ws.iter_rows(min_row=2, values_only=False):
        cci_val = row[COL_CCI].value
        if cci_val is None:
            continue
        try:
            cci_str = str(int(float(cci_val)))
        except (TypeError, ValueError):
            continue
        if cci_str in a_reprocessar:
            linha_por_cci[cci_str] = row[0].row

    # Lista 1: ja existentes -> marcar como PDF baixado (apontando pra pasta antiga)
    linhas_marcar_baixado = [
        (linha_por_cci[cci], cci, mapa_ja_baixados[cci])
        for cci in ja_tem_pdf if cci in linha_por_cci
    ]
    # Lista 2: realmente falta -> limpar marcador pra reprocessar
    linhas_limpar = [
        (linha_por_cci[cci], cci)
        for cci in falta_baixar if cci in linha_por_cci
    ]

    print(f"\nAcoes na planilha:")
    print(f"  Marcar como JA BAIXADO (apontando pra outra pasta) : {len(linhas_marcar_baixado)}")
    print(f"  Limpar marcador pra reprocessar                    : {len(linhas_limpar)}")

    if not linhas_marcar_baixado and not linhas_limpar:
        print("Nada a fazer.")
        return

    resp = input("\nConfirma essas acoes? [s/N]: ").strip().lower()
    if resp != "s":
        print("Cancelado.")
        return

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    for linha, cci, caminho in linhas_marcar_baixado:
        pasta_origem = caminho.parent.name
        ws.cell(row=linha, column=COL_OBS_IPTU + 1).value = (
            f"PDF baixado {agora} - {caminho.name} (de {pasta_origem})"
        )
    for linha, _ in linhas_limpar:
        ws.cell(row=linha, column=COL_OBS_IPTU + 1).value = None

    try:
        wb.save(planilha)
    except PermissionError:
        print("[ERRO] Feche a planilha no Excel e tente de novo.")
        sys.exit(1)

    print(f"\n[OK] Planilha atualizada.")
    print("\n" + "=" * 60)
    if linhas_limpar:
        print(f"Sobraram {len(linhas_limpar)} CCIs pra REBAIXAR com data 29/05/2026.")
        print()
        print('  python baixar_iptu.py --data 29/05/2026 \\')
        print('                        --pasta IPTUs_Baixados_29_05_2026_RETRY')
    else:
        print("TODOS os CCIs falhados ja tinham PDF em outra pasta.")
        print("Nao precisa rebaixar nada. So foram remarcados na planilha.")
    print("=" * 60)


if __name__ == "__main__":
    main()
