# -*- coding: utf-8 -*-
"""
RESTAURAR MARCADORES - Natela Analytics

Le o falhas.log gerado pelo baixar_iptu.py e restaura os marcadores
de SEM_DEBITO e BOLETO_INDISPONIVEL na planilha.

Util quando voce limpou o cache (limpar_cache.py --tudo) e quer voltar
a marcacao dos lotes ja analisados, sem ter que reprocessar todos.

Comportamento:
  - Le falhas.log linha a linha
  - Identifica CCIs marcados como SemDebito ou BoletoIndisponivel
  - Para cada CCI, marca a linha correspondente na planilha
  - PRESERVA marcadores existentes (ex: PDF baixado, Guia enviada)

Uso:
  python restaurar_marcadores.py
      -> mostra resumo + pede confirmacao
"""
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

import openpyxl

NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]
COL_CCI       = 10   # 0-indexed
COL_OBS_IPTU  = 12   # 0-indexed
ARQ_LOG       = "falhas.log"

# Padroes que identificam classificacao final do lote
PADRAO_SEM_DEBITO    = re.compile(r"SemDebito\s*:", re.I)
PADRAO_INDISPONIVEL  = re.compile(r"BoletoIndisponivel\s*:", re.I)
# Formato de linha do log: "TIMESTAMP\tCCI=N\tMOTIVO"
PADRAO_LINHA         = re.compile(r"CCI\s*=\s*(\d+)\s*\t(.+)$")


def achar_planilha():
    for nome in NOMES_PLANILHA:
        p = Path(nome)
        if p.exists():
            return p
    for f in Path(".").glob("*.xlsx"):
        if "iptu" in f.name.lower():
            return f
    return None


def parsear_log(arq_log):
    """Le falhas.log e devolve dois sets: ccis_sem_debito, ccis_indisponivel."""
    sem_debito = set()
    indisponivel = set()
    total_linhas = 0
    if not Path(arq_log).exists():
        return sem_debito, indisponivel, total_linhas
    with open(arq_log, "r", encoding="utf-8", errors="ignore") as f:
        for linha in f:
            total_linhas += 1
            m = PADRAO_LINHA.search(linha)
            if not m:
                continue
            cci = m.group(1)
            motivo = m.group(2)
            if PADRAO_SEM_DEBITO.search(motivo):
                sem_debito.add(cci)
            elif PADRAO_INDISPONIVEL.search(motivo):
                indisponivel.add(cci)
    # Se um CCI aparece em ambos (foi indisponivel uma vez e depois sem debito),
    # prevalece o mais recente. Pra simplicidade, mantem em ambos e quem ler
    # decide. Mas vou remover de indisponivel se tambem aparece em sem_debito.
    indisponivel -= sem_debito
    return sem_debito, indisponivel, total_linhas


def main():
    arq_log = Path(ARQ_LOG)
    print(f"Lendo log: {arq_log.resolve()}")
    if not arq_log.exists():
        print(f"[ERRO] {ARQ_LOG} nao encontrado.")
        sys.exit(1)

    sem_debito, indisponivel, total_linhas = parsear_log(arq_log)
    print(f"  {total_linhas} linhas no log")
    print(f"  CCIs marcados como SEM_DEBITO   : {len(sem_debito)}")
    print(f"  CCIs marcados como INDISPONIVEL : {len(indisponivel)}")

    if not sem_debito and not indisponivel:
        print("Nada para restaurar.")
        return

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    print(f"\nPlanilha: {planilha}")
    print("Tentando abrir (retry caso baixar_iptu esteja escrevendo)...")
    wb = None
    for tentativa in range(8):
        try:
            wb = openpyxl.load_workbook(planilha)
            break
        except Exception as e:
            print(f"  tentativa {tentativa+1}/8 falhou ({type(e).__name__}); aguardando 2s...")
            time.sleep(2)
    if wb is None:
        print("[ERRO] Nao consegui ler a planilha. FECHE o baixar_iptu.py se estiver rodando.")
        sys.exit(1)
    ws = wb.active

    # Indexa as linhas por CCI
    linhas_por_cci = {}
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=False), 2):
        cci_val = row[COL_CCI].value
        if cci_val is None:
            continue
        try:
            cci_str = str(int(float(cci_val)))
        except (TypeError, ValueError):
            continue
        linhas_por_cci[cci_str] = i

    # Verifica conflitos: linhas ja marcadas
    a_marcar_sd  = []   # (cci, linha, valor_atual)
    a_marcar_ind = []
    conflitos    = 0

    for cci in sem_debito:
        linha = linhas_por_cci.get(cci)
        if linha is None:
            continue
        cell = ws.cell(row=linha, column=COL_OBS_IPTU + 1)
        val_atual = (str(cell.value) if cell.value else "").strip().lower()
        if "sem_debito" in val_atual:
            continue  # ja marcado, ok
        if "pdf baixado" in val_atual:
            conflitos += 1
            continue  # tem PDF baixado, melhor nao sobrescrever
        a_marcar_sd.append((cci, linha))

    for cci in indisponivel:
        linha = linhas_por_cci.get(cci)
        if linha is None:
            continue
        cell = ws.cell(row=linha, column=COL_OBS_IPTU + 1)
        val_atual = (str(cell.value) if cell.value else "").strip().lower()
        if "indisponivel" in val_atual:
            continue
        if "pdf baixado" in val_atual or "sem_debito" in val_atual:
            conflitos += 1
            continue
        a_marcar_ind.append((cci, linha))

    print(f"\nResumo:")
    print(f"  Vou MARCAR SEM_DEBITO em      : {len(a_marcar_sd)} linhas")
    print(f"  Vou MARCAR INDISPONIVEL em    : {len(a_marcar_ind)} linhas")
    if conflitos:
        print(f"  Conflitos (preservados como estao): {conflitos}")

    if not a_marcar_sd and not a_marcar_ind:
        print("Tudo ja esta marcado. Nada a fazer.")
        return

    resp = input("\nConfirma a restauracao? [s/N]: ").strip().lower()
    if resp != "s":
        print("Cancelado.")
        return

    from datetime import datetime
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    for cci, linha in a_marcar_sd:
        cell = ws.cell(row=linha, column=COL_OBS_IPTU + 1)
        cell.value = f"PDF baixado {agora} - SEM_DEBITO_CCI{cci}"

    for cci, linha in a_marcar_ind:
        cell = ws.cell(row=linha, column=COL_OBS_IPTU + 1)
        cell.value = f"PDF baixado {agora} - BOLETO_INDISPONIVEL_CCI{cci}"

    try:
        wb.save(planilha)
    except PermissionError:
        print("[ERRO] Feche a planilha no Excel e tente novamente.")
        sys.exit(1)

    total_marcados = len(a_marcar_sd) + len(a_marcar_ind)
    print(f"\n[OK] {total_marcados} marcadores restaurados.")
    print(f"     SEM_DEBITO    : {len(a_marcar_sd)}")
    print(f"     INDISPONIVEL  : {len(a_marcar_ind)}")
    print(f"\nProxima rodada do baixar_iptu.py vai PULAR esses lotes automaticamente.")


if __name__ == "__main__":
    main()
