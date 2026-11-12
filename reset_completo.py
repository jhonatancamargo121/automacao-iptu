# -*- coding: utf-8 -*-
"""
RESET COMPLETO - Natela Analytics

Prepara para uma nova rodada de download com data diferente:
  1. Renomeia a pasta IPTUs_Baixados/ para IPTUs_Baixados_BACKUP_{data}/
     (preserva os PDFs antigos como historico)
  2. Limpa os marcadores 'PDF baixado', 'SEM_DEBITO', 'INDISPONIVEL' da planilha
  3. Apos confirmar, todos os lotes voltam a ser pendentes

Uso:
  python reset_completo.py
      -> Pede confirmacao antes de qualquer coisa.
"""
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]
PASTA_PDFS = "IPTUs_Baixados"
COL_OBS_IPTU = 12

MARCADORES = (
    "pdf baixado",
    "sem_debito",
    "indisponivel",
    "boleto_indisponivel",
)


def achar_planilha():
    for nome in NOMES_PLANILHA:
        p = Path(nome)
        if p.exists():
            return p
    for f in Path(".").glob("*.xlsx"):
        if "iptu" in f.name.lower():
            return f
    return None


def main():
    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    pasta_pdfs = Path(PASTA_PDFS)
    pdfs_existem = pasta_pdfs.exists()

    # Conta PDFs e marcadores
    pdfs_count = 0
    if pdfs_existem:
        pdfs_count = sum(1 for _ in pasta_pdfs.rglob("*.pdf"))

    import time
    wb = None
    for _ in range(5):
        try:
            wb = openpyxl.load_workbook(planilha)
            break
        except Exception:
            time.sleep(2)
    if wb is None:
        print("[ERRO] Planilha esta sendo escrita pelo baixar_iptu. "
              "FECHE-O antes de fazer reset.")
        sys.exit(1)
    ws = wb.active

    marcadas = 0
    for row in ws.iter_rows(min_row=2):
        cell = row[COL_OBS_IPTU]
        val = (str(cell.value) if cell.value else "").strip().lower()
        if any(m in val for m in MARCADORES):
            marcadas += 1

    print("=" * 60)
    print("  RESET COMPLETO - Natela Analytics")
    print("=" * 60)
    print(f"\nPlanilha: {planilha}")
    print(f"  Linhas marcadas (PDF baixado/SEM_DEBITO/etc): {marcadas}")
    print(f"\nPasta de PDFs: {pasta_pdfs.resolve()}")
    print(f"  PDFs encontrados: {pdfs_count}")
    print("\nO que vou fazer:")
    if pdfs_existem and pdfs_count > 0:
        print(f"  1. Renomear '{PASTA_PDFS}/' -> 'IPTUs_Baixados_BACKUP_{{data}}/'")
    print(f"  2. Limpar {marcadas} marcadores da coluna de obs na planilha")
    print(f"  3. Apos isso, RODAR 'python baixar_iptu.py' vai reprocessar TUDO")
    print("     (com a nova data de pagamento)")
    print()
    resp = input("CONFIRMA? Digite 'RESET' para prosseguir: ").strip()
    if resp.upper() != "RESET":
        print("Cancelado.")
        return

    # 1) Move PDFs antigos pra backup
    if pdfs_existem and pdfs_count > 0:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = Path(f"IPTUs_Baixados_BACKUP_{ts}")
        try:
            shutil.move(str(pasta_pdfs), str(backup))
            print(f"[OK] PDFs movidos para: {backup.resolve()}")
        except Exception as e:
            print(f"[ERRO] Falha ao mover pasta: {e}")
            sys.exit(1)

    # 2) Limpa marcadores na planilha
    for row in ws.iter_rows(min_row=2):
        cell = row[COL_OBS_IPTU]
        val = (str(cell.value) if cell.value else "").strip().lower()
        if any(m in val for m in MARCADORES):
            cell.value = None

    try:
        wb.save(planilha)
        print(f"[OK] {marcadas} marcadores limpos. Planilha salva.")
    except PermissionError:
        print("[ERRO] Feche a planilha no Excel e tente de novo.")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  RESET COMPLETO. Proximo passo:")
    print("    python baixar_iptu.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
