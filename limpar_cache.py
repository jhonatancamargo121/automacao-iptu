# -*- coding: utf-8 -*-
"""
Limpa marcadores das tentativas anteriores do baixar_iptu.py na planilha.

Comportamento padrao (RECOMENDADO):
  - Limpa SO 'PDF baixado ...'
  - PRESERVA 'SEM_DEBITO' e 'INDISPONIVEL' (esses lotes ja foram analisados,
    nao precisa re-verificar - economiza muito tempo)

Modo --tudo:
  - Limpa TODOS os marcadores ('PDF baixado', 'SEM_DEBITO', 'INDISPONIVEL')
  - Use isso so se quiser reanalisar totalmente

Uso:
  python limpar_cache.py         -> limpa so PDF baixado (preserva SEM_DEBITO)
  python limpar_cache.py --tudo  -> limpa tudo (reset completo de markers)
"""
import argparse
from pathlib import Path
import openpyxl

PLANILHA = "IPTU - CIDADE EXEMPLO (2).xlsx"
COL_OBS_IPTU = 12   # 0-indexed (coluna M)

MARCADORES_PDF = ("pdf baixado",)
MARCADORES_OUTROS = ("sem_debito", "indisponivel", "boleto_indisponivel")


def main():
    parser = argparse.ArgumentParser(description="Limpa marcadores do baixar_iptu na planilha")
    parser.add_argument("--tudo", action="store_true",
                        help="Limpa TODOS os marcadores (incluindo SEM_DEBITO/INDISPONIVEL)")
    args = parser.parse_args()

    if args.tudo:
        marcadores = MARCADORES_PDF + MARCADORES_OUTROS
        print("MODO: --tudo (limpa TODOS marcadores - vai reanalisar tudo)")
    else:
        marcadores = MARCADORES_PDF
        print("MODO: padrao (limpa SO 'PDF baixado', preserva SEM_DEBITO/INDISPONIVEL)")
        print("       (use --tudo para limpar tambem os ja-analisados)")

    p = Path(PLANILHA)
    if not p.exists():
        p = Path("IPTU - CIDADE EXEMPLO (2).xlsx")
    if not p.exists():
        print("[ERRO] planilha nao encontrada.")
        return

    wb = openpyxl.load_workbook(p)
    ws = wb.active
    limpas = 0
    preservadas_outras = 0  # SEM_DEBITO/INDISPONIVEL preservados (so no modo padrao)
    preservadas_demais = 0  # outras (ex: "Guia enviada e-mail cliente")

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
        cell = row[COL_OBS_IPTU]
        val = (str(cell.value) if cell.value else "").strip().lower()
        if not val:
            continue
        # Verifica se contem algum marcador alvo
        eh_pdf       = any(m in val for m in MARCADORES_PDF)
        eh_outro     = any(m in val for m in MARCADORES_OUTROS)

        if args.tudo:
            if eh_pdf or eh_outro:
                cell.value = None
                limpas += 1
            else:
                preservadas_demais += 1
        else:
            # Modo padrao: limpa so PDF baixado, preserva sem_debito/indisponivel
            if eh_pdf and not eh_outro:
                cell.value = None
                limpas += 1
            elif eh_outro:
                preservadas_outras += 1
            else:
                preservadas_demais += 1

    if limpas == 0:
        print("Nada para limpar (nenhum marcador alvo encontrado).")
        return

    print(f"\nResumo:")
    print(f"  Vou LIMPAR              : {limpas} linhas")
    if not args.tudo:
        print(f"  PRESERVAR SEM_DEBITO    : {preservadas_outras} linhas (ja analisados)")
    print(f"  PRESERVAR demais        : {preservadas_demais} linhas (ex: Guia enviada)")

    resp = input("\nConfirma? [s/N]: ").strip().lower()
    if resp != "s":
        print("Cancelado.")
        return

    wb.save(p)
    print(f"[OK] {limpas} celulas limpas em '{p}'.")


if __name__ == "__main__":
    main()
