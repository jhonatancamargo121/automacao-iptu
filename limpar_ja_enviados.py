# -*- coding: utf-8 -*-
"""
LIMPAR JA ENVIADOS - Natela Analytics

Le a planilha, identifica os CCIs ja marcados como
'Guia enviada e-mail cliente' e MOVE os PDFs correspondentes do
IPTUs_Final_Consolidado/ pra IPTUs_Ja_Enviados_Arquivo/.

Assim a pasta consolidada fica so com os IPTUs que ainda PRECISAM ser enviados.
Os movidos ficam preservados na pasta de arquivo (nao apaga).

Uso:
  python limpar_ja_enviados.py
      -> mostra resumo e pede confirmacao
  python limpar_ja_enviados.py --executar
      -> sem perguntar
  python limpar_ja_enviados.py --apagar
      -> apaga em vez de mover (CUIDADO: irreversivel)
  python limpar_ja_enviados.py --origem IPTUs_Final --executar
      -> usa outra pasta de origem
"""
import argparse
import re
import shutil
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
COL_STATUS    = 0
COL_CCI       = 10
COL_OBS_IPTU  = 12

PASTA_ORIGEM_PADRAO   = "IPTUs_Final_Consolidado"
PASTA_ARQUIVO_PADRAO  = "IPTUs_Ja_Enviados_Arquivo"

MARCADOR_ENVIADO = "guia enviada e-mail cliente"


def achar_planilha():
    for nome in NOMES_PLANILHA:
        p = Path(nome)
        if p.exists():
            return p
    for f in Path(".").glob("*.xlsx"):
        if "iptu" in f.name.lower():
            return f
    return None


def carregar_ccis_ja_enviados(planilha_path):
    """Retorna set de CCIs marcados como 'Guia enviada e-mail cliente'."""
    print(f"Lendo planilha: {planilha_path}")
    wb = None
    for _ in range(8):
        try:
            wb = openpyxl.load_workbook(planilha_path, read_only=True)
            break
        except Exception:
            time.sleep(2)
    if wb is None:
        print("[ERRO] Planilha sendo escrita por outro processo.")
        sys.exit(1)
    ws = wb.active

    enviados = set()
    detalhes = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_OBS_IPTU:
            continue
        cci = row[COL_CCI]
        if not cci:
            continue
        obs = str(row[COL_OBS_IPTU] or "").strip().lower()
        if MARCADOR_ENVIADO not in obs:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue
        enviados.add(cci_str)
        detalhes.append((i, cci_str, str(row[COL_OBS_IPTU])[:80]))
    wb.close()
    return enviados, detalhes


def main():
    parser = argparse.ArgumentParser(
        description="Remove (move) da pasta consolidada os IPTUs ja enviados"
    )
    parser.add_argument("--executar", action="store_true",
                        help="Pula a confirmacao")
    parser.add_argument("--apagar", action="store_true",
                        help="Apaga em vez de mover (IRREVERSIVEL)")
    parser.add_argument("--origem", type=str, default=PASTA_ORIGEM_PADRAO,
                        help=f"Pasta origem (padrao: {PASTA_ORIGEM_PADRAO})")
    parser.add_argument("--arquivo", type=str, default=PASTA_ARQUIVO_PADRAO,
                        help=f"Pasta de arquivo (padrao: {PASTA_ARQUIVO_PADRAO})")
    args = parser.parse_args()

    print("=" * 70)
    print("  Natela Analytics - Limpar PDFs Ja Enviados")
    print("=" * 70)

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    enviados, detalhes = carregar_ccis_ja_enviados(planilha)
    print(f"\nCCIs marcados como 'Guia enviada e-mail cliente': {len(enviados)}")
    if detalhes:
        print("\nAmostra (primeiros 10):")
        for linha, cci, obs in detalhes[:10]:
            print(f"  linha {linha:>5}  CCI {cci:>8}  obs='{obs}'")

    if not enviados:
        print("\nNada marcado como enviado. Nao ha o que limpar.")
        return

    pasta_origem = Path(args.origem)
    if not pasta_origem.exists():
        print(f"\n[ERRO] Pasta origem nao existe: {pasta_origem.resolve()}")
        sys.exit(1)

    # Inventario dos PDFs na pasta consolidada
    pdfs_na_pasta = []
    for pdf in pasta_origem.rglob("*.pdf"):
        m = re.search(r"CCI(\d+)", pdf.name, re.I)
        if m:
            cci = m.group(1)
            pdfs_na_pasta.append((cci, pdf))

    # Casa CCIs enviados com PDFs no disco
    a_remover = [(cci, pdf) for cci, pdf in pdfs_na_pasta if cci in enviados]
    nao_encontrados = enviados - {cci for cci, _ in pdfs_na_pasta}

    print(f"\nPasta origem        : {pasta_origem.resolve()}")
    print(f"PDFs na pasta       : {len(pdfs_na_pasta)}")
    print(f"PDFs a remover      : {len(a_remover)}  (casaram com CCIs enviados)")
    if nao_encontrados:
        print(f"CCIs sem PDF aqui   : {len(nao_encontrados)}  (ja removidos ou em outra pasta)")

    if not a_remover:
        print("\nNada a remover.")
        return

    acao = "APAGAR" if args.apagar else "MOVER PARA ARQUIVO"
    pasta_arquivo = Path(args.arquivo)
    print(f"\n>>> ACAO: {acao}")
    if not args.apagar:
        print(f"    Destino: {pasta_arquivo.resolve()}")
    print(f"    Quantidade: {len(a_remover)} PDFs")

    if not args.executar:
        resp = input(f"\nConfirma {acao}? [s/N]: ").strip().lower()
        if resp != "s":
            print("Cancelado.")
            return

    if not args.apagar:
        pasta_arquivo.mkdir(parents=True, exist_ok=True)

    sucesso = 0
    erros = 0
    log = []
    for cci, pdf in a_remover:
        try:
            rel = pdf.relative_to(pasta_origem)
            if args.apagar:
                pdf.unlink()
                log.append(("APAGADO", cci, str(rel)))
            else:
                destino = pasta_arquivo / rel
                destino.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(pdf), str(destino))
                log.append(("MOVIDO", cci, str(rel)))
            sucesso += 1
        except Exception as e:
            erros += 1
            log.append(("ERRO", cci, f"{pdf}: {e}"))

    # Limpa pastas vazias resultantes
    pastas_removidas = 0
    for d in sorted(pasta_origem.rglob("*"), key=lambda p: -len(p.parts)):
        if d.is_dir():
            try:
                d.rmdir()  # so remove se estiver vazia
                pastas_removidas += 1
            except OSError:
                pass

    # Salva log
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path(f"limpeza_enviados_{ts}.log")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"# Limpeza {ts}\n")
        f.write(f"# Pasta origem: {pasta_origem.resolve()}\n")
        if not args.apagar:
            f.write(f"# Pasta arquivo: {pasta_arquivo.resolve()}\n")
        for entry in log:
            f.write("\t".join(str(x) for x in entry) + "\n")

    print(f"\n[OK] {sucesso} PDFs processados.")
    if erros:
        print(f"[!!] {erros} erros (ver {log_path})")
    if pastas_removidas:
        print(f"     {pastas_removidas} pasta(s) vazia(s) removida(s)")
    print(f"     Log: {log_path.resolve()}")

    # Resumo final
    restantes = sum(1 for _ in pasta_origem.rglob("*.pdf"))
    print(f"\nPDFs RESTANTES em '{pasta_origem.name}/' (pendentes de envio): {restantes}")
    if not args.apagar:
        arquivados = sum(1 for _ in pasta_arquivo.rglob("*.pdf"))
        print(f"PDFs ARQUIVADOS em '{pasta_arquivo.name}/' (ja enviados):    {arquivados}")


if __name__ == "__main__":
    main()
