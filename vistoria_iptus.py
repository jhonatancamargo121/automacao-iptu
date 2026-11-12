# -*- coding: utf-8 -*-
"""
Vistoria dos IPTUs baixados.

Le a planilha, para cada linha marcada como "PDF baixado":
  - Verifica se o arquivo PDF correspondente existe no disco
  - Valida que e um PDF real (>0 bytes e comeca com %PDF)
  - Lista os faltantes
  - (opcional) Limpa o marcador para o baixar_iptu.py reprocessar

Uso:
  python vistoria_iptus.py
      -> Apenas relata, NAO modifica a planilha
  python vistoria_iptus.py --corrigir
      -> Tambem limpa o marcador 'PDF baixado' dos faltantes
      -> NA proxima rodada do baixar_iptu.py, eles serao re-baixados

  python vistoria_iptus.py --lista-faltantes
      -> Salva lista CSV com CCIs faltantes
"""
import argparse
import re
import sys
from pathlib import Path
from datetime import datetime

import openpyxl

# Mesmas constantes do baixar_iptu.py
NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]
PASTA_PDFS = "IPTUs_Baixados"

COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13


def achar_planilha():
    script_dir = Path(__file__).resolve().parent
    for base in (Path.cwd(), script_dir):
        for nome in NOMES_PLANILHA:
            p = base / nome
            if p.exists():
                return p
    for base in (Path.cwd(), script_dir):
        for f in base.glob("*.xlsx"):
            if "iptu" in f.name.lower():
                return f
    return None


def achar_pdf_por_cci(cci, pasta_pdfs):
    """Procura recursivamente um PDF cujo nome contenha 'CCI{n}_' ou 'CCI{n}.'"""
    base = Path(pasta_pdfs)
    if not base.exists():
        return None
    marcadores = (f"CCI{cci}_", f"CCI{cci}.")
    for pdf in base.rglob("*.pdf"):
        if any(m in pdf.name for m in marcadores):
            return pdf
    return None


def eh_pdf_valido(caminho):
    """Verifica se o arquivo existe, tem tamanho >0 e comeca com %PDF."""
    try:
        if not caminho or not caminho.exists():
            return False, "nao existe"
        tam = caminho.stat().st_size
        if tam == 0:
            return False, "vazio (0 bytes)"
        with open(caminho, "rb") as f:
            cabecalho = f.read(8)
        if not cabecalho.startswith(b"%PDF"):
            return False, f"nao e PDF (cabecalho={cabecalho[:6]!r})"
        return True, f"OK ({tam} bytes)"
    except Exception as e:
        return False, f"erro: {e}"


def main():
    parser = argparse.ArgumentParser(
        description="Vistoria dos IPTUs baixados (verifica integridade do disco vs planilha)"
    )
    parser.add_argument("--corrigir", action="store_true",
                        help="Limpa o marcador 'PDF baixado' dos faltantes "
                             "para que sejam reprocessados na proxima rodada.")
    parser.add_argument("--lista-faltantes", action="store_true",
                        help="Salva CSV com a lista de CCIs faltantes")
    args = parser.parse_args()

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    print(f"Lendo planilha: {planilha}")
    import time
    wb = None
    for tentativa in range(8):
        try:
            wb = openpyxl.load_workbook(planilha)
            break
        except Exception as e:
            print(f"  [.] tentativa {tentativa+1}/8 ({type(e).__name__}); aguardando 2s...")
            time.sleep(2)
    if wb is None:
        print("[ERRO] Nao consegui ler a planilha (baixar_iptu escrevendo?).")
        sys.exit(1)
    ws = wb.active

    total_marcado_baixado = 0
    total_ok              = 0
    total_faltantes       = 0
    total_corrompidos     = 0
    total_sem_debito      = 0
    total_indisponivel    = 0

    faltantes = []   # [(linha, cci, comprador, etapa, motivo)]
    corrompidos = []

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        status = str(row[COL_STATUS] or "").strip().lower()
        cci    = row[COL_CCI]
        obs    = str(row[COL_OBS_IPTU] or "").strip()
        if status != "vendido" or not cci:
            continue

        obs_low = obs.lower()
        # Contadores informativos
        if "sem_debito" in obs_low or "sem debito" in obs_low:
            total_sem_debito += 1
            continue
        if "indisponivel" in obs_low or "boleto_indisponivel" in obs_low:
            total_indisponivel += 1
            continue
        if "pdf baixado" not in obs_low:
            continue

        total_marcado_baixado += 1
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            cci_str = str(cci).strip()

        comprador = str(row[COL_COMPRADOR] or "").strip()
        etapa     = str(row[COL_ETAPA]     or "").strip()

        # Procura o PDF
        caminho = achar_pdf_por_cci(cci_str, PASTA_PDFS)
        valido, info = eh_pdf_valido(caminho)
        if valido:
            total_ok += 1
        else:
            if caminho is None or not caminho.exists():
                total_faltantes += 1
                faltantes.append((i, cci_str, comprador, etapa, info))
            else:
                total_corrompidos += 1
                corrompidos.append((i, cci_str, comprador, etapa, info))

    # Relatorio
    print("\n" + "=" * 60)
    print("VISTORIA DOS IPTUs BAIXADOS")
    print("=" * 60)
    print(f"  Marcados como 'PDF baixado' na planilha : {total_marcado_baixado}")
    print(f"  PDFs OK no disco                        : {total_ok}")
    print(f"  PDFs FALTANTES (nao acharam no disco)   : {total_faltantes}")
    print(f"  PDFs CORROMPIDOS (vazios/invalidos)     : {total_corrompidos}")
    print(f"")
    print(f"  Marcados como SEM_DEBITO (info)         : {total_sem_debito}")
    print(f"  Marcados como INDISPONIVEL (info)       : {total_indisponivel}")
    print("=" * 60)

    if faltantes:
        print(f"\n=== FALTANTES (primeiros 20 de {len(faltantes)}) ===")
        for linha, cci_str, comprador, etapa, info in faltantes[:20]:
            print(f"  linha {linha:>5} CCI {cci_str:>8} {comprador[:30]:<30} | {info}")

    if corrompidos:
        print(f"\n=== CORROMPIDOS ({len(corrompidos)}) ===")
        for linha, cci_str, comprador, etapa, info in corrompidos[:20]:
            print(f"  linha {linha:>5} CCI {cci_str:>8} {comprador[:30]:<30} | {info}")

    # Salva CSV se pedido
    if args.lista_faltantes and (faltantes or corrompidos):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = Path(f"faltantes_{ts}.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("linha;cci;comprador;etapa;motivo;tipo\n")
            for linha, cci_str, comprador, etapa, info in faltantes:
                f.write(f"{linha};{cci_str};{comprador};{etapa};{info};FALTANTE\n")
            for linha, cci_str, comprador, etapa, info in corrompidos:
                f.write(f"{linha};{cci_str};{comprador};{etapa};{info};CORROMPIDO\n")
        print(f"\n[OK] Lista salva em: {csv_path.resolve()}")

    # Corrige a planilha se pedido
    if args.corrigir and (faltantes or corrompidos):
        total_corrigir = len(faltantes) + len(corrompidos)
        resp = input(f"\nVai limpar o marcador 'PDF baixado' de {total_corrigir} "
                     f"linha(s) para reprocessar. Confirma? [s/N]: ").strip().lower()
        if resp != "s":
            print("Cancelado.")
            return

        for linha, *_ in faltantes + corrompidos:
            ws.cell(row=linha, column=COL_OBS_IPTU + 1).value = None

        try:
            wb.save(planilha)
            print(f"[OK] {total_corrigir} linha(s) reprocessaveis. "
                  f"Rode 'python baixar_iptu.py' pra re-baixar.")
        except PermissionError:
            print("[ERRO] Feche a planilha no Excel e tente novamente.")
            sys.exit(1)


if __name__ == "__main__":
    main()
