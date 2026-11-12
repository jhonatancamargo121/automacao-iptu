# -*- coding: utf-8 -*-
"""
VERIFICADOR DE PASTAS LOCAIS x PLANILHA - Natela Analytics

Para cada comprador que deveria ter PDF (marcado como 'PDF baixado' ou
'Guia enviada' na planilha), verifica se a pasta ETAPA{N}/COMPRADOR/
existe localmente E se ela tem o PDF dentro.

Mostra:
  - Compradores com pasta + PDF OK
  - Compradores com pasta mas SEM PDF
  - Compradores SEM PASTA (a planilha diz que baixou mas a pasta nao existe)
  - Compradores em multiplas etapas - mostra cada etapa
  - Diferencas de normalizacao do nome (acentos, espacos)

Saidas:
  - Terminal com resumo
  - CSV detalhado de tudo que esta inconsistente

Uso:
  python verificar_pastas_local.py
      -> verifica IPTUs_Final_Consolidado por padrao
  python verificar_pastas_local.py --pasta IPTUs_Baixados
      -> verifica outra pasta especifica
  python verificar_pastas_local.py --todas
      -> agrega de todas pastas IPTUs_* da pasta atual
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
PASTA_PADRAO = "IPTUs_Final_Consolidado"
COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13


def achar_planilha():
    for nome in NOMES_PLANILHA:
        p = Path(nome)
        if p.exists():
            return p
    for f in Path(".").glob("*.xlsx"):
        if "iptu" in f.name.lower():
            return f
    return None


def limpar_nome(nome):
    """Mesma normalizacao usada pelo baixar_iptu (preserva acentos)."""
    nome = (nome or "").upper().strip()
    nome = re.sub(r'[<>:"/\\|?*]', "", nome)
    nome = re.sub(r"\s+", "_", nome)
    return nome[:80] or "SEM_NOME"


def normalizar_sem_acento(s):
    """Pra comparar nomes sem que acentos atrapalhem."""
    import unicodedata
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper()


def extrair_etapa(s):
    m = re.search(r"(\d+)", str(s or ""))
    return m.group(1) if m else ""


def main():
    parser = argparse.ArgumentParser(
        description="Verificador de consistencia planilha x pastas locais"
    )
    parser.add_argument("--pasta", type=str, default=PASTA_PADRAO,
                        help=f"Pasta a verificar (padrao: {PASTA_PADRAO})")
    parser.add_argument("--todas", action="store_true",
                        help="Verifica TODAS pastas IPTUs_*/ da pasta atual")
    args = parser.parse_args()

    print("=" * 75)
    print("  Natela Analytics - Verificador de Pastas Locais")
    print("=" * 75)

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    # Pastas a verificar
    if args.todas:
        pastas = sorted([d for d in Path(".").iterdir()
                         if d.is_dir() and d.name.lower().startswith("iptus_")])
        print(f"\nVerificando TODAS as pastas IPTUs_*/:")
    else:
        p = Path(args.pasta)
        if not p.exists():
            print(f"[ERRO] Pasta nao existe: {p.resolve()}")
            sys.exit(1)
        pastas = [p]
        print(f"\nVerificando: {p.resolve()}")
    for d in pastas:
        cnt = sum(1 for _ in d.rglob("*.pdf")) if d.exists() else 0
        print(f"  - {d.name:<40} ({cnt} PDFs)")

    # Mapeia tudo que existe no disco: {etapa: {nome_pasta_norm: caminho}}
    disco = defaultdict(dict)
    # Tambem mapeia sem-acento -> nome real
    disco_sem_acento = defaultdict(dict)
    for base in pastas:
        for d in base.rglob("*"):
            if not d.is_dir():
                continue
            rel = d.relative_to(base).parts
            if len(rel) != 2:
                continue
            etapa, comprador = rel
            if not etapa.upper().startswith("ETAPA"):
                continue
            disco[etapa][comprador] = d
            disco_sem_acento[etapa][normalizar_sem_acento(comprador)] = comprador

    # Le planilha
    print(f"\nLendo planilha: {planilha}")
    wb = None
    for _ in range(8):
        try:
            wb = openpyxl.load_workbook(planilha, read_only=True)
            break
        except Exception:
            time.sleep(2)
    if wb is None:
        print("[ERRO] Planilha sendo escrita por outro processo.")
        sys.exit(1)
    ws = wb.active

    # Casa cada lote que deveria ter PDF com sua pasta no disco
    ok = []                # comprador OK, pasta + PDF presentes
    sem_pdf_na_pasta = []  # pasta existe mas vazia
    sem_pasta = []         # pasta inexistente
    casamento_aproximado = []  # achou com nome ligeiramente diferente

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        obs_low = str(row[COL_OBS_IPTU] or "").strip().lower()
        # So verifica lotes que deveriam ter PDF
        if not ("pdf baixado" in obs_low or "guia enviada e-mail cliente" in obs_low):
            continue
        # Exclui SEM_DEBITO e INDISPONIVEL (esses nao tem PDF mesmo)
        if "sem_debito" in obs_low or "sem debito" in obs_low:
            continue
        if "indisponivel" in obs_low or "boleto_indisponivel" in obs_low:
            continue

        cci = row[COL_CCI]
        if not cci:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue

        comprador = str(row[COL_COMPRADOR] or "").strip()
        etapa_n = extrair_etapa(row[COL_ETAPA])
        if not etapa_n or not comprador:
            continue
        etapa_str = f"ETAPA{etapa_n}"
        nome_esperado = limpar_nome(comprador)

        entry = {
            "linha": i, "cci": cci_str,
            "comprador": comprador, "nome_esperado": nome_esperado,
            "etapa": etapa_str,
        }

        if etapa_str not in disco:
            entry["motivo"] = f"etapa {etapa_str} nao existe no disco"
            sem_pasta.append(entry)
            continue

        # Match exato
        if nome_esperado in disco[etapa_str]:
            pasta_dir = disco[etapa_str][nome_esperado]
            pdfs = list(pasta_dir.glob("*.pdf"))
            # Procura PDF que contem o CCI
            tem_pdf_cci = any(f"CCI{cci_str}" in p.name for p in pdfs)
            if tem_pdf_cci:
                ok.append({**entry, "pasta_real": nome_esperado})
            else:
                entry["motivo"] = f"pasta tem {len(pdfs)} PDFs mas nenhum com CCI{cci_str}"
                entry["pasta_real"] = nome_esperado
                sem_pdf_na_pasta.append(entry)
            continue

        # Match aproximado (sem acentos)
        sem_acento_alvo = normalizar_sem_acento(nome_esperado)
        if sem_acento_alvo in disco_sem_acento.get(etapa_str, {}):
            pasta_real = disco_sem_acento[etapa_str][sem_acento_alvo]
            entry["motivo"] = (f"nome divergente: esperava '{nome_esperado}' "
                               f"mas no disco esta '{pasta_real}'")
            entry["pasta_real"] = pasta_real
            casamento_aproximado.append(entry)
            continue

        # Nao tem mesmo
        entry["motivo"] = f"pasta {etapa_str}/{nome_esperado} NAO EXISTE no disco"
        sem_pasta.append(entry)

    wb.close()

    # Resumo
    total = len(ok) + len(sem_pdf_na_pasta) + len(sem_pasta) + len(casamento_aproximado)
    print("\n" + "=" * 75)
    print(f"  RESUMO")
    print("=" * 75)
    print(f"  Total verificado (lotes que deveriam ter PDF) : {total}")
    print(f"  OK (pasta + PDF presentes)                    : {len(ok)}")
    print(f"  PASTA OK mas sem PDF do CCI esperado          : {len(sem_pdf_na_pasta)}")
    print(f"  Match aproximado (nome divergente)            : {len(casamento_aproximado)}")
    print(f"  PASTA FALTANDO no disco                       : {len(sem_pasta)}")
    print("=" * 75)

    if sem_pasta:
        print(f"\n=== PASTAS FALTANDO (primeiras 30) ===")
        for e in sem_pasta[:30]:
            print(f"  CCI {e['cci']:>8} | linha {e['linha']:>5} | "
                  f"{e['etapa']}/{e['nome_esperado'][:40]:<40} | {e['motivo'][:60]}")

    if casamento_aproximado:
        print(f"\n=== NOMES DIVERGENTES (primeiras 30) ===")
        for e in casamento_aproximado[:30]:
            print(f"  CCI {e['cci']:>8} | {e['etapa']}/{e['nome_esperado'][:30]:<30}")
            print(f"             disco tem: {e['pasta_real']}")

    if sem_pdf_na_pasta:
        print(f"\n=== PASTAS SEM O PDF DO CCI (primeiras 30) ===")
        for e in sem_pdf_na_pasta[:30]:
            print(f"  CCI {e['cci']:>8} | {e['etapa']}/{e['nome_esperado'][:40]}")
            print(f"             {e['motivo']}")

    # CSV
    if (sem_pasta or sem_pdf_na_pasta or casamento_aproximado):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = Path(f"verificacao_pastas_{ts}.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("tipo;linha;cci;etapa;nome_esperado;pasta_real;comprador_planilha;motivo\n")
            for tipo, lista in (
                ("SEM_PASTA", sem_pasta),
                ("NOME_DIVERGENTE", casamento_aproximado),
                ("SEM_PDF_CCI", sem_pdf_na_pasta),
            ):
                for e in lista:
                    f.write(f"{tipo};{e['linha']};{e['cci']};{e['etapa']};"
                            f"{e['nome_esperado']};{e.get('pasta_real','')};"
                            f"{e['comprador']};{e['motivo']}\n")
        print(f"\n[OK] CSV: {csv_path.resolve()}")


if __name__ == "__main__":
    main()
