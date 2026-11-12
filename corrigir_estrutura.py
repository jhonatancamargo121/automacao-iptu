# -*- coding: utf-8 -*-
"""
CORRIGIR ESTRUTURA LOCAL - Natela Analytics

Para cada CCI marcado como 'PDF baixado' ou 'Guia enviada' na planilha:
  1. Procura o PDF em TODAS as pastas IPTUs_*/ (recursivo)
  2. Se achou mas esta em ETAPA/COMPRADOR errado, MOVE pra correto
  3. Se nao achou em nenhum lugar, lista como FALTANTE_REAL (precisa re-baixar)

Destino padrao: IPTUs_Final_Consolidado/ETAPA{N}/{COMPRADOR}/

Apos rodar isso, a estrutura local fica certinha. Voce sincroniza com o
OneDrive (3 opcoes - veja README no final).

Uso:
  python corrigir_estrutura.py
      -> dry-run: mostra o que faria sem mover nada
  python corrigir_estrutura.py --executar
      -> faz os moves de verdade
  python corrigir_estrutura.py --destino IPTUs_Para_OneDrive --executar
      -> usa outro destino
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
PASTA_DESTINO_PADRAO = "IPTUs_Final_Consolidado"

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
    nome = (nome or "").upper().strip()
    nome = re.sub(r'[<>:"/\\|?*]', "", nome)
    nome = re.sub(r"\s+", "_", nome)
    return nome[:80] or "SEM_NOME"


def extrair_etapa(s):
    m = re.search(r"(\d+)", str(s or ""))
    return m.group(1) if m else ""


def carregar_esperados(planilha_path):
    """Retorna lista de dicts: lotes que deveriam ter PDF."""
    print(f"Lendo planilha: {planilha_path}")
    wb = None
    for _ in range(8):
        try:
            wb = openpyxl.load_workbook(planilha_path, read_only=True)
            break
        except Exception:
            time.sleep(2)
    if wb is None:
        print("[ERRO] Planilha sendo escrita.")
        sys.exit(1)
    ws = wb.active

    esperados = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        obs_low = str(row[COL_OBS_IPTU] or "").strip().lower()
        if not ("pdf baixado" in obs_low or "guia enviada e-mail cliente" in obs_low):
            continue
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
        esperados.append({
            "linha": i,
            "cci": cci_str,
            "comprador": comprador,
            "etapa": etapa_n,
            "etapa_str": f"ETAPA{etapa_n}",
            "nome_pasta": limpar_nome(comprador),
        })
    wb.close()
    return esperados


def indexar_pdfs_existentes():
    """
    Varre todas as pastas IPTUs_*/ e devolve {cci: [list of (caminho, source)]}.
    """
    pdfs_por_cci = defaultdict(list)
    pastas = [d for d in Path(".").iterdir()
              if d.is_dir() and d.name.lower().startswith("iptus_")]
    for base in pastas:
        for pdf in base.rglob("*.pdf"):
            m = re.search(r"CCI(\d+)", pdf.name, re.I)
            if m:
                pdfs_por_cci[m.group(1)].append((pdf, base.name))
    return pdfs_por_cci, [p.name for p in pastas]


def main():
    parser = argparse.ArgumentParser(
        description="Corrige a estrutura local: move cada PDF pra ETAPA/COMPRADOR certo"
    )
    parser.add_argument("--executar", action="store_true",
                        help="Faz os moves de verdade. Sem essa flag eh dry-run.")
    parser.add_argument("--destino", type=str, default=PASTA_DESTINO_PADRAO,
                        help=f"Pasta de destino (padrao: {PASTA_DESTINO_PADRAO})")
    args = parser.parse_args()

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    print("=" * 75)
    print("  Natela Analytics - Corrigir Estrutura Local")
    print(f"  Modo: {'EXECUTAR (move arquivos)' if args.executar else 'DRY-RUN (apenas relata)'}")
    print("=" * 75)

    destino_base = Path(args.destino)
    esperados = carregar_esperados(planilha)
    print(f"\nLotes que deveriam ter PDF: {len(esperados)}")

    pdfs_por_cci, pastas_origem = indexar_pdfs_existentes()
    print(f"\nPastas escaneadas:")
    for p in pastas_origem:
        cnt = sum(1 for _ in Path(p).rglob('*.pdf'))
        print(f"  - {p:<40} {cnt:>5} PDFs")
    print(f"\nTotal de CCIs unicos com PDF em disco: {len(pdfs_por_cci)}")

    # Classifica cada esperado
    ja_no_lugar       = []   # PDF ja esta no destino correto
    mover_pra_destino = []   # PDF existe em outra pasta, precisa mover
    copiar_pra_destino = []  # PDF existe em outra pasta, copiar (manter origem)
    faltante_real     = []   # nao tem PDF em lugar nenhum

    for e in esperados:
        cci = e["cci"]
        destino_pasta = destino_base / e["etapa_str"] / e["nome_pasta"]
        pdfs = pdfs_por_cci.get(cci, [])
        if not pdfs:
            faltante_real.append(e)
            continue
        # Procura se ja tem no destino correto
        ja_no_destino = None
        for pdf, src in pdfs:
            try:
                if pdf.is_relative_to(destino_pasta):
                    ja_no_destino = pdf
                    break
            except ValueError:
                pass
            # Fallback se Python < 3.9
            try:
                pdf.relative_to(destino_pasta)
                ja_no_destino = pdf
                break
            except ValueError:
                pass
        if ja_no_destino:
            ja_no_lugar.append({**e, "pdf": ja_no_destino})
            continue
        # PDF existe mas em outra pasta - vamos pegar o primeiro
        pdf_origem, src = pdfs[0]
        # Se origem eh outra pasta IPTUs_* (nao a destino), MOVE
        # Se origem ja eh em destino_base mas em subpasta errada, MOVE
        mover_pra_destino.append({
            **e, "pdf_origem": pdf_origem, "pasta_origem": src,
            "destino": destino_pasta,
        })

    # Resumo
    print("\n" + "=" * 75)
    print("  RESUMO DA CORRECAO")
    print("=" * 75)
    print(f"  Lotes esperados (com PDF baixado)         : {len(esperados)}")
    print(f"  Ja estao no lugar correto                 : {len(ja_no_lugar)}")
    print(f"  Vao ser MOVIDOS pra ETAPA/COMPRADOR certo : {len(mover_pra_destino)}")
    print(f"  FALTANTES REAIS (nao tem PDF em lugar nenhum): {len(faltante_real)}")
    print("=" * 75)

    if mover_pra_destino:
        print(f"\n=== A MOVER (primeiros 20 de {len(mover_pra_destino)}) ===")
        for e in mover_pra_destino[:20]:
            print(f"  CCI {e['cci']:>8} | {e['pasta_origem']:<35} -> {e['etapa_str']}/{e['nome_pasta']}")

    if faltante_real:
        print(f"\n=== FALTANTES REAIS ({len(faltante_real)} - PRECISA REBAIXAR) ===")
        for e in faltante_real[:20]:
            print(f"  CCI {e['cci']:>8} | linha {e['linha']:>5} | "
                  f"{e['etapa_str']}/{e['nome_pasta'][:40]}")
        if len(faltante_real) > 20:
            print(f"  ... e mais {len(faltante_real) - 20} no CSV")

    # CSV
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = Path(f"correcao_estrutura_{ts}.csv")
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write("acao;cci;etapa;nome_pasta;origem;destino;linha_planilha\n")
        for e in ja_no_lugar:
            f.write(f"OK;{e['cci']};{e['etapa_str']};{e['nome_pasta']};{e.get('pdf','')};;{e['linha']}\n")
        for e in mover_pra_destino:
            f.write(f"MOVER;{e['cci']};{e['etapa_str']};{e['nome_pasta']};"
                    f"{e['pdf_origem']};{e['destino']};{e['linha']}\n")
        for e in faltante_real:
            f.write(f"FALTANTE_REAL;{e['cci']};{e['etapa_str']};{e['nome_pasta']};;;{e['linha']}\n")
    print(f"\n[OK] CSV: {csv_path.resolve()}")

    # Executa moves se pedido
    if args.executar and mover_pra_destino:
        resp = input(f"\nVai MOVER {len(mover_pra_destino)} PDFs. Confirma? [s/N]: ").strip().lower()
        if resp != "s":
            print("Cancelado.")
            return
        sucesso = erros = 0
        for e in mover_pra_destino:
            try:
                e["destino"].mkdir(parents=True, exist_ok=True)
                dst = e["destino"] / e["pdf_origem"].name
                if dst.exists():
                    ts2 = datetime.now().strftime("%H%M%S")
                    dst = e["destino"] / f"{dst.stem}_{ts2}.pdf"
                shutil.move(str(e["pdf_origem"]), str(dst))
                sucesso += 1
            except Exception as ex:
                print(f"  [ERRO] CCI {e['cci']}: {ex}")
                erros += 1
        print(f"\n[OK] {sucesso} PDFs movidos. {erros} erros.")

        # Limpa pastas vazias nas origens
        print("\nLimpando pastas vazias...")
        for p in sorted(Path(".").rglob("*"), key=lambda x: -len(x.parts)):
            if p.is_dir():
                try:
                    p.rmdir()
                except OSError:
                    pass
    elif mover_pra_destino:
        print(f"\n>>> DRY-RUN: nada foi movido. Rode com --executar pra aplicar.")

    print("\n" + "=" * 75)
    print("  PROXIMO PASSO - SINCRONIZAR COM ONEDRIVE")
    print("=" * 75)
    print("  Opcao 1 (mais simples): instalar o cliente OneDrive Desktop e")
    print("    sincronizar a pasta IPTUs_Final_Consolidado/. Tudo que voce")
    print("    arrumar localmente aparece automaticamente na nuvem.")
    print()
    print("  Opcao 2: zipar a pasta IPTUs_Final_Consolidado e fazer upload")
    print("    manual pelo navegador (arrastar e soltar). OneDrive aceita ate")
    print("    100GB por upload.")
    print()
    print("  Opcao 3: instalar 'rclone' (gratuito, open-source) e rodar:")
    print("    rclone sync IPTUs_Final_Consolidado/ remote:IPTUs_Baixados/")
    print("  Sincroniza incremental - so manda diferenca, varias vezes mais rapido.")
    print("=" * 75)


if __name__ == "__main__":
    main()
