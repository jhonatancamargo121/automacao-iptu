# -*- coding: utf-8 -*-
"""
MESCLAR PASTAS - Natela Analytics

Junta os IPTUs baixados em diferentes pastas (datas diferentes) numa
unica pasta consolidada. Mantem a estrutura ETAPA{N}/{COMPRADOR}/.

Regra de prioridade: a primeira pasta da lista PASTAS_ORIGEM "vence"
quando ha o mesmo CCI em duas pastas. Use isso para preferir a data
mais recente (29/06) e deixar a 29/05 so para os faltantes.

Uso:
  python mesclar_pastas.py
      -> mostra o que vai fazer e pede confirmacao
  python mesclar_pastas.py --executar
      -> roda sem perguntar (atalho)
  python mesclar_pastas.py --destino IPTUs_Final --executar
      -> sobrescreve a pasta destino
"""
import argparse
import re
import shutil
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# Ordem de prioridade: a PRIMEIRA pasta da lista vence se o CCI existir em
# mais de uma. Em geral: pasta com data mais recente / mais correta primeiro.
PASTAS_ORIGEM = [
    "IPTUs_Baixados_29_06_2026",          # data principal (29/06)
    "IPTUs_Baixados_29_05_2026_RETRY",    # retry com data 29/05 dos faltantes
    "IPTUs_Baixados",                     # rodada inicial (preservada)
]

PASTA_DESTINO_PADRAO = "IPTUs_Final_Consolidado"


def extrair_cci(nome):
    m = re.search(r"CCI(\d+)", nome, re.I)
    return m.group(1) if m else None


def main():
    parser = argparse.ArgumentParser(
        description="Mescla pastas de IPTUs baixados em datas diferentes"
    )
    parser.add_argument("--executar", action="store_true",
                        help="Pula a confirmacao e copia direto")
    parser.add_argument("--destino", type=str, default=PASTA_DESTINO_PADRAO,
                        help=f"Pasta destino (padrao: {PASTA_DESTINO_PADRAO})")
    parser.add_argument("--origens", nargs="+", default=None,
                        help="Sobrescreve a lista de pastas origem (em ordem de prioridade)")
    parser.add_argument("--mover", action="store_true",
                        help="Move arquivos em vez de copiar (cuidado: esvazia origem)")
    args = parser.parse_args()

    origens = args.origens or PASTAS_ORIGEM
    destino = Path(args.destino)

    print("=" * 70)
    print("  Natela Analytics - Mesclagem de Pastas IPTU")
    print("=" * 70)
    print(f"\nPasta destino: {destino.resolve()}")
    print(f"\nPastas origem (em ordem de prioridade):")
    for i, p in enumerate(origens, 1):
        path = Path(p)
        existe = path.exists()
        cnt = sum(1 for _ in path.rglob("*.pdf")) if existe else 0
        marca = "OK" if existe else "NAO EXISTE"
        print(f"  {i}. {p:<40} [{marca}]  PDFs={cnt}")

    # Inventario por CCI
    inventario = defaultdict(list)   # cci -> [(prioridade, caminho_origem, pasta_origem)]
    for prio, pasta_origem_nome in enumerate(origens):
        pasta_origem = Path(pasta_origem_nome)
        if not pasta_origem.exists():
            continue
        for pdf in pasta_origem.rglob("*.pdf"):
            cci = extrair_cci(pdf.name)
            if not cci:
                continue
            inventario[cci].append((prio, pdf, pasta_origem))

    if not inventario:
        print("\n[ERRO] Nenhum PDF encontrado nas pastas origem.")
        sys.exit(1)

    # Para cada CCI, escolhe o de maior prioridade (menor numero)
    a_copiar = []   # [(cci, caminho_origem, pasta_origem)]
    descartados = []
    for cci, lista in inventario.items():
        lista.sort(key=lambda x: x[0])
        primeiro = lista[0]
        a_copiar.append((cci, primeiro[1], primeiro[2]))
        for outro in lista[1:]:
            descartados.append((cci, outro[1], outro[2]))

    print(f"\nInventario:")
    print(f"  CCIs unicos encontrados : {len(inventario)}")
    print(f"  PDFs a copiar/mover     : {len(a_copiar)}")
    print(f"  PDFs duplicados (descartados pela prioridade) : {len(descartados)}")

    # Quebra por pasta origem (mostra de onde vem o que vai pra pasta final)
    por_origem = defaultdict(int)
    for cci, _, pasta_origem in a_copiar:
        por_origem[pasta_origem.name] += 1
    print(f"\nDistribuicao no destino por pasta origem:")
    for nome, cnt in sorted(por_origem.items(), key=lambda kv: -kv[1]):
        print(f"  {nome:<40} {cnt:>5} PDFs")

    if descartados:
        print(f"\nAmostra de descartados (CCI ja vem de pasta de maior prioridade):")
        for cci, pdf, _ in descartados[:10]:
            print(f"  CCI {cci}  descartado: {pdf}")

    acao = "MOVER" if args.mover else "COPIAR"
    print(f"\n>>> ACAO: {acao} {len(a_copiar)} PDFs para '{destino}/'")
    print(f"    Estrutura ETAPA/COMPRADOR sera preservada.")

    if not args.executar:
        resp = input(f"\nConfirma {acao}? [s/N]: ").strip().lower()
        if resp != "s":
            print("Cancelado.")
            return

    # Executa a copia/mover
    destino.mkdir(parents=True, exist_ok=True)
    copiados = 0
    erros = 0
    log = []
    for cci, pdf, pasta_origem in a_copiar:
        try:
            rel = pdf.relative_to(pasta_origem)
            destino_arq = destino / rel
            destino_arq.parent.mkdir(parents=True, exist_ok=True)
            if args.mover:
                shutil.move(str(pdf), str(destino_arq))
            else:
                shutil.copy2(str(pdf), str(destino_arq))
            copiados += 1
            log.append(("OK", cci, str(rel), pasta_origem.name))
        except Exception as e:
            erros += 1
            log.append(("ERRO", cci, str(pdf), str(e)))

    # Salva log
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path(f"mesclagem_{ts}.log")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"# Mesclagem {ts}\n")
        f.write(f"# Destino: {destino.resolve()}\n")
        f.write(f"# Origens (prioridade): {origens}\n")
        for entry in log:
            f.write("\t".join(str(x) for x in entry) + "\n")

    print(f"\n[OK] {copiados} PDFs {acao.lower()}ados para '{destino}/'")
    if erros:
        print(f"[!!] {erros} erros (ver {log_path})")
    print(f"     Log detalhado: {log_path.resolve()}")
    print(f"     Total no destino agora: {sum(1 for _ in destino.rglob('*.pdf'))} PDFs")


if __name__ == "__main__":
    main()
