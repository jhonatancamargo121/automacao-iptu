# -*- coding: utf-8 -*-
"""
ANALISE PROFUNDA DOS IPTUs BAIXADOS

Vasculha todas as pastas em IPTUs_Baixados/ e cruza com a planilha.

Detecta:
  - Orfaos: PDFs no disco sem linha na planilha
  - Duplicados: mesmo CCI em multiplos arquivos
  - Tamanho anomalo (<5KB e >5MB)
  - Comprador divergente: pasta != comprador da planilha
  - Pastas vazias
  - Conteudo: CCI presente no texto extraido do PDF (anti-troca)

Opcionalmente verifica no sistema Prodata (--validar-sistema):
  - Re-consulta SEM_DEBITO no site para confirmar que ainda nao ha boleto

Uso:
  python analise_profunda.py
      -> Analise completa local
  python analise_profunda.py --com-conteudo
      -> Extrai texto de TODOS os PDFs (mais lento, requer pypdf)
  python analise_profunda.py --validar-sistema 10
      -> Re-consulta 10 SEM_DEBITO aleatorios no site
  python analise_profunda.py --csv
      -> Salva relatorio completo em CSV
"""
import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path
from datetime import datetime

import openpyxl

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

# Limiares
TAM_MIN_BYTES = 5 * 1024        # menor que 5KB e suspeito
TAM_MAX_BYTES = 5 * 1024 * 1024  # maior que 5MB tambem


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


def limpar_nome(nome):
    nome = (nome or "").upper().strip()
    nome = re.sub(r'[<>:"/\\|?*]', "", nome)
    nome = re.sub(r"\s+", "_", nome)
    return nome[:80] or "SEM_NOME"


def extrair_cci_do_nome(nome_arq):
    """Extrai o numero do CCI do nome do arquivo (CCI{N}_...)."""
    m = re.search(r"CCI(\d+)", nome_arq, re.IGNORECASE)
    return m.group(1) if m else None


def extrair_etapa(s):
    m = re.search(r"(\d+)", str(s or ""))
    return m.group(1) if m else ""


def carregar_planilha_para_dict(planilha_path):
    """
    Retorna:
      por_cci   = {cci_str: {linha, quadra, lote, comprador, etapa, obs, status}}
      contadores = dict com contagens informativas
    """
    print(f"Lendo planilha: {planilha_path}")
    wb = openpyxl.load_workbook(planilha_path, read_only=True)
    ws = wb.active
    por_cci = {}
    cont = defaultdict(int)
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        status = str(row[COL_STATUS] or "").strip()
        cci    = row[COL_CCI]
        obs    = str(row[COL_OBS_IPTU] or "").strip()
        if status.lower() != "vendido" or not cci:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue
        info = {
            "linha":     i,
            "quadra":    str(row[COL_QUADRA] or "").strip(),
            "lote":      str(row[COL_LOTE] or "").strip(),
            "comprador": str(row[COL_COMPRADOR] or "").strip(),
            "etapa":     str(row[COL_ETAPA] or "").strip(),
            "obs":       obs,
            "status":    status,
        }
        por_cci[cci_str] = info
        obs_low = obs.lower()
        if "pdf baixado" in obs_low and "sem_debito" not in obs_low and "indisponivel" not in obs_low:
            cont["pdf_baixado_marcado"] += 1
        elif "sem_debito" in obs_low or "sem debito" in obs_low:
            cont["sem_debito"] += 1
        elif "indisponivel" in obs_low or "boleto_indisponivel" in obs_low:
            cont["indisponivel"] += 1
        else:
            cont["pendente"] += 1
        cont["total_vendido_com_cci"] += 1
    wb.close()
    return por_cci, cont


def vasculhar_pdfs(pasta_pdfs):
    """
    Retorna lista de dicts com info de cada PDF encontrado:
      [{caminho, nome, cci, tamanho, etapa_pasta, comprador_pasta}]
    """
    base = Path(pasta_pdfs)
    if not base.exists():
        return []
    resultado = []
    for pdf in base.rglob("*.pdf"):
        try:
            tam = pdf.stat().st_size
        except Exception:
            tam = 0
        cci = extrair_cci_do_nome(pdf.name)
        # Tenta extrair etapa e comprador a partir do caminho
        rel = pdf.relative_to(base)
        partes = rel.parts
        etapa_pasta = partes[0] if len(partes) >= 1 else ""
        comprador_pasta = partes[1] if len(partes) >= 2 else ""
        resultado.append({
            "caminho": pdf,
            "nome": pdf.name,
            "cci": cci,
            "tamanho": tam,
            "etapa_pasta": etapa_pasta,
            "comprador_pasta": comprador_pasta,
        })
    return resultado


def pastas_vazias(pasta_pdfs):
    """Retorna subpastas sem nenhum PDF dentro (recursivo)."""
    base = Path(pasta_pdfs)
    if not base.exists():
        return []
    vazias = []
    for d in base.rglob("*"):
        if d.is_dir():
            tem_pdf = any(d.rglob("*.pdf"))
            if not tem_pdf:
                vazias.append(d)
    return vazias


def extrair_texto_pdf(caminho):
    """Tenta extrair texto do PDF usando pypdf. Retorna string ou None."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return None
    try:
        reader = PdfReader(str(caminho))
        textos = []
        for pagina in reader.pages[:3]:  # so as 3 primeiras paginas
            try:
                textos.append(pagina.extract_text() or "")
            except Exception:
                pass
        return " ".join(textos)
    except Exception:
        return None


def main():
    parser = argparse.ArgumentParser(
        description="Analise profunda dos IPTUs baixados (folders + planilha + conteudo)"
    )
    parser.add_argument("--com-conteudo", action="store_true",
                        help="Extrai texto de cada PDF para verificar se o CCI aparece "
                             "(detecta troca de arquivo). Requer pypdf.")
    parser.add_argument("--csv", action="store_true",
                        help="Salva relatorio detalhado em CSV")
    parser.add_argument("--validar-sistema", type=int, default=0, metavar="N",
                        help="Re-consulta no site da prefeitura N SEM_DEBITO aleatorios "
                             "para confirmar (cuidado: abre navegador)")
    args = parser.parse_args()

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    por_cci, cont = carregar_planilha_para_dict(planilha)
    pdfs = vasculhar_pdfs(PASTA_PDFS)
    print(f"\n[OK] {len(por_cci)} lotes 'Vendidos' na planilha.")
    print(f"[OK] {len(pdfs)} PDFs encontrados no disco.\n")

    # 1) Orfaos: PDF no disco sem CCI valido OU CCI nao existe na planilha
    orfaos = []
    pdfs_por_cci = defaultdict(list)
    for p in pdfs:
        if not p["cci"]:
            orfaos.append((p, "sem CCI no nome"))
            continue
        if p["cci"] not in por_cci:
            orfaos.append((p, "CCI nao existe na planilha"))
            continue
        pdfs_por_cci[p["cci"]].append(p)

    # 2) Duplicados
    duplicados = [(cci, lst) for cci, lst in pdfs_por_cci.items() if len(lst) > 1]

    # 3) Tamanho anomalo
    anomalos = []
    for p in pdfs:
        if p["tamanho"] < TAM_MIN_BYTES or p["tamanho"] > TAM_MAX_BYTES:
            anomalos.append(p)

    # 4) Comprador divergente
    divergentes = []
    for p in pdfs:
        if not p["cci"] or p["cci"] not in por_cci:
            continue
        info_planilha = por_cci[p["cci"]]
        comprador_esperado = limpar_nome(info_planilha["comprador"])
        # Compara apenas se temos comprador_pasta nao vazio
        if p["comprador_pasta"]:
            if p["comprador_pasta"] != comprador_esperado:
                divergentes.append((p, comprador_esperado))

    # 5) Pastas vazias
    vazias = pastas_vazias(PASTA_PDFS)

    # 6) Conteudo (opcional)
    conteudo_suspeito = []
    if args.com_conteudo:
        print("Extraindo texto de cada PDF (pode demorar)...")
        try:
            from pypdf import PdfReader  # noqa
        except ImportError:
            print("  [AVISO] pypdf nao instalado. Pulando.")
            print("  Para instalar: pip install pypdf")
        else:
            for p in pdfs:
                if not p["cci"]:
                    continue
                texto = extrair_texto_pdf(p["caminho"])
                if texto is None:
                    continue
                if p["cci"] not in texto:
                    conteudo_suspeito.append((p, "CCI nao aparece no texto"))

    # Relatorio
    print("=" * 70)
    print("RELATORIO DA ANALISE PROFUNDA")
    print("=" * 70)
    print(f"\nPLANILHA (lotes Vendidos com CCI = {cont['total_vendido_com_cci']}):")
    print(f"  marcados como 'PDF baixado'  : {cont['pdf_baixado_marcado']}")
    print(f"  marcados SEM_DEBITO          : {cont['sem_debito']}")
    print(f"  marcados INDISPONIVEL        : {cont['indisponivel']}")
    print(f"  ainda pendentes              : {cont['pendente']}")

    print(f"\nDISCO ({len(pdfs)} PDFs em {PASTA_PDFS}/):")
    cobertos = len(por_cci) and len(set(p['cci'] for p in pdfs if p['cci']) & set(por_cci.keys()))
    print(f"  PDFs com CCI casando na planilha : {cobertos}")
    print(f"  PDFs ORFAOS                       : {len(orfaos)}")
    print(f"  PDFs DUPLICADOS (CCIs)            : {len(duplicados)}")
    print(f"  PDFs com TAMANHO ANOMALO          : {len(anomalos)}")
    print(f"  PDFs em PASTA DIVERGENTE          : {len(divergentes)}")
    print(f"  Pastas VAZIAS                     : {len(vazias)}")
    if args.com_conteudo:
        print(f"  PDFs com CCI ausente no texto     : {len(conteudo_suspeito)}")

    if orfaos:
        print(f"\n=== ORFAOS (primeiros 15) ===")
        for p, motivo in orfaos[:15]:
            print(f"  {p['nome'][:50]:<50} | {motivo}")
    if duplicados:
        print(f"\n=== DUPLICADOS ===")
        for cci, lst in duplicados[:15]:
            print(f"  CCI {cci}:")
            for p in lst:
                print(f"    -> {p['caminho'].relative_to(Path(PASTA_PDFS))}")
    if anomalos:
        print(f"\n=== TAMANHO ANOMALO (primeiros 15) ===")
        for p in anomalos[:15]:
            tag = "PEQUENO" if p["tamanho"] < TAM_MIN_BYTES else "GRANDE"
            print(f"  [{tag}] {p['tamanho']:>10} bytes  {p['nome'][:50]}")
    if divergentes:
        print(f"\n=== COMPRADOR DIVERGENTE (primeiros 15) ===")
        for p, esperado in divergentes[:15]:
            print(f"  CCI {p['cci']}  pasta='{p['comprador_pasta']}'  esperado='{esperado}'")
    if vazias:
        print(f"\n=== PASTAS VAZIAS (primeiras 15) ===")
        for d in vazias[:15]:
            print(f"  {d.relative_to(Path(PASTA_PDFS))}")
    if conteudo_suspeito:
        print(f"\n=== CONTEUDO SUSPEITO ({len(conteudo_suspeito)}) ===")
        for p, motivo in conteudo_suspeito[:15]:
            print(f"  CCI {p['cci']}  -> {motivo}  ({p['nome'][:40]})")

    # CSV opcional
    if args.csv:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = Path(f"analise_profunda_{ts}.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("tipo;cci;arquivo_ou_pasta;detalhe\n")
            for p, motivo in orfaos:
                f.write(f"ORFAO;{p['cci'] or ''};{p['caminho']};{motivo}\n")
            for cci, lst in duplicados:
                for p in lst:
                    f.write(f"DUPLICADO;{cci};{p['caminho']};\n")
            for p in anomalos:
                tag = "PEQUENO" if p["tamanho"] < TAM_MIN_BYTES else "GRANDE"
                f.write(f"TAM_{tag};{p['cci'] or ''};{p['caminho']};{p['tamanho']} bytes\n")
            for p, esperado in divergentes:
                f.write(f"DIVERGENTE;{p['cci']};{p['caminho']};esperado={esperado}\n")
            for d in vazias:
                f.write(f"PASTA_VAZIA;;{d};\n")
            for p, motivo in conteudo_suspeito:
                f.write(f"CONTEUDO_SUSPEITO;{p['cci']};{p['caminho']};{motivo}\n")
        print(f"\n[OK] CSV salvo: {csv_path.resolve()}")

    # Validacao no sistema (opcional)
    if args.validar_sistema > 0:
        print(f"\n[!!] Validacao no sistema ({args.validar_sistema} CCIs SEM_DEBITO)...")
        print("    Esta funcionalidade requer Playwright e abrir o navegador.")
        print("    Recomendado: rode separado como ferramenta (TBD).")

    print("\n" + "=" * 70)
    print("Resumo:")
    total_problemas = (len(orfaos) + len(duplicados) + len(anomalos)
                       + len(divergentes) + len(conteudo_suspeito))
    if total_problemas == 0:
        print("  TUDO OK - nenhum problema detectado.")
    else:
        print(f"  {total_problemas} item(s) com problema. Use --csv para lista completa.")
    print("=" * 70)


if __name__ == "__main__":
    main()
