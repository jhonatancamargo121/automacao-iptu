# -*- coding: utf-8 -*-
"""
CHECAGEM FINAL - Natela Analytics

Faz a validacao completa dos IPTUs antes de subir pra nuvem (OneDrive).

Verifica:
  1. INTEGRIDADE de cada PDF (existe, >5KB, header %PDF valido)
  2. NOMENCLATURA (formato ETAPA{N}_QD{X}_LT{Y}_CCI{Z}_NAME.pdf)
  3. ESTRUTURA DE PASTAS (ETAPA{N}/COMPRADOR/)
  4. CRUZAMENTO com planilha (todo CCI marcado tem PDF / todo PDF tem CCI)
  5. DUPLICATAS (mesmo CCI em multiplos arquivos)
  6. COMPRADOR DIVERGENTE (pasta != comprador da planilha)
  7. COBERTURA por etapa
  8. VEREDITO final: PRONTO ou REVISAR

Saidas:
  - Terminal com summary
  - HTML detalhado (checagem_final_natela_analytics.html)
  - CSV com itens problematicos

Uso:
  python checagem_final.py
  python checagem_final.py --pasta IPTUs_Final_Consolidado --abrir
"""
import argparse
import json
import re
import sys
import time
import webbrowser
from collections import defaultdict, Counter
from datetime import datetime
from html import escape
from pathlib import Path

import openpyxl

NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]
PASTA_PADRAO = "IPTUs_Final_Consolidado"
SAIDA_HTML   = "checagem_final_natela_analytics.html"

COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13

TAM_MIN = 5 * 1024            # 5KB
TAM_MAX = 5 * 1024 * 1024     # 5MB

PADRAO_NOME_OK = re.compile(
    r"^ETAPA(\d+)_QD([^_]+)_LT([^_]+)_CCI(\d+)_(.+)\.pdf$",
    re.IGNORECASE,
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


def limpar_nome(nome):
    nome = (nome or "").upper().strip()
    nome = re.sub(r'[<>:"/\\|?*]', "", nome)
    nome = re.sub(r"\s+", "_", nome)
    return nome[:80] or "SEM_NOME"


def extrair_cci(s):
    m = re.search(r"CCI(\d+)", s, re.I)
    return m.group(1) if m else None


def validar_pdf(caminho):
    """Verifica integridade do PDF. Retorna (valido, info)."""
    try:
        tam = caminho.stat().st_size
    except Exception as e:
        return False, f"erro stat: {e}"
    if tam == 0:
        return False, "0 bytes"
    if tam < TAM_MIN:
        return False, f"muito pequeno ({tam} bytes)"
    if tam > TAM_MAX:
        return False, f"muito grande ({tam:,} bytes)"
    try:
        with open(caminho, "rb") as f:
            cabecalho = f.read(8)
        if not cabecalho.startswith(b"%PDF"):
            return False, f"sem header %PDF (header={cabecalho[:6]!r})"
    except Exception as e:
        return False, f"erro leitura: {e}"
    return True, f"OK ({tam:,} bytes)"


def carregar_planilha(planilha_path):
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
    por_cci = {}
    cont = Counter()
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        if str(row[COL_STATUS] or "").strip().lower() != "vendido":
            continue
        cci = row[COL_CCI]
        if not cci:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue
        obs = str(row[COL_OBS_IPTU] or "").strip().lower()
        if "guia enviada e-mail cliente" in obs:
            cat = "enviado"
        elif "sem_debito" in obs or "sem debito" in obs:
            cat = "sem_debito"
        elif "indisponivel" in obs:
            cat = "indisponivel"
        elif "pdf baixado" in obs:
            cat = "baixado"
        else:
            cat = "pendente"
        por_cci[cci_str] = {
            "linha": i,
            "quadra": str(row[COL_QUADRA] or "").strip(),
            "lote":   str(row[COL_LOTE] or "").strip(),
            "comprador": str(row[COL_COMPRADOR] or "").strip(),
            "etapa": re.search(r"\d+", str(row[COL_ETAPA] or "")).group(0)
                     if re.search(r"\d+", str(row[COL_ETAPA] or "")) else "",
            "categoria": cat,
        }
        cont[cat] += 1
    wb.close()
    return por_cci, cont


def vasculhar_pasta(pasta):
    """Retorna lista de dicts: caminho, nome, cci, etapa_pasta, comprador_pasta, tamanho."""
    pasta = Path(pasta)
    res = []
    if not pasta.exists():
        return res
    for pdf in pasta.rglob("*.pdf"):
        try:
            tam = pdf.stat().st_size
        except Exception:
            tam = 0
        rel = pdf.relative_to(pasta).parts
        etapa_p = rel[0] if len(rel) >= 1 else ""
        comp_p  = rel[1] if len(rel) >= 2 else ""
        res.append({
            "caminho": pdf,
            "nome": pdf.name,
            "cci": extrair_cci(pdf.name) or "",
            "etapa_pasta": etapa_p,
            "comprador_pasta": comp_p,
            "tamanho": tam,
        })
    return res


def main():
    parser = argparse.ArgumentParser(description="Checagem final dos IPTUs - Natela Analytics")
    parser.add_argument("--pasta", type=str, default=PASTA_PADRAO,
                        help=f"Pasta a validar (padrao: {PASTA_PADRAO})")
    parser.add_argument("--abrir", action="store_true",
                        help="Abre o HTML no navegador apos gerar")
    args = parser.parse_args()

    print("=" * 70)
    print("  Natela Analytics - Checagem Final pra Upload na Nuvem")
    print("=" * 70)

    pasta = Path(args.pasta)
    if not pasta.exists():
        print(f"[ERRO] Pasta nao existe: {pasta.resolve()}")
        sys.exit(1)

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    por_cci, cont = carregar_planilha(planilha)
    pdfs = vasculhar_pasta(pasta)
    print(f"\nPasta validada: {pasta.resolve()}")
    print(f"PDFs encontrados: {len(pdfs)}")

    # 1. Integridade
    corrompidos = []     # arquivos com problema fisico
    for p in pdfs:
        ok, info = validar_pdf(p["caminho"])
        if not ok:
            corrompidos.append({**p, "motivo": info})

    # 2. Nomenclatura
    nome_invalido = []
    for p in pdfs:
        if not PADRAO_NOME_OK.match(p["nome"]):
            nome_invalido.append({**p, "motivo": "nome nao bate ETAPA{N}_QD{X}_LT{Y}_CCI{Z}_..."})

    # 3. Cruzamento PDF -> planilha
    sem_cci_no_nome = [p for p in pdfs if not p["cci"]]
    orfaos = [p for p in pdfs if p["cci"] and p["cci"] not in por_cci]

    # 4. Cruzamento planilha -> PDF (marcados como baixado mas sem PDF)
    ccis_no_disco = {p["cci"] for p in pdfs if p["cci"]}
    marcados_baixado = {c for c, i in por_cci.items() if i["categoria"] in ("baixado", "enviado")}
    falta_no_disco = marcados_baixado - ccis_no_disco

    # 5. Duplicados
    por_cci_pdfs = defaultdict(list)
    for p in pdfs:
        if p["cci"]:
            por_cci_pdfs[p["cci"]].append(p)
    duplicados = [(c, lst) for c, lst in por_cci_pdfs.items() if len(lst) > 1]

    # 6. Comprador divergente
    divergentes = []
    for p in pdfs:
        if p["cci"] and p["cci"] in por_cci and p["comprador_pasta"]:
            esperado = limpar_nome(por_cci[p["cci"]]["comprador"])
            if p["comprador_pasta"] != esperado:
                divergentes.append({**p, "esperado": esperado})

    # 7. Cobertura por etapa
    cobertura_etapa = defaultdict(lambda: {"esperado": 0, "no_disco": 0})
    for cci, info in por_cci.items():
        et = f"ETAPA{info['etapa']}" if info["etapa"] else "?"
        if info["categoria"] in ("baixado", "enviado"):
            cobertura_etapa[et]["esperado"] += 1
            if cci in ccis_no_disco:
                cobertura_etapa[et]["no_disco"] += 1

    # 8. Veredito
    total_problemas = (len(corrompidos) + len(nome_invalido) + len(sem_cci_no_nome)
                       + len(orfaos) + len(falta_no_disco) + len(duplicados)
                       + len(divergentes))
    pronto = (total_problemas == 0)

    # Terminal report
    print("\n" + "-" * 70)
    print("PLANILHA:")
    print(f"  Vendido com CCI    : {sum(cont.values())}")
    print(f"  baixado            : {cont.get('baixado', 0)}")
    print(f"  enviado            : {cont.get('enviado', 0)}")
    print(f"  sem_debito         : {cont.get('sem_debito', 0)}")
    print(f"  indisponivel       : {cont.get('indisponivel', 0)}")
    print(f"  pendente           : {cont.get('pendente', 0)}")
    print("\nDISCO:")
    print(f"  PDFs encontrados   : {len(pdfs)}")
    print(f"  CCIs unicos no disco: {len(ccis_no_disco)}")
    print("\nPROBLEMAS:")
    print(f"  Corrompidos        : {len(corrompidos)}")
    print(f"  Nome invalido      : {len(nome_invalido)}")
    print(f"  Sem CCI no nome    : {len(sem_cci_no_nome)}")
    print(f"  Orfaos (sem planilha): {len(orfaos)}")
    print(f"  Marcados sem PDF   : {len(falta_no_disco)}")
    print(f"  Duplicados (mesmo CCI): {len(duplicados)}")
    print(f"  Comprador divergente: {len(divergentes)}")
    print("\nCOBERTURA POR ETAPA:")
    for et in sorted(cobertura_etapa.keys()):
        d = cobertura_etapa[et]
        pct = (d["no_disco"] / d["esperado"] * 100) if d["esperado"] else 0
        print(f"  {et:<12}: {d['no_disco']}/{d['esperado']} PDFs ({pct:.1f}%)")
    print()
    print("=" * 70)
    if pronto:
        print("  >>> VEREDITO: PRONTO PARA UPLOAD NA NUVEM <<<")
    else:
        print(f"  >>> VEREDITO: REVISAR ({total_problemas} item(s) problematico(s)) <<<")
    print("=" * 70)

    # CSV detalhado
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if total_problemas:
        csv_path = Path(f"checagem_problemas_{ts}.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("tipo;cci;arquivo;detalhe\n")
            for p in corrompidos:
                f.write(f"CORROMPIDO;{p['cci']};{p['caminho']};{p['motivo']}\n")
            for p in nome_invalido:
                f.write(f"NOME_INVALIDO;{p['cci']};{p['caminho']};{p['motivo']}\n")
            for p in sem_cci_no_nome:
                f.write(f"SEM_CCI;;{p['caminho']};nome sem CCI{'{N}'}\n")
            for p in orfaos:
                f.write(f"ORFAO;{p['cci']};{p['caminho']};CCI nao existe na planilha\n")
            for cci in falta_no_disco:
                info = por_cci.get(cci, {})
                f.write(f"FALTA_PDF;{cci};(no disco);"
                        f"marcado como {info.get('categoria','?')} na planilha\n")
            for cci, lst in duplicados:
                for p in lst:
                    f.write(f"DUPLICADO;{cci};{p['caminho']};\n")
            for p in divergentes:
                f.write(f"DIVERGENTE;{p['cci']};{p['caminho']};esperado={p['esperado']}\n")
        print(f"\nCSV detalhado: {csv_path.resolve()}")

    # HTML
    html_path = Path(SAIDA_HTML)
    html = gerar_html({
        "pasta": str(pasta.resolve()),
        "planilha": planilha.name,
        "agora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "pronto": pronto,
        "total_problemas": total_problemas,
        "planilha_cont": dict(cont),
        "total_pdfs": len(pdfs),
        "ccis_no_disco": len(ccis_no_disco),
        "corrompidos": corrompidos,
        "nome_invalido": nome_invalido,
        "orfaos": orfaos,
        "falta_no_disco": [(c, por_cci.get(c, {})) for c in falta_no_disco],
        "duplicados": duplicados,
        "divergentes": divergentes,
        "cobertura_etapa": {k: dict(v) for k, v in cobertura_etapa.items()},
    })
    html_path.write_text(html, encoding="utf-8")
    print(f"HTML: {html_path.resolve()}")

    if args.abrir:
        webbrowser.open(html_path.resolve().as_uri())


def gerar_html(d):
    cor_veredito = "#16a34a" if d["pronto"] else "#dc2626"
    txt_veredito = "PRONTO PARA UPLOAD" if d["pronto"] else f"REVISAR ({d['total_problemas']} problemas)"

    def tabela_problemas(titulo, items, cols):
        rows = []
        for it in items[:50]:
            row = "".join(f"<td>{escape(str(it.get(c, '')))}</td>" for c in cols)
            rows.append(f"<tr>{row}</tr>")
        if not rows:
            rows = [f"<tr><td colspan='{len(cols)}' class='vazio'>Nenhum.</td></tr>"]
        ths = "".join(f"<th>{escape(c)}</th>" for c in cols)
        n = len(items)
        ctd = f" ({n})" if n else ""
        return f"<h3>{escape(titulo)}{ctd}</h3><table><thead><tr>{ths}</tr></thead><tbody>{''.join(rows)}</tbody></table>"

    cob_html = "".join(
        f"<tr><td>{escape(k)}</td><td class='num'>{v['no_disco']}/{v['esperado']}</td>"
        f"<td class='num'>{(v['no_disco']/v['esperado']*100 if v['esperado'] else 0):.1f}%</td></tr>"
        for k, v in sorted(d["cobertura_etapa"].items())
    )

    corrompidos_html = tabela_problemas(
        "PDFs Corrompidos", d["corrompidos"], ["cci", "nome", "tamanho", "motivo"])
    nome_html = tabela_problemas(
        "Nomenclatura Invalida", d["nome_invalido"], ["cci", "nome", "motivo"])
    orfaos_html = tabela_problemas(
        "Orfaos (no disco sem planilha)", d["orfaos"], ["cci", "nome", "etapa_pasta"])

    falta_rows = []
    for cci, info in d["falta_no_disco"][:50]:
        falta_rows.append(
            f"<tr><td>{escape(cci)}</td><td>{escape(info.get('comprador',''))}</td>"
            f"<td>ETAPA{info.get('etapa','')}</td><td>{escape(info.get('categoria',''))}</td></tr>"
        )
    if not falta_rows:
        falta_rows = ["<tr><td colspan='4' class='vazio'>Nenhum.</td></tr>"]
    falta_html = (f"<h3>Marcados como baixado mas SEM PDF no disco ({len(d['falta_no_disco'])})</h3>"
                  f"<table><thead><tr><th>CCI</th><th>Comprador</th><th>Etapa</th><th>Categoria</th></tr></thead>"
                  f"<tbody>{''.join(falta_rows)}</tbody></table>")

    dup_rows = []
    for cci, lst in d["duplicados"][:30]:
        for p in lst:
            dup_rows.append(f"<tr><td>{escape(cci)}</td><td>{escape(str(p['caminho']))}</td></tr>")
    if not dup_rows:
        dup_rows = ["<tr><td colspan='2' class='vazio'>Nenhum.</td></tr>"]
    dup_html = (f"<h3>Duplicados ({len(d['duplicados'])} CCIs)</h3>"
                f"<table><thead><tr><th>CCI</th><th>Arquivo</th></tr></thead>"
                f"<tbody>{''.join(dup_rows)}</tbody></table>")

    div_html = tabela_problemas(
        "Comprador Divergente", d["divergentes"], ["cci", "comprador_pasta", "esperado"])

    cont = d["planilha_cont"]
    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>Natela Analytics - Checagem Final</title>
<style>
  * {{box-sizing:border-box;margin:0;padding:0}}
  body {{font-family:-apple-system,"Segoe UI",Roboto,sans-serif;background:#f4f6fa;color:#1f2937;padding:24px}}
  .container{{max-width:1300px;margin:0 auto}}
  header{{background:linear-gradient(135deg,#0f172a,#1e3a8a 60%,#3b82f6);color:white;padding:28px 32px;border-radius:16px;box-shadow:0 10px 30px rgba(15,23,42,.25);margin-bottom:24px}}
  .brand{{display:flex;align-items:center;gap:16px}}
  .logo{{width:56px;height:56px;background:white;color:#1e3a8a;border-radius:14px;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:20px}}
  header h1{{font-size:26px}}
  header .sub{{font-size:14px;opacity:.85;margin-top:4px}}
  header .meta{{font-size:13px;opacity:.75;margin-top:12px}}
  .veredito{{background:{cor_veredito};color:white;padding:24px;border-radius:16px;margin-bottom:24px;font-size:24px;font-weight:800;text-align:center;letter-spacing:.5px;box-shadow:0 6px 20px rgba(0,0,0,.12)}}
  .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:22px}}
  .card{{background:white;border-radius:12px;padding:16px 18px;box-shadow:0 2px 8px rgba(0,0,0,.04);border-left:4px solid #6366f1}}
  .card .label{{font-size:11px;color:#6b7280;text-transform:uppercase;font-weight:600;letter-spacing:.5px}}
  .card .val{{font-size:28px;font-weight:800;color:#111827;margin-top:4px}}
  .card.ok{{border-left-color:#22c55e}}
  .card.ok .val{{color:#16a34a}}
  .card.warn{{border-left-color:#eab308}}
  .card.warn .val{{color:#ca8a04}}
  .card.err{{border-left-color:#dc2626}}
  .card.err .val{{color:#b91c1c}}
  .panel{{background:white;border-radius:12px;padding:22px;box-shadow:0 2px 8px rgba(0,0,0,.04);margin-bottom:20px}}
  .panel h2{{font-size:17px;font-weight:700;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid #e5e7eb}}
  h3{{font-size:14px;color:#374151;margin:18px 0 8px;font-weight:700}}
  table{{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:10px}}
  th,td{{padding:8px 10px;text-align:left;border-bottom:1px solid #f3f4f6}}
  th{{background:#f9fafb;font-size:11px;color:#374151;text-transform:uppercase;font-weight:700}}
  tr:hover td{{background:#fafbfc}}
  .num{{text-align:right;font-weight:600;font-variant-numeric:tabular-nums}}
  .vazio{{text-align:center;color:#9ca3af;font-style:italic;padding:16px}}
  footer{{text-align:center;color:#9ca3af;font-size:12px;margin-top:24px;padding-top:18px}}
  footer strong{{color:#1e3a8a}}
</style></head><body><div class="container">
<header>
  <div class="brand">
    <div class="logo">NATELA</div>
    <div>
      <h1>Natela Analytics</h1>
      <div class="sub">Checagem Final pra Upload na Nuvem</div>
    </div>
  </div>
  <div class="meta">Gerado em {d['agora']} &middot; Pasta: <code>{escape(d['pasta'])}</code> &middot; Planilha: {escape(d['planilha'])}</div>
</header>

<div class="veredito">{txt_veredito}</div>

<div class="cards">
  <div class="card"><div class="label">PDFs no disco</div><div class="val">{d['total_pdfs']:,}</div></div>
  <div class="card ok"><div class="label">Baixado</div><div class="val">{cont.get('baixado',0):,}</div></div>
  <div class="card ok"><div class="label">Enviado</div><div class="val">{cont.get('enviado',0):,}</div></div>
  <div class="card"><div class="label">Sem Debito</div><div class="val">{cont.get('sem_debito',0):,}</div></div>
  <div class="card warn"><div class="label">Pendente</div><div class="val">{cont.get('pendente',0):,}</div></div>
  <div class="card {'err' if d['total_problemas']>0 else 'ok'}"><div class="label">Problemas</div><div class="val">{d['total_problemas']:,}</div></div>
</div>

<div class="panel">
  <h2>Cobertura por Etapa</h2>
  <table>
    <thead><tr><th>Etapa</th><th class="num">No disco / Esperado</th><th class="num">%</th></tr></thead>
    <tbody>{cob_html}</tbody>
  </table>
</div>

<div class="panel">
  <h2>Itens Problematicos</h2>
  {corrompidos_html}
  {nome_html}
  {orfaos_html}
  {falta_html}
  {dup_html}
  {div_html}
</div>

<footer>
  <strong>Natela Analytics</strong> &middot; Checagem Final IPTU 2026 &middot;
  Gerado em {d['agora']}
</footer>
</div></body></html>
"""


if __name__ == "__main__":
    main()
