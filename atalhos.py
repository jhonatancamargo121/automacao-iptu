# -*- coding: utf-8 -*-
"""
ATALHOS DOS IPTUs - Natela Analytics

HTML enxuto, focado APENAS em atalhos. 1 card por comprador, com botoes
clicaveis que abrem direto a pasta no OneDrive corporativo.

Recursos:
  - Busca em tempo real (digite parte do nome no topo)
  - Filtros: tudo / so enviados / so pendentes / com PDF / sem PDF
  - Badge ENVIADO verde nos ja-enviados
  - Botao por etapa com cor distinta se tem PDF (azul) ou nao (amarelo)

Uso:
  python atalhos.py
  python atalhos.py --abrir
"""
import argparse
import json
import re
import sys
import time
import webbrowser
from collections import defaultdict
from datetime import datetime
from html import escape
from pathlib import Path
from urllib.parse import quote

import openpyxl

NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]
SAIDA_HTML = "atalhos_iptu.html"

COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13

# URL canonica do OneDrive corporativo
URL_ONEDRIVE_BASE = (
    "https://exemplo-my.sharepoint.com/personal/"
    "usuario_exemplo_com_br/_layouts/15/onedrive.aspx"
)
PASTA_RAIZ_ONEDRIVE = (
    "/personal/usuario_exemplo_com_br/Documents/IPTUs_Baixados"
)

# Pastas locais para mapear datas e cruzar
PASTAS_LOCAIS = [
    "IPTUs_Final_Consolidado",
    "IPTUs_Baixados_29_06_2026",
    "IPTUs_Baixados_29_05_2026_RETRY",
    "IPTUs_Baixados",
]


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


def gerar_url(etapa_str, nome_pasta):
    caminho = f"{PASTA_RAIZ_ONEDRIVE}/{etapa_str}/{nome_pasta}"
    return f"{URL_ONEDRIVE_BASE}?id={quote(caminho, safe='/')}"


def indexar_pdfs_locais():
    """{cci: True} pra cada CCI com PDF achado em qualquer pasta local."""
    com_pdf = set()
    for p_nome in PASTAS_LOCAIS:
        p = Path(p_nome)
        if not p.exists():
            continue
        for pdf in p.rglob("*.pdf"):
            m = re.search(r"CCI(\d+)", pdf.name, re.I)
            if m:
                com_pdf.add(m.group(1))
    return com_pdf


def main():
    parser = argparse.ArgumentParser(description="Gera HTML de atalhos pros IPTUs")
    parser.add_argument("--abrir", action="store_true", help="Abre o HTML apos gerar")
    parser.add_argument("--saida", type=str, default=SAIDA_HTML)
    args = parser.parse_args()

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)
    print(f"Lendo planilha: {planilha}")

    wb = None
    for _ in range(8):
        try:
            wb = openpyxl.load_workbook(planilha, read_only=True)
            break
        except Exception:
            time.sleep(2)
    if wb is None:
        print("[ERRO] Planilha sendo escrita.")
        sys.exit(1)
    ws = wb.active

    print("Indexando PDFs locais...")
    ccis_com_pdf = indexar_pdfs_locais()
    print(f"  CCIs com PDF local: {len(ccis_com_pdf)}")

    # Estrutura: por_comprador[nome] = {nome_pasta, etapas: {etapa_n: [(cci, qd, lt, enviado, tem_pdf)]}, qtd_enviado, qtd_pendente}
    por_comprador = defaultdict(lambda: {
        "nome_pasta": "",
        "etapas": defaultdict(list),
        "qtd_enviado": 0,
        "qtd_pendente": 0,
        "qtd_com_pdf": 0,
        "qtd_sem_pdf": 0,
        "status_set": set(),
    })

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        cci = row[COL_CCI]
        if not cci:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue
        comprador = str(row[COL_COMPRADOR] or "").strip() or "(sem nome)"
        quadra    = str(row[COL_QUADRA] or "").strip()
        lote      = str(row[COL_LOTE] or "").strip()
        etapa_n   = extrair_etapa(row[COL_ETAPA])
        status    = str(row[COL_STATUS] or "").strip()
        obs_low   = str(row[COL_OBS_IPTU] or "").strip().lower()
        nome_pasta = limpar_nome(comprador)

        enviado = "guia enviada e-mail cliente" in obs_low
        tem_pdf = cci_str in ccis_com_pdf

        d = por_comprador[comprador]
        d["nome_pasta"] = nome_pasta
        if status:
            d["status_set"].add(status)
        if etapa_n:
            d["etapas"][etapa_n].append({
                "cci": cci_str, "quadra": quadra, "lote": lote,
                "enviado": enviado, "tem_pdf": tem_pdf,
            })
        if enviado:
            d["qtd_enviado"] += 1
        else:
            d["qtd_pendente"] += 1
        if tem_pdf:
            d["qtd_com_pdf"] += 1
        else:
            d["qtd_sem_pdf"] += 1
    wb.close()

    print(f"Compradores: {len(por_comprador)}")

    # Reagrupa: por_etapa[etapa_n] = lista de (comprador, lotes_dessa_etapa, dados_globais_comprador)
    por_etapa = defaultdict(list)
    for comprador, d in por_comprador.items():
        for etapa_n, lotes in d["etapas"].items():
            por_etapa[etapa_n].append((comprador, lotes, d))

    # Ordena etapas e dentro de cada uma, compradores em ordem alfabetica
    etapas_ordenadas = sorted(por_etapa.keys(),
                              key=lambda e: int(e) if e.isdigit() else 999)

    secoes_html = []
    for etapa_n in etapas_ordenadas:
        registros = sorted(por_etapa[etapa_n], key=lambda x: x[0].upper())
        cards_secao = []
        # Conta cada categoria nessa etapa
        n_enviado = n_pendente = n_com_pdf = n_sem_pdf = 0
        for comprador, lotes, d in registros:
            nome_pasta = d["nome_pasta"]
            todos_enviados = all(l["enviado"] for l in lotes)
            algum_enviado  = any(l["enviado"] for l in lotes)
            tem_pdf_etapa  = any(l["tem_pdf"] for l in lotes)
            qtd = len(lotes)

            classe = "btn-ok" if tem_pdf_etapa else "btn-pendente"
            if todos_enviados:
                classe += " btn-enviado"
            url = gerar_url(f"ETAPA{etapa_n}", nome_pasta)
            badge_btn = f" ({qtd})" if qtd > 1 else ""
            botao = (
                f'<a class="btn-etapa {classe}" href="{escape(url)}" target="_blank" '
                f'title="Abre ETAPA{etapa_n}/{escape(nome_pasta)} no OneDrive">'
                f'ETAPA{etapa_n}{badge_btn} <span class="arrow">&#x2197;</span></a>'
            )

            classes_card = ["card"]
            if todos_enviados:
                classes_card.append("card-enviado")
                n_enviado += 1
            else:
                classes_card.append("card-pendente")
                n_pendente += 1
            if tem_pdf_etapa:
                classes_card.append("card-com-pdf")
                n_com_pdf += 1
            else:
                classes_card.append("card-sem-pdf")
                n_sem_pdf += 1
            classe_str = " ".join(classes_card)
            busca = (comprador + " " + nome_pasta).upper()

            badge_top = ""
            if todos_enviados:
                badge_top = '<span class="badge badge-enviado">ENVIADO</span>'
            elif algum_enviado:
                qtd_env = sum(1 for l in lotes if l["enviado"])
                badge_top = f'<span class="badge badge-parcial">{qtd_env}/{qtd} ENVIADO</span>'

            # Detalhes dos lotes da etapa
            detalhes_lotes = ", ".join(
                f"CCI {l['cci']} (Qd{l['quadra']} Lt{l['lote']})"
                for l in lotes
            )

            cards_secao.append(
                f'<div class="{classe_str}" data-busca="{escape(busca)}" data-etapa="{etapa_n}">'
                f'<div class="card-head">'
                f'<div class="card-nome">'
                f'<div class="nome-real">{escape(comprador)}</div>'
                f'<code class="nome-pasta">{escape(nome_pasta)}</code>'
                f'</div>'
                f'<div class="card-badges">{badge_top}</div>'
                f'</div>'
                f'<div class="card-botoes">{botao}</div>'
                f'<div class="card-lotes">{escape(detalhes_lotes)}</div>'
                f'</div>'
            )

        secao = (
            f'<section class="secao-etapa" data-etapa="{etapa_n}">'
            f'<div class="secao-header">'
            f'<h2>ETAPA {etapa_n}</h2>'
            f'<div class="secao-stats">'
            f'<span class="stat-pill stat-total">{len(registros)} compradores</span>'
            f'<span class="stat-pill stat-enviado">{n_enviado} enviados</span>'
            f'<span class="stat-pill stat-pendente">{n_pendente} pendentes</span>'
            f'<span class="stat-pill stat-com-pdf">{n_com_pdf} c/PDF</span>'
            f'</div></div>'
            f'<div class="cards-grid">{"".join(cards_secao)}</div>'
            f'</section>'
        )
        secoes_html.append(secao)

    cards_html = "\n".join(secoes_html) if secoes_html else "<div class='vazio'>Nenhum comprador.</div>"
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    total_comp = len(por_comprador)
    total_enviado = sum(1 for d in por_comprador.values()
                       if d["qtd_enviado"] > 0 and d["qtd_pendente"] == 0)
    total_pendente = sum(1 for d in por_comprador.values() if d["qtd_pendente"] > 0)
    total_com_pdf = sum(1 for d in por_comprador.values() if d["qtd_com_pdf"] > 0)
    total_cards = sum(len(v) for v in por_etapa.values())

    html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Natela Analytics - Atalhos IPTUs</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, "Segoe UI", Roboto, sans-serif;
         background: #f4f6fa; color: #1f2937; padding: 20px; }}
  .container {{ max-width: 1400px; margin: 0 auto; }}
  header {{
    background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #3b82f6 100%);
    color: white; padding: 24px 28px; border-radius: 14px;
    box-shadow: 0 8px 24px rgba(15,23,42,.25); margin-bottom: 20px;
  }}
  .brand {{ display: flex; align-items: center; gap: 14px; }}
  .logo {{ width: 50px; height: 50px; background: white; color: #1e3a8a;
           border-radius: 12px; display:flex; align-items:center; justify-content:center;
           font-weight: 900; font-size: 18px; }}
  header h1 {{ font-size: 22px; }}
  header .sub {{ font-size: 13px; opacity: .85; margin-top: 3px; }}
  header .meta {{ font-size: 12px; opacity: .75; margin-top: 10px; }}

  .toolbar {{
    background: white; border-radius: 12px; padding: 16px 18px;
    box-shadow: 0 2px 8px rgba(0,0,0,.05); margin-bottom: 16px;
    display: flex; flex-wrap: wrap; gap: 12px; align-items: center;
  }}
  .busca-input {{
    flex: 1; min-width: 240px;
    padding: 10px 14px; border: 2px solid #e5e7eb; border-radius: 8px;
    font-size: 14px; font-family: inherit;
    transition: border-color 0.15s;
  }}
  .busca-input:focus {{ outline: none; border-color: #3b82f6; }}
  .filtros {{ display: flex; gap: 6px; flex-wrap: wrap; }}
  .filtro {{
    padding: 8px 14px; border-radius: 7px; border: 1px solid #d1d5db;
    background: white; color: #374151; cursor: pointer; font-size: 13px;
    font-weight: 600; transition: all 0.15s; font-family: inherit;
  }}
  .filtro:hover {{ border-color: #3b82f6; color: #3b82f6; }}
  .filtro.ativo {{
    background: #1e3a8a; color: white; border-color: #1e3a8a;
  }}
  .contador {{ font-size: 12px; color: #6b7280; margin-left: 8px; }}
  #total-visivel {{ font-weight: 700; color: #1e3a8a; }}

  .secao-etapa {{ margin-bottom: 28px; }}
  .secao-header {{
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 12px; padding: 14px 18px;
    background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
    border-radius: 10px; color: white;
    box-shadow: 0 2px 8px rgba(30,58,138,0.2);
  }}
  .secao-header h2 {{ font-size: 18px; font-weight: 800; letter-spacing: 0.5px; }}
  .secao-stats {{ display: flex; gap: 8px; flex-wrap: wrap; }}
  .stat-pill {{
    background: rgba(255,255,255,0.2); padding: 4px 10px; border-radius: 12px;
    font-size: 11px; font-weight: 700;
  }}
  .cards-grid {{
    display: grid; grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
    gap: 12px;
  }}
  .card-lotes {{
    font-size: 11px; color: #6b7280; margin-top: 8px;
    padding-top: 8px; border-top: 1px solid #f3f4f6;
    font-family: Consolas, monospace;
  }}
  .card {{
    background: white; border-radius: 12px; padding: 16px 18px;
    box-shadow: 0 2px 6px rgba(0,0,0,.04);
    border-left: 4px solid #d1d5db;
    transition: all 0.15s;
  }}
  .card:hover {{ box-shadow: 0 6px 16px rgba(0,0,0,.08); transform: translateY(-1px); }}
  .card.card-enviado {{ border-left-color: #22c55e; }}
  .card.card-pendente.card-com-pdf {{ border-left-color: #3b82f6; }}
  .card.card-pendente.card-sem-pdf {{ border-left-color: #eab308; }}
  .card-head {{
    display: flex; justify-content: space-between; align-items: flex-start;
    margin-bottom: 12px; gap: 10px;
  }}
  .card-nome {{ flex: 1; }}
  .nome-real {{ font-size: 15px; font-weight: 700; color: #111827; }}
  .nome-pasta {{
    font-size: 10px; color: #6b7280; background: #f3f4f6;
    padding: 1px 5px; border-radius: 3px; display: inline-block;
    margin-top: 3px; font-family: Consolas, monospace;
  }}
  .card-badges {{ display: flex; flex-direction: column; gap: 4px; align-items: flex-end; }}
  .badge {{
    display: inline-block; font-size: 10px; font-weight: 700;
    padding: 2px 7px; border-radius: 4px; letter-spacing: 0.3px;
    white-space: nowrap;
  }}
  .badge-enviado {{ background: #16a34a; color: white; }}
  .badge-parcial {{ background: #f59e0b; color: white; }}
  .badge-status {{ background: #e0e7ff; color: #3730a3; }}

  .card-botoes {{ display: flex; flex-wrap: wrap; gap: 6px; }}
  .btn-etapa {{
    text-decoration: none;
    display: inline-flex; align-items: center; gap: 4px;
    padding: 7px 12px; border-radius: 6px; font-size: 12px;
    font-weight: 700; transition: all 0.15s;
    border: 1px solid;
  }}
  .btn-etapa.btn-ok {{
    background: #eff6ff; border-color: #bfdbfe; color: #1e3a8a;
  }}
  .btn-etapa.btn-ok:hover {{
    background: #3b82f6; color: white; border-color: #2563eb;
  }}
  .btn-etapa.btn-pendente {{
    background: #fef3c7; border-color: #fcd34d; color: #92400e;
  }}
  .btn-etapa.btn-pendente:hover {{
    background: #f59e0b; color: white; border-color: #d97706;
  }}
  .btn-etapa.btn-enviado {{
    background: #dcfce7 !important; border-color: #86efac !important; color: #166534 !important;
  }}
  .btn-etapa.btn-enviado:hover {{
    background: #16a34a !important; color: white !important;
  }}
  .arrow {{ font-size: 13px; }}

  .vazio {{
    grid-column: 1/-1; text-align: center; color: #9ca3af;
    padding: 40px; font-style: italic;
  }}

  footer {{ margin-top: 28px; text-align: center; color: #9ca3af;
            font-size: 12px; padding-top: 18px; }}
  footer strong {{ color: #1e3a8a; }}
</style>
</head><body><div class="container">

<header>
  <div class="brand">
    <div class="logo">NATELA</div>
    <div>
      <h1>Natela Analytics</h1>
      <div class="sub">Atalhos IPTUs - {total_comp:,} compradores</div>
    </div>
  </div>
  <div class="meta">Gerado em {agora} &middot; Click no botao da etapa abre direto a pasta no OneDrive corporativo</div>
</header>

<div class="toolbar">
  <input type="text" id="busca" class="busca-input"
         placeholder="Buscar por nome do comprador..." autofocus>
  <div class="filtros">
    <button class="filtro ativo" data-filtro="todos">Todos</button>
    <button class="filtro" data-filtro="enviado">Enviados</button>
    <button class="filtro" data-filtro="pendente">Pendentes</button>
    <button class="filtro" data-filtro="com-pdf">Com PDF</button>
    <button class="filtro" data-filtro="sem-pdf">Sem PDF</button>
  </div>
  <span class="contador">Mostrando <span id="total-visivel">{total_comp}</span> de {total_comp}</span>
</div>

<div class="cards-grid" id="grid">
{cards_html}
</div>

<footer>
  <strong>Natela Analytics</strong> &middot; {total_enviado} ja enviados &middot;
  {total_pendente} pendentes &middot; {total_com_pdf} tem PDF no disco
</footer>

</div>

<script>
const inputBusca = document.getElementById('busca');
const filtroBtns = document.querySelectorAll('.filtro');
const cards = document.querySelectorAll('.card');
const contador = document.getElementById('total-visivel');
let filtroAtivo = 'todos';

function aplicarFiltros() {{
    const termo = inputBusca.value.trim().toUpperCase();
    let visiveis = 0;
    const cardsPorEtapa = {{}};
    cards.forEach(card => {{
        const matchBusca = !termo || (card.dataset.busca || '').includes(termo);
        let matchFiltro = true;
        if (filtroAtivo === 'enviado')  matchFiltro = card.classList.contains('card-enviado');
        if (filtroAtivo === 'pendente') matchFiltro = card.classList.contains('card-pendente');
        if (filtroAtivo === 'com-pdf')  matchFiltro = card.classList.contains('card-com-pdf');
        if (filtroAtivo === 'sem-pdf')  matchFiltro = card.classList.contains('card-sem-pdf');
        const visivel = matchBusca && matchFiltro;
        card.style.display = visivel ? '' : 'none';
        if (visivel) {{
            visiveis++;
            const et = card.dataset.etapa || '?';
            cardsPorEtapa[et] = (cardsPorEtapa[et] || 0) + 1;
        }}
    }});
    // Esconde secoes que ficaram vazias
    document.querySelectorAll('.secao-etapa').forEach(s => {{
        const et = s.dataset.etapa;
        s.style.display = (cardsPorEtapa[et] || 0) > 0 ? '' : 'none';
    }});
    contador.textContent = visiveis;
}}

inputBusca.addEventListener('input', aplicarFiltros);
filtroBtns.forEach(btn => {{
    btn.addEventListener('click', () => {{
        filtroBtns.forEach(b => b.classList.remove('ativo'));
        btn.classList.add('ativo');
        filtroAtivo = btn.dataset.filtro;
        aplicarFiltros();
    }});
}});
</script>
</body></html>
"""

    saida = Path(args.saida)
    saida.write_text(html, encoding="utf-8")
    print(f"\n[OK] Gerado: {saida.resolve()}")
    print(f"     Tamanho: {saida.stat().st_size / 1024:.1f} KB")
    print(f"     Cards: {total_comp}")
    print(f"     Ja enviados (todos os lotes): {total_enviado}")
    print(f"     Com pendentes               : {total_pendente}")
    print(f"     Tem PDF no disco            : {total_com_pdf}")

    if args.abrir:
        webbrowser.open(saida.resolve().as_uri())
        print("     Aberto no navegador.")


if __name__ == "__main__":
    main()
