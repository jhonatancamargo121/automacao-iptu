# -*- coding: utf-8 -*-
"""
GERA DASHBOARD HTML - Natela Analytics

Le a planilha e a pasta IPTUs_Baixados/ e gera um HTML autocontido com:
  - Cards de overview (total, baixados, sem debito, pendentes, etc.)
  - Grafico donut da distribuicao
  - Grafico de barras por etapa
  - Tabelas de orfaos, duplicados, anomalos
  - Branding "Natela Analytics" em todo lugar

Saida: dashboard_natela_analytics.html (abre no navegador padrao)

Uso:
  python gerar_dashboard.py
  python gerar_dashboard.py --abrir       (abre no navegador apos gerar)
"""
import argparse
import json
import re
import sys
import webbrowser
from collections import defaultdict, Counter
from pathlib import Path
from datetime import datetime
from html import escape

import openpyxl

NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]
PASTA_PDFS = "IPTUs_Baixados"
SAIDA_HTML = "dashboard_natela_analytics.html"

COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13

TAM_MIN = 5 * 1024
TAM_MAX = 5 * 1024 * 1024


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


def extrair_etapa(s):
    m = re.search(r"(\d+)", str(s or ""))
    return m.group(1) if m else ""


def extrair_cci_do_nome(nome):
    m = re.search(r"CCI(\d+)", nome, re.IGNORECASE)
    return m.group(1) if m else None


def coletar_dados():
    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    print(f"Lendo planilha: {planilha}")
    import time
    wb = None
    for tentativa in range(8):
        try:
            wb = openpyxl.load_workbook(planilha, read_only=True)
            break
        except Exception as e:
            print(f"  [.] tentativa {tentativa+1}/8 ({type(e).__name__}); aguardando 2s...")
            time.sleep(2)
    if wb is None:
        print("[ERRO] Nao consegui ler a planilha (baixar_iptu escrevendo?).")
        sys.exit(1)
    ws = wb.active

    por_cci = {}
    por_etapa_status = defaultdict(lambda: defaultdict(int))
    por_status_global = Counter()

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        status_raw = str(row[COL_STATUS] or "").strip()
        cci = row[COL_CCI]
        obs = str(row[COL_OBS_IPTU] or "").strip()
        if status_raw.lower() != "vendido" or not cci:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue

        etapa = extrair_etapa(row[COL_ETAPA])
        etapa_str = f"ETAPA{etapa}" if etapa else "SEM_ETAPA"
        obs_low = obs.lower()

        if "guia enviada e-mail cliente" in obs_low:
            cat = "enviado"
        elif "sem_debito" in obs_low or "sem debito" in obs_low:
            cat = "sem_debito"
        elif "indisponivel" in obs_low or "boleto_indisponivel" in obs_low:
            cat = "indisponivel"
        elif "pdf baixado" in obs_low:
            cat = "baixado"
        else:
            cat = "pendente"

        info = {
            "linha": i,
            "quadra": str(row[COL_QUADRA] or "").strip(),
            "lote": str(row[COL_LOTE] or "").strip(),
            "comprador": str(row[COL_COMPRADOR] or "").strip(),
            "etapa": etapa_str,
            "obs": obs,
            "categoria": cat,
        }
        por_cci[cci_str] = info
        por_status_global[cat] += 1
        por_etapa_status[etapa_str][cat] += 1

    wb.close()

    # Vasculha disco
    base = Path(PASTA_PDFS)
    pdfs = []
    pdfs_por_cci = defaultdict(list)
    if base.exists():
        for pdf in base.rglob("*.pdf"):
            try:
                tam = pdf.stat().st_size
            except Exception:
                tam = 0
            cci = extrair_cci_do_nome(pdf.name)
            rel = pdf.relative_to(base)
            partes = rel.parts
            etapa_pasta = partes[0] if len(partes) >= 1 else ""
            comprador_pasta = partes[1] if len(partes) >= 2 else ""
            info = {
                "caminho": str(rel),
                "nome": pdf.name,
                "cci": cci,
                "tamanho": tam,
                "etapa_pasta": etapa_pasta,
                "comprador_pasta": comprador_pasta,
            }
            pdfs.append(info)
            if cci:
                pdfs_por_cci[cci].append(info)

    # Analise
    orfaos = []
    for p in pdfs:
        if not p["cci"]:
            orfaos.append({**p, "motivo": "sem CCI no nome"})
        elif p["cci"] not in por_cci:
            orfaos.append({**p, "motivo": "CCI nao existe na planilha"})

    duplicados = []
    for cci, lst in pdfs_por_cci.items():
        if len(lst) > 1:
            duplicados.append({"cci": cci, "arquivos": [p["caminho"] for p in lst]})

    anomalos = []
    for p in pdfs:
        if p["tamanho"] < TAM_MIN:
            anomalos.append({**p, "tipo": "PEQUENO"})
        elif p["tamanho"] > TAM_MAX:
            anomalos.append({**p, "tipo": "GRANDE"})

    divergentes = []
    for p in pdfs:
        if p["cci"] and p["cci"] in por_cci and p["comprador_pasta"]:
            esperado = limpar_nome(por_cci[p["cci"]]["comprador"])
            if p["comprador_pasta"] != esperado:
                divergentes.append({**p, "esperado": esperado})

    # Tamanho medio dos PDFs validos
    tamanhos = [p["tamanho"] for p in pdfs if TAM_MIN <= p["tamanho"] <= TAM_MAX]
    tam_medio = sum(tamanhos) / len(tamanhos) if tamanhos else 0
    tam_total_mb = sum(p["tamanho"] for p in pdfs) / (1024 * 1024)

    return {
        "planilha_path": str(planilha),
        "por_cci": por_cci,
        "por_etapa_status": dict(por_etapa_status),
        "por_status_global": dict(por_status_global),
        "pdfs": pdfs,
        "pdfs_por_cci_count": {cci: len(lst) for cci, lst in pdfs_por_cci.items()},
        "orfaos": orfaos,
        "duplicados": duplicados,
        "anomalos": anomalos,
        "divergentes": divergentes,
        "tam_medio_kb": tam_medio / 1024,
        "tam_total_mb": tam_total_mb,
    }


def gerar_html(dados):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    total_lotes = len(dados["por_cci"])
    total_pdfs = len(dados["pdfs"])
    status = dados["por_status_global"]
    n_baixado = status.get("baixado", 0)
    n_enviado = status.get("enviado", 0)
    n_sem_debito = status.get("sem_debito", 0)
    n_indispon = status.get("indisponivel", 0)
    n_pendente = status.get("pendente", 0)
    pct_concluido = ((n_baixado + n_enviado + n_sem_debito + n_indispon)
                     / total_lotes * 100) if total_lotes else 0

    # Dados para os graficos
    donut_labels = ["Baixado", "Enviado", "Sem debito", "Indisponivel", "Pendente"]
    donut_data = [n_baixado, n_enviado, n_sem_debito, n_indispon, n_pendente]

    etapas_ordenadas = sorted(
        dados["por_etapa_status"].keys(),
        key=lambda e: int(re.search(r"\d+", e).group()) if re.search(r"\d+", e) else 999,
    )
    bar_categorias = ["baixado", "enviado", "sem_debito", "indisponivel", "pendente"]
    bar_datasets = []
    cores_categ = {
        "baixado":      "#22c55e",
        "enviado":      "#3b82f6",
        "sem_debito":   "#a3a3a3",
        "indisponivel": "#f97316",
        "pendente":     "#eab308",
    }
    for cat in bar_categorias:
        bar_datasets.append({
            "label": cat.replace("_", " ").title(),
            "data": [dados["por_etapa_status"][e].get(cat, 0) for e in etapas_ordenadas],
            "backgroundColor": cores_categ[cat],
        })

    # Top 15 compradores com mais PDFs baixados
    contagem_comprador = Counter()
    for cci, info in dados["por_cci"].items():
        if info["categoria"] in ("baixado", "enviado"):
            contagem_comprador[info["comprador"]] += 1
    top_compradores = contagem_comprador.most_common(15)

    # Tabelas
    def linha_tabela_orfaos(o):
        return (
            f"<tr><td>{escape(o['nome'])}</td>"
            f"<td>{escape(o['cci'] or '-')}</td>"
            f"<td>{o['tamanho']:,}</td>"
            f"<td>{escape(o['motivo'])}</td></tr>"
        )

    def linha_tabela_dup(d):
        arquivos_html = "<br>".join(escape(a) for a in d["arquivos"])
        return f"<tr><td>{escape(d['cci'])}</td><td>{arquivos_html}</td></tr>"

    def linha_tabela_anomalo(a):
        return (
            f"<tr><td>{escape(a.get('tipo', ''))}</td>"
            f"<td>{escape(a['cci'] or '-')}</td>"
            f"<td>{a['tamanho']:,}</td>"
            f"<td>{escape(a['nome'])}</td></tr>"
        )

    def linha_tabela_div(d):
        return (
            f"<tr><td>{escape(d['cci'])}</td>"
            f"<td>{escape(d['comprador_pasta'])}</td>"
            f"<td>{escape(d['esperado'])}</td></tr>"
        )

    def linha_top_comprador(c, n, max_n):
        pct = (n / max_n * 100) if max_n else 0
        return (
            f"<tr><td>{escape(c[:50])}</td>"
            f"<td><div class='bar-wrap'><div class='bar' style='width:{pct:.1f}%'></div>"
            f"<span class='bar-label'>{n}</span></div></td></tr>"
        )

    max_top = top_compradores[0][1] if top_compradores else 1

    tabela_orfaos = "".join(linha_tabela_orfaos(o) for o in dados["orfaos"][:50])
    tabela_dup = "".join(linha_tabela_dup(d) for d in dados["duplicados"][:50])
    tabela_anomalo = "".join(linha_tabela_anomalo(a) for a in dados["anomalos"][:50])
    tabela_div = "".join(linha_tabela_div(d) for d in dados["divergentes"][:50])
    tabela_top = "".join(linha_top_comprador(c, n, max_top) for c, n in top_compradores)

    if not tabela_orfaos: tabela_orfaos = "<tr><td colspan='4' class='vazio'>Nenhum orfao detectado.</td></tr>"
    if not tabela_dup:    tabela_dup = "<tr><td colspan='2' class='vazio'>Nenhum duplicado detectado.</td></tr>"
    if not tabela_anomalo: tabela_anomalo = "<tr><td colspan='4' class='vazio'>Nenhum anomalo detectado.</td></tr>"
    if not tabela_div:    tabela_div = "<tr><td colspan='3' class='vazio'>Nenhuma divergencia.</td></tr>"
    if not tabela_top:    tabela_top = "<tr><td colspan='2' class='vazio'>Nenhum dado.</td></tr>"

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Natela Analytics - Dashboard IPTU 2026</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, "Segoe UI", Roboto, sans-serif;
    background: #f4f6fa;
    color: #1f2937;
    padding: 24px;
  }}
  .container {{ max-width: 1400px; margin: 0 auto; }}
  header {{
    background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
    color: white;
    padding: 28px 32px;
    border-radius: 16px;
    box-shadow: 0 10px 30px rgba(30, 58, 138, 0.2);
    margin-bottom: 24px;
  }}
  header .brand {{
    display: flex;
    align-items: center;
    gap: 16px;
  }}
  header .logo {{
    width: 56px; height: 56px;
    background: white;
    color: #1e3a8a;
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 900; font-size: 20px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
  }}
  header h1 {{ font-size: 26px; font-weight: 700; letter-spacing: -0.5px; }}
  header .subtitle {{ font-size: 14px; opacity: 0.85; margin-top: 4px; }}
  header .meta {{ margin-top: 12px; font-size: 13px; opacity: 0.75; }}

  .grid-cards {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
  }}
  .card {{
    background: white;
    border-radius: 12px;
    padding: 18px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    border-left: 4px solid #6366f1;
  }}
  .card .label {{ font-size: 12px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }}
  .card .value {{ font-size: 30px; font-weight: 800; color: #111827; margin-top: 6px; }}
  .card .pct {{ font-size: 12px; color: #9ca3af; margin-top: 2px; }}
  .card.success {{ border-left-color: #22c55e; }}
  .card.success .value {{ color: #16a34a; }}
  .card.warn {{ border-left-color: #eab308; }}
  .card.warn .value {{ color: #ca8a04; }}
  .card.danger {{ border-left-color: #f97316; }}
  .card.danger .value {{ color: #ea580c; }}
  .card.info {{ border-left-color: #3b82f6; }}
  .card.info .value {{ color: #2563eb; }}
  .card.muted {{ border-left-color: #9ca3af; }}
  .card.muted .value {{ color: #6b7280; }}

  .row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px; }}
  .row-1col {{ display: grid; grid-template-columns: 1fr; gap: 24px; margin-bottom: 24px; }}
  @media (max-width: 900px) {{ .row {{ grid-template-columns: 1fr; }} }}
  .panel {{
    background: white;
    border-radius: 12px;
    padding: 22px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
  }}
  .panel h2 {{
    font-size: 16px; font-weight: 700; color: #111827;
    margin-bottom: 16px;
    padding-bottom: 10px; border-bottom: 1px solid #e5e7eb;
    display: flex; align-items: center; gap: 8px;
  }}
  .panel h2::before {{
    content: ""; display: inline-block;
    width: 4px; height: 16px;
    background: #3b82f6; border-radius: 2px;
  }}

  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th, td {{ padding: 8px 12px; text-align: left; border-bottom: 1px solid #f3f4f6; }}
  th {{ background: #f9fafb; font-weight: 600; color: #374151; font-size: 12px; text-transform: uppercase; letter-spacing: 0.3px; }}
  tr:hover td {{ background: #f9fafb; }}
  .vazio {{ text-align: center; color: #9ca3af; font-style: italic; padding: 20px; }}

  .bar-wrap {{
    position: relative;
    height: 18px;
    background: #f3f4f6;
    border-radius: 4px;
    overflow: hidden;
  }}
  .bar {{
    height: 100%;
    background: linear-gradient(90deg, #3b82f6 0%, #60a5fa 100%);
    border-radius: 4px;
  }}
  .bar-label {{
    position: absolute; right: 6px; top: 50%; transform: translateY(-50%);
    font-size: 11px; font-weight: 600; color: #1f2937;
  }}

  footer {{
    margin-top: 32px;
    text-align: center;
    color: #9ca3af;
    font-size: 12px;
    padding-top: 20px;
  }}
  footer strong {{ color: #1e3a8a; }}

  .chart-container {{ position: relative; height: 280px; }}
</style>
</head>
<body>
<div class="container">

<header>
  <div class="brand">
    <div class="logo">NATELA</div>
    <div>
      <h1>Natela Analytics</h1>
      <div class="subtitle">Dashboard de IPTU 2026 - Cidade Exemplo</div>
    </div>
  </div>
  <div class="meta">Relatorio gerado em {agora} | Planilha: {escape(Path(dados['planilha_path']).name)}</div>
</header>

<div class="grid-cards">
  <div class="card info">
    <div class="label">Total Vendido (com CCI)</div>
    <div class="value">{total_lotes:,}</div>
  </div>
  <div class="card success">
    <div class="label">PDFs Baixados</div>
    <div class="value">{n_baixado:,}</div>
    <div class="pct">{(n_baixado/total_lotes*100 if total_lotes else 0):.1f}%</div>
  </div>
  <div class="card info">
    <div class="label">E-mails Enviados</div>
    <div class="value">{n_enviado:,}</div>
    <div class="pct">{(n_enviado/total_lotes*100 if total_lotes else 0):.1f}%</div>
  </div>
  <div class="card muted">
    <div class="label">Sem Debito</div>
    <div class="value">{n_sem_debito:,}</div>
    <div class="pct">{(n_sem_debito/total_lotes*100 if total_lotes else 0):.1f}%</div>
  </div>
  <div class="card danger">
    <div class="label">Indisponivel</div>
    <div class="value">{n_indispon:,}</div>
  </div>
  <div class="card warn">
    <div class="label">Pendentes</div>
    <div class="value">{n_pendente:,}</div>
    <div class="pct">{(n_pendente/total_lotes*100 if total_lotes else 0):.1f}%</div>
  </div>
  <div class="card success">
    <div class="label">% Concluido</div>
    <div class="value">{pct_concluido:.1f}%</div>
  </div>
  <div class="card info">
    <div class="label">PDFs no Disco</div>
    <div class="value">{total_pdfs:,}</div>
    <div class="pct">{dados['tam_total_mb']:.1f} MB total</div>
  </div>
</div>

<div class="row">
  <div class="panel">
    <h2>Distribuicao por Status</h2>
    <div class="chart-container">
      <canvas id="donutChart"></canvas>
    </div>
  </div>
  <div class="panel">
    <h2>Status por Etapa</h2>
    <div class="chart-container">
      <canvas id="barChart"></canvas>
    </div>
  </div>
</div>

<div class="row">
  <div class="panel">
    <h2>Top 15 Compradores (PDFs baixados/enviados)</h2>
    <table>
      <thead><tr><th>Comprador</th><th>Lotes</th></tr></thead>
      <tbody>{tabela_top}</tbody>
    </table>
  </div>
  <div class="panel">
    <h2>Integridade Geral</h2>
    <table>
      <tbody>
        <tr><td>PDFs orfaos no disco</td><td><strong>{len(dados['orfaos'])}</strong></td></tr>
        <tr><td>CCIs com PDFs duplicados</td><td><strong>{len(dados['duplicados'])}</strong></td></tr>
        <tr><td>PDFs com tamanho anomalo</td><td><strong>{len(dados['anomalos'])}</strong></td></tr>
        <tr><td>PDFs em pasta divergente</td><td><strong>{len(dados['divergentes'])}</strong></td></tr>
        <tr><td>Tamanho medio dos PDFs validos</td><td>{dados['tam_medio_kb']:.1f} KB</td></tr>
        <tr><td>Armazenamento total</td><td>{dados['tam_total_mb']:.2f} MB</td></tr>
      </tbody>
    </table>
  </div>
</div>

<div class="row-1col">
  <div class="panel">
    <h2>PDFs Orfaos (no disco mas sem linha na planilha)</h2>
    <table>
      <thead><tr><th>Arquivo</th><th>CCI</th><th>Tamanho (bytes)</th><th>Motivo</th></tr></thead>
      <tbody>{tabela_orfaos}</tbody>
    </table>
  </div>
</div>

<div class="row">
  <div class="panel">
    <h2>PDFs Duplicados</h2>
    <table>
      <thead><tr><th>CCI</th><th>Arquivos</th></tr></thead>
      <tbody>{tabela_dup}</tbody>
    </table>
  </div>
  <div class="panel">
    <h2>Tamanho Anomalo</h2>
    <table>
      <thead><tr><th>Tipo</th><th>CCI</th><th>Bytes</th><th>Arquivo</th></tr></thead>
      <tbody>{tabela_anomalo}</tbody>
    </table>
  </div>
</div>

<div class="row-1col">
  <div class="panel">
    <h2>Comprador Divergente (pasta != comprador da planilha)</h2>
    <table>
      <thead><tr><th>CCI</th><th>Pasta</th><th>Esperado</th></tr></thead>
      <tbody>{tabela_div}</tbody>
    </table>
  </div>
</div>

<footer>
  <strong>Natela Analytics</strong> &middot; Sistema de Analise de IPTU 2026 &middot;
  Gerado por <strong>Natela Analytics</strong> em {agora}
</footer>

</div>

<script>
const cores = {{
  baixado: '#22c55e',
  enviado: '#3b82f6',
  sem_debito: '#a3a3a3',
  indisponivel: '#f97316',
  pendente: '#eab308',
}};

new Chart(document.getElementById('donutChart'), {{
  type: 'doughnut',
  data: {{
    labels: {json.dumps(donut_labels)},
    datasets: [{{
      data: {json.dumps(donut_data)},
      backgroundColor: ['#22c55e', '#3b82f6', '#a3a3a3', '#f97316', '#eab308'],
      borderWidth: 2,
      borderColor: '#fff',
    }}]
  }},
  options: {{
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{
      legend: {{ position: 'right', labels: {{ font: {{ size: 12 }} }} }},
      title: {{ display: true, text: 'Natela Analytics', color: '#9ca3af', font: {{ size: 10, weight: 'normal' }}, padding: {{ bottom: 10 }} }}
    }}
  }}
}});

new Chart(document.getElementById('barChart'), {{
  type: 'bar',
  data: {{
    labels: {json.dumps(etapas_ordenadas)},
    datasets: {json.dumps(bar_datasets)}
  }},
  options: {{
    responsive: true,
    maintainAspectRatio: false,
    scales: {{
      x: {{ stacked: true }},
      y: {{ stacked: true, beginAtZero: true }}
    }},
    plugins: {{
      legend: {{ position: 'top', labels: {{ font: {{ size: 11 }} }} }},
      title: {{ display: true, text: 'Natela Analytics', color: '#9ca3af', font: {{ size: 10, weight: 'normal' }}, padding: {{ bottom: 8 }} }}
    }}
  }}
}});
</script>
</body>
</html>
"""
    return html


def main():
    parser = argparse.ArgumentParser(
        description="Gera dashboard HTML Natela Analytics com analise completa"
    )
    parser.add_argument("--abrir", action="store_true",
                        help="Abre o HTML gerado no navegador padrao")
    parser.add_argument("--saida", type=str, default=SAIDA_HTML,
                        help=f"Nome do arquivo HTML de saida (padrao: {SAIDA_HTML})")
    args = parser.parse_args()

    print("Coletando dados...")
    dados = coletar_dados()
    print("Gerando HTML...")
    html = gerar_html(dados)

    saida = Path(args.saida)
    saida.write_text(html, encoding="utf-8")
    print(f"\n[OK] Dashboard gerado: {saida.resolve()}")
    print(f"     Tamanho: {saida.stat().st_size / 1024:.1f} KB")

    if args.abrir:
        webbrowser.open(saida.resolve().as_uri())
        print("     Aberto no navegador padrao.")
    else:
        print(f"\nPara abrir: clique 2x em '{saida.name}' ou rode:")
        print(f"   python gerar_dashboard.py --abrir")


if __name__ == "__main__":
    main()
