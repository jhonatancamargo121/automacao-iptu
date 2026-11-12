# -*- coding: utf-8 -*-
"""
RELATORIO COMPLETO DO PROJETO - Natela Analytics

Gera um HTML rico que documenta TUDO que foi feito no projeto de automacao
do IPTU 2026 - Cidade Exemplo.

Conteudo:
  - Resumo executivo (cards)
  - Cronologia das fases
  - Metricas de eficiencia (tempo manual vs automatizado)
  - Status atual (por etapa, top compradores)
  - Lista de iniciativas (scripts) implementadas
  - Decisoes tecnicas chave
  - Proximos passos

Saida: relatorio_completo_natela_analytics.html

Uso:
  python relatorio_completo.py
  python relatorio_completo.py --abrir
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
PASTAS_PDFS = [
    "IPTUs_Baixados",
    "IPTUs_Baixados_29_06_2026",
    "IPTUs_Baixados_29_05_2026_RETRY",
]
PASTAS_PROVAS = ["Provas_Envio", "Provas_Envio_Manual"]
SAIDA_HTML = "relatorio_completo_natela_analytics.html"

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


def extrair_etapa(s):
    m = re.search(r"(\d+)", str(s or ""))
    return m.group(1) if m else ""


def coletar():
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
        print("[ERRO] planilha sendo escrita por outro processo.")
        sys.exit(1)
    ws = wb.active

    status_geral = Counter()
    por_etapa = defaultdict(lambda: Counter())
    top_compradores = Counter()
    compradores_distintos = set()
    total_vendido = 0

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        if str(row[COL_STATUS] or "").strip().lower() != "vendido":
            continue
        if not row[COL_CCI]:
            continue
        total_vendido += 1
        obs = str(row[COL_OBS_IPTU] or "").strip().lower()
        etapa = extrair_etapa(row[COL_ETAPA])
        etapa_str = f"ETAPA{etapa}" if etapa else "?"
        comprador = str(row[COL_COMPRADOR] or "").strip()
        if comprador:
            compradores_distintos.add(comprador)

        if "guia enviada e-mail cliente" in obs:
            cat = "enviado"
        elif "sem_debito" in obs or "sem debito" in obs:
            cat = "sem_debito"
        elif "indisponivel" in obs or "boleto_indisponivel" in obs:
            cat = "indisponivel"
        elif "pdf baixado" in obs:
            cat = "baixado"
        else:
            cat = "pendente"

        status_geral[cat] += 1
        por_etapa[etapa_str][cat] += 1
        if cat in ("baixado", "enviado") and comprador:
            top_compradores[comprador] += 1
    wb.close()

    # Vasculha as pastas
    pastas_info = []
    total_pdfs_disco = 0
    tam_total = 0
    for pasta_nome in PASTAS_PDFS:
        p = Path(pasta_nome)
        if not p.exists():
            pastas_info.append({"nome": pasta_nome, "existe": False, "pdfs": 0, "mb": 0.0})
            continue
        cnt = 0
        tam = 0
        for pdf in p.rglob("*.pdf"):
            cnt += 1
            try:
                tam += pdf.stat().st_size
            except Exception:
                pass
        pastas_info.append({
            "nome": pasta_nome, "existe": True,
            "pdfs": cnt, "mb": tam / (1024 * 1024),
        })
        total_pdfs_disco += cnt
        tam_total += tam

    # Provas (screenshots) salvas
    provas_info = []
    total_provas = 0
    for pasta_nome in PASTAS_PROVAS:
        p = Path(pasta_nome)
        if not p.exists():
            provas_info.append({"nome": pasta_nome, "existe": False, "arquivos": 0})
            continue
        cnt = sum(1 for _ in p.rglob("*"))
        provas_info.append({"nome": pasta_nome, "existe": True, "arquivos": cnt})
        total_provas += cnt

    # Scripts do projeto
    scripts_info = []
    for nome in (
        "baixar_iptu.py", "enviar_iptu.py", "capturar_prova.py",
        "limpar_cache.py", "reprocessar_falhas.py", "restaurar_marcadores.py",
        "vistoria_iptus.py", "analise_profunda.py",
        "gerar_dashboard.py", "relatorio_baixados.py", "relatorio_completo.py",
        "reset_completo.py",
    ):
        p = Path(nome)
        if p.exists():
            scripts_info.append({
                "nome": nome,
                "linhas": sum(1 for _ in open(p, encoding="utf-8", errors="ignore")),
                "kb": p.stat().st_size / 1024,
            })

    return {
        "planilha_path": str(planilha),
        "total_vendido": total_vendido,
        "compradores_distintos": len(compradores_distintos),
        "status_geral": dict(status_geral),
        "por_etapa": {k: dict(v) for k, v in por_etapa.items()},
        "top_compradores": top_compradores.most_common(15),
        "pastas_info": pastas_info,
        "provas_info": provas_info,
        "scripts_info": scripts_info,
        "total_pdfs_disco": total_pdfs_disco,
        "tam_total_mb": tam_total / (1024 * 1024),
        "total_provas": total_provas,
    }


SCRIPT_DESCRICOES = {
    "baixar_iptu.py":
        "Automacao principal. Le planilha, abre site da prefeitura via Playwright, "
        "preenche CCI, navega 8 modais, captura PDF do Jasper. Suporta --data, --pasta, "
        "--teste, --cci, --inicio.",
    "enviar_iptu.py":
        "Envio dos boletos via Outlook Web automatizado. Login com 2FA (sessao salva), "
        "compose, anexo via file chooser, screenshot de prova, atualiza planilha.",
    "capturar_prova.py":
        "Captura manual de tela (F9). OCR pelo Tesseract extrai quadra/lote/nome do "
        "corpo do e-mail. Salva PDF organizado por subpasta do cliente.",
    "limpar_cache.py":
        "Limpa marcadores 'PDF baixado' da planilha. Modo padrao preserva SEM_DEBITO "
        "e INDISPONIVEL (otimiza re-rodadas).",
    "reprocessar_falhas.py":
        "Le falhas.log, cruza com pastas existentes, marca CCIs ja-baixados em outra "
        "pasta, prepara lista de re-baixar. Gera comando exato pra rodar.",
    "restaurar_marcadores.py":
        "Recupera marcadores SEM_DEBITO e INDISPONIVEL a partir do falhas.log, "
        "caso o cache tenha sido limpo por engano.",
    "vistoria_iptus.py":
        "Auditoria de integridade: compara planilha vs disco, detecta PDFs faltantes "
        "ou corrompidos, opcionalmente limpa marcadores pra reprocessar.",
    "analise_profunda.py":
        "Analise forense: PDFs orfaos (no disco sem linha), duplicados, tamanho anomalo, "
        "comprador divergente, pastas vazias. Opcional: OCR de conteudo via pypdf.",
    "gerar_dashboard.py":
        "Dashboard HTML Natela Analytics: cards, donut, barras por etapa, top compradores, "
        "tabelas de problemas. Visao operacional.",
    "relatorio_baixados.py":
        "Relatorio HTML por comprador com botoes que copiam o nome e abrem direto a "
        "pasta correspondente no OneDrive corporativo.",
    "relatorio_completo.py":
        "Este relatorio que voce esta lendo agora. Sintese estrategica do projeto.",
    "reset_completo.py":
        "Move IPTUs_Baixados/ pra backup e zera markers. Permite recomecar limpo "
        "(usado quando muda data de pagamento).",
}


CRONOLOGIA = [
    {
        "fase": "Fase 1 - Diagnostico Manual",
        "duracao": "Pre-projeto",
        "desc": (
            "Levantamento da rotina manual: cada IPTU exige ~3 min de cliques no site "
            "da prefeitura (8 modais), salvar PDF, renomear, enviar e-mail. "
            "1.558 lotes pendentes = ~78 horas de trabalho manual repetitivo."
        ),
    },
    {
        "fase": "Fase 2 - Automacao do Download",
        "duracao": "Implementacao",
        "desc": (
            "Construcao do baixar_iptu.py em Playwright async. 30+ iteracoes para "
            "domar o Angular dinamico: filtro anti-widget de chat, detecao de modais, "
            "navegacao do datepicker, captura do Jasper via response listener."
        ),
    },
    {
        "fase": "Fase 3 - Robustez e Resume",
        "duracao": "Hardening",
        "desc": (
            "Suporte a interrupcao/retomada (planilha salva por lote), detecao de "
            "SEM_DEBITO e INDISPONIVEL, listener de resposta do Jasper (cookies + "
            "request.get pra capturar PDF mesmo se a aba fechar)."
        ),
    },
    {
        "fase": "Fase 4 - Velocidade",
        "duracao": "Tuning",
        "desc": (
            "Reducao de PAUSA_ENTRE_CLIQUES (3s -> 1.5s -> 2.2s otimo). "
            "Estimativa de tempo total caiu de ~21h para ~2-4h. "
            "Flags --data, --pasta para multiplas rodadas com datas diferentes."
        ),
    },
    {
        "fase": "Fase 5 - Envio de E-mails",
        "duracao": "Implementacao",
        "desc": (
            "enviar_iptu.py com Outlook Web automatizado: login 2FA persistido, "
            "compose, autocomplete handling, anexo via expect_file_chooser, "
            "screenshot de prova, atualizacao da planilha com marcador 'Guia enviada'."
        ),
    },
    {
        "fase": "Fase 6 - Observabilidade",
        "duracao": "Suite analitica",
        "desc": (
            "Vistoria de integridade, analise profunda (orfaos/duplicados), dashboards "
            "HTML Natela Analytics com Chart.js, relatorio por comprador com deep-link "
            "OneDrive."
        ),
    },
    {
        "fase": "Fase 7 - Operacao",
        "duracao": "Em curso",
        "desc": (
            "Rodadas de download com data 29/05 e 29/06. Reprocessamento inteligente "
            "(reprocessar_falhas.py cruza pastas pra nao re-baixar boletos que ja temos)."
        ),
    },
]


DECISOES_TECNICAS = [
    ("Playwright async em vez de Selenium",
     "Lida melhor com Angular SPA, melhor controle de eventos e file chooser."),
    ("Captura de PDF via response listener (nao re-fetch URL)",
     "Servidor Jasper invalida o timestamp apos 1a leitura. Listener pega o body "
     "diretamente da resposta original."),
    ("Cruzamento entre pastas no reprocessamento",
     "Se ja temos o PDF de outra rodada (29/05), nao re-baixa. Economiza ~30s por lote."),
    ("Subpasta por comprador",
     "Mesma estrutura nas pastas de download E de provas. Cliente pode ter multiplos "
     "lotes agrupados no mesmo diretorio."),
    ("Markers no Excel como fonte de verdade",
     "Status do lote e gravado na propria planilha. Resume automatico via filtro."),
    ("CLI flags em vez de constantes",
     "--data, --pasta, --cci permitem rodadas paralelas sem editar codigo."),
    ("Detecao multi-estrategia no calendario do Outlook",
     "Header de mes/ano sao identificados via 3 estrategias em cascata "
     "(button.uib-title, btn-default, role=heading)."),
    ("OCR no painel direito do Outlook (crop 45%+)",
     "Tesseract era confundido pela lista lateral de e-mails. Crop pre-OCR resolveu."),
]


METRICAS_OTIM = [
    ("Tempo manual estimado (3 min/lote x 1558)", "~78h", "100%"),
    ("Tempo automatizado (com pauses 2.2s)", "~3-4h", "~5%"),
    ("Economia operacional",                "~74h", "-95%"),
    ("Custo de manutencao do script",       "0h/lote", "-"),
]


def gerar_html(d):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    status = d["status_geral"]
    n_baixado = status.get("baixado", 0)
    n_enviado = status.get("enviado", 0)
    n_sd = status.get("sem_debito", 0)
    n_ind = status.get("indisponivel", 0)
    n_pend = status.get("pendente", 0)
    total = d["total_vendido"]
    pct_concluido = ((n_baixado + n_enviado + n_sd + n_ind) / total * 100) if total else 0

    # Donut
    donut_labels = ["Baixado", "Enviado", "Sem debito", "Indisponivel", "Pendente"]
    donut_data = [n_baixado, n_enviado, n_sd, n_ind, n_pend]

    # Barras por etapa
    etapas = sorted(d["por_etapa"].keys(),
                    key=lambda e: int(re.search(r"\d+", e).group()) if re.search(r"\d+", e) else 999)
    cores = {
        "baixado": "#22c55e", "enviado": "#3b82f6",
        "sem_debito": "#a3a3a3", "indisponivel": "#f97316", "pendente": "#eab308",
    }
    bar_datasets = [
        {"label": cat.replace("_", " ").title(),
         "data": [d["por_etapa"][e].get(cat, 0) for e in etapas],
         "backgroundColor": cores[cat]}
        for cat in ("baixado", "enviado", "sem_debito", "indisponivel", "pendente")
    ]

    # Cronologia
    cronologia_html = "".join(
        f'<div class="fase">'
        f'<div class="fase-num">{i}</div>'
        f'<div class="fase-body">'
        f'<div class="fase-titulo">{escape(f["fase"])} <span class="fase-tag">{escape(f["duracao"])}</span></div>'
        f'<div class="fase-desc">{escape(f["desc"])}</div>'
        f'</div></div>'
        for i, f in enumerate(CRONOLOGIA, 1)
    )

    # Decisoes
    decisoes_html = "".join(
        f"<tr><td><strong>{escape(t)}</strong></td><td>{escape(desc)}</td></tr>"
        for t, desc in DECISOES_TECNICAS
    )

    # Metricas
    metricas_html = "".join(
        f"<tr><td>{escape(label)}</td><td class='num'>{escape(valor)}</td><td class='num'>{escape(pct)}</td></tr>"
        for label, valor, pct in METRICAS_OTIM
    )

    # Pastas
    pastas_html = "".join(
        f"<tr><td><code>{escape(p['nome'])}</code></td>"
        f"<td>{'OK' if p['existe'] else '<span class=cinza>nao existe</span>'}</td>"
        f"<td class='num'>{p['pdfs']:,}</td>"
        f"<td class='num'>{p['mb']:.1f} MB</td></tr>"
        for p in d["pastas_info"]
    )
    provas_html = "".join(
        f"<tr><td><code>{escape(p['nome'])}</code></td>"
        f"<td>{'OK' if p['existe'] else '<span class=cinza>nao existe</span>'}</td>"
        f"<td class='num'>{p['arquivos']:,}</td></tr>"
        for p in d["provas_info"]
    )

    # Top compradores
    max_top = d["top_compradores"][0][1] if d["top_compradores"] else 1
    top_html = "".join(
        f"<tr><td>{escape(c[:50])}</td>"
        f"<td><div class='bar-wrap'><div class='bar' style='width:{(n/max_top*100):.1f}%'></div>"
        f"<span class='bar-label'>{n}</span></div></td></tr>"
        for c, n in d["top_compradores"]
    )
    if not top_html:
        top_html = "<tr><td colspan='2' class='vazio'>Nenhum dado.</td></tr>"

    # Scripts
    scripts_html = "".join(
        f"<tr><td><code>{escape(s['nome'])}</code></td>"
        f"<td>{escape(SCRIPT_DESCRICOES.get(s['nome'], ''))}</td>"
        f"<td class='num'>{s['linhas']:,}</td>"
        f"<td class='num'>{s['kb']:.1f} KB</td></tr>"
        for s in d["scripts_info"]
    )

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Natela Analytics - Relatorio Completo do Projeto</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, "Segoe UI", Roboto, sans-serif;
         background: #f4f6fa; color: #1f2937; padding: 24px; }}
  .container {{ max-width: 1400px; margin: 0 auto; }}
  header {{
    background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #3b82f6 100%);
    color: white; padding: 32px 36px; border-radius: 18px;
    box-shadow: 0 12px 36px rgba(15,23,42,0.25); margin-bottom: 28px;
  }}
  .brand {{ display: flex; align-items: center; gap: 18px; }}
  .logo {{ width: 64px; height: 64px; background: white; color: #1e3a8a;
           border-radius: 16px; display:flex; align-items:center; justify-content:center;
           font-weight: 900; font-size: 22px; box-shadow: 0 4px 14px rgba(0,0,0,0.15); }}
  header h1 {{ font-size: 30px; letter-spacing: -0.5px; }}
  header .sub {{ font-size: 15px; opacity: 0.9; margin-top: 6px; }}
  header .meta {{ font-size: 13px; opacity: 0.7; margin-top: 14px; }}
  header .tagline {{
    background: rgba(255,255,255,0.15); padding: 8px 14px; border-radius: 8px;
    display: inline-block; margin-top: 14px; font-size: 13px;
  }}
  h2 {{ font-size: 20px; font-weight: 700; margin: 32px 0 16px; color: #0f172a;
        display: flex; align-items: center; gap: 10px; }}
  h2::before {{ content: ""; width: 6px; height: 22px; background: #3b82f6; border-radius: 3px; }}
  .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 16px; margin-bottom: 24px; }}
  .card {{ background: white; border-radius: 14px; padding: 20px 22px;
           box-shadow: 0 2px 10px rgba(0,0,0,0.05); border-left: 5px solid #6366f1; }}
  .card .label {{ font-size: 11px; color: #6b7280; text-transform: uppercase;
                  font-weight: 700; letter-spacing: 0.6px; }}
  .card .val {{ font-size: 32px; font-weight: 800; color: #111827; margin-top: 6px; }}
  .card .pct {{ font-size: 12px; color: #9ca3af; margin-top: 2px; }}
  .card.success {{ border-left-color: #22c55e; }} .card.success .val {{ color: #16a34a; }}
  .card.info    {{ border-left-color: #3b82f6; }} .card.info .val    {{ color: #2563eb; }}
  .card.warn    {{ border-left-color: #eab308; }} .card.warn .val    {{ color: #ca8a04; }}
  .card.danger  {{ border-left-color: #f97316; }} .card.danger .val  {{ color: #ea580c; }}
  .card.muted   {{ border-left-color: #9ca3af; }} .card.muted .val   {{ color: #6b7280; }}
  .row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px; }}
  .row-1 {{ display: grid; grid-template-columns: 1fr; gap: 24px; margin-bottom: 24px; }}
  @media (max-width: 900px) {{ .row {{ grid-template-columns: 1fr; }} }}
  .panel {{ background: white; border-radius: 14px; padding: 24px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05); }}
  .panel-titulo {{ font-size: 17px; font-weight: 700; color: #111827;
                   margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid #e5e7eb; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th, td {{ padding: 10px 12px; text-align: left; border-bottom: 1px solid #f3f4f6;
            vertical-align: top; }}
  th {{ background: #f9fafb; font-weight: 700; color: #374151; font-size: 11px;
        text-transform: uppercase; letter-spacing: 0.4px; }}
  tr:hover td {{ background: #fafbfc; }}
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }}
  code {{ background: #f3f4f6; padding: 2px 6px; border-radius: 4px;
          font-family: Consolas, Monaco, monospace; font-size: 12px; }}
  .cinza {{ color: #9ca3af; font-style: italic; }}
  .vazio {{ text-align: center; color: #9ca3af; font-style: italic; padding: 20px; }}
  .bar-wrap {{ position: relative; height: 20px; background: #f3f4f6; border-radius: 4px; overflow: hidden; }}
  .bar {{ height: 100%; background: linear-gradient(90deg, #16a34a 0%, #22c55e 100%); }}
  .bar-label {{ position: absolute; right: 8px; top: 50%; transform: translateY(-50%);
                font-size: 11px; font-weight: 700; color: #1f2937; }}
  .fase {{ display: flex; gap: 18px; padding: 16px; border-radius: 10px;
           background: #f9fafb; margin-bottom: 12px; }}
  .fase-num {{ width: 42px; height: 42px; background: linear-gradient(135deg, #3b82f6, #1e3a8a);
               color: white; border-radius: 10px; font-weight: 800; font-size: 18px;
               display: flex; align-items: center; justify-content: center;
               flex-shrink: 0; }}
  .fase-titulo {{ font-weight: 700; font-size: 15px; color: #111827; margin-bottom: 4px; }}
  .fase-tag {{ background: #dbeafe; color: #1e40af; font-size: 11px; padding: 2px 8px;
               border-radius: 10px; margin-left: 8px; font-weight: 600; }}
  .fase-desc {{ font-size: 13px; color: #4b5563; line-height: 1.55; }}
  .chart-container {{ position: relative; height: 320px; }}
  .destaque {{ background: linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%);
               border-left: 5px solid #16a34a; padding: 18px 22px; border-radius: 12px;
               font-size: 14px; color: #14532d; margin-bottom: 16px; }}
  .destaque strong {{ color: #052e16; }}
  footer {{ margin-top: 40px; text-align: center; color: #9ca3af; font-size: 13px;
            padding: 20px 0; border-top: 1px solid #e5e7eb; }}
  footer strong {{ color: #1e3a8a; }}
</style>
</head>
<body>
<div class="container">

<header>
  <div class="brand">
    <div class="logo">NATELA</div>
    <div>
      <h1>Natela Analytics</h1>
      <div class="sub">Relatorio Completo do Projeto - Automacao IPTU 2026</div>
    </div>
  </div>
  <div class="tagline">Cidade Exemplo &middot; Natela Urbanismo</div>
  <div class="meta">Gerado em {agora} &middot; Planilha: {escape(Path(d['planilha_path']).name)}</div>
</header>

<h2>Resumo Executivo</h2>
<div class="destaque">
  <strong>Projeto entrega o equivalente a ~74h de trabalho manual em ~3-4h de execucao automatizada.</strong>
  Ate o momento, {n_baixado:,} boletos baixados, {n_sd:,} confirmados como sem debito,
  e {n_pend:,} ainda em fila. Cobertura atual: <strong>{pct_concluido:.1f}%</strong> dos {total:,} lotes Vendidos.
</div>

<div class="cards">
  <div class="card info"><div class="label">Total Vendido (com CCI)</div><div class="val">{total:,}</div></div>
  <div class="card info"><div class="label">Compradores distintos</div><div class="val">{d['compradores_distintos']:,}</div></div>
  <div class="card success"><div class="label">PDFs Baixados</div><div class="val">{n_baixado:,}</div><div class="pct">{(n_baixado/total*100 if total else 0):.1f}%</div></div>
  <div class="card info"><div class="label">E-mails Enviados</div><div class="val">{n_enviado:,}</div></div>
  <div class="card muted"><div class="label">Sem Debito</div><div class="val">{n_sd:,}</div></div>
  <div class="card danger"><div class="label">Indisponivel</div><div class="val">{n_ind:,}</div></div>
  <div class="card warn"><div class="label">Pendentes</div><div class="val">{n_pend:,}</div></div>
  <div class="card success"><div class="label">% Concluido</div><div class="val">{pct_concluido:.1f}%</div></div>
  <div class="card info"><div class="label">PDFs no Disco</div><div class="val">{d['total_pdfs_disco']:,}</div><div class="pct">{d['tam_total_mb']:.1f} MB</div></div>
  <div class="card info"><div class="label">Provas geradas</div><div class="val">{d['total_provas']:,}</div></div>
</div>

<div class="row">
  <div class="panel">
    <div class="panel-titulo">Distribuicao Global</div>
    <div class="chart-container"><canvas id="donut"></canvas></div>
  </div>
  <div class="panel">
    <div class="panel-titulo">Status por Etapa</div>
    <div class="chart-container"><canvas id="bar"></canvas></div>
  </div>
</div>

<h2>Cronologia das Iniciativas</h2>
<div class="panel">
  {cronologia_html}
</div>

<h2>Metricas de Otimizacao</h2>
<div class="panel">
  <table>
    <thead><tr><th>Metrica</th><th class="num">Valor</th><th class="num">% do manual</th></tr></thead>
    <tbody>{metricas_html}</tbody>
  </table>
</div>

<h2>Decisoes Tecnicas Chave</h2>
<div class="panel">
  <table>
    <thead><tr><th>Decisao</th><th>Motivacao / Resultado</th></tr></thead>
    <tbody>{decisoes_html}</tbody>
  </table>
</div>

<h2>Scripts do Projeto</h2>
<div class="panel">
  <table>
    <thead><tr><th>Arquivo</th><th>Funcao</th><th class="num">Linhas</th><th class="num">Tamanho</th></tr></thead>
    <tbody>{scripts_html}</tbody>
  </table>
</div>

<div class="row">
  <div class="panel">
    <div class="panel-titulo">Pastas de PDFs</div>
    <table>
      <thead><tr><th>Pasta</th><th>Status</th><th class="num">PDFs</th><th class="num">Tam</th></tr></thead>
      <tbody>{pastas_html}</tbody>
    </table>
  </div>
  <div class="panel">
    <div class="panel-titulo">Pastas de Provas</div>
    <table>
      <thead><tr><th>Pasta</th><th>Status</th><th class="num">Arquivos</th></tr></thead>
      <tbody>{provas_html}</tbody>
    </table>
  </div>
</div>

<h2>Top 15 Compradores (PDFs baixados/enviados)</h2>
<div class="panel">
  <table>
    <thead><tr><th>Comprador</th><th>Lotes</th></tr></thead>
    <tbody>{top_html}</tbody>
  </table>
</div>

<h2>Proximos Passos Sugeridos</h2>
<div class="panel">
  <ol style="margin-left:18px; font-size:14px; line-height:1.7; color:#374151;">
    <li><strong>Finalizar o download dos {n_pend:,} pendentes</strong> com data 29/06/2026
        rodando <code>python baixar_iptu.py</code>.</li>
    <li><strong>Reprocessar os boletos indisponiveis</strong> com data alternativa 29/05/2026:
        <code>python reprocessar_falhas.py --incluir-indisponivel</code> +
        <code>python baixar_iptu.py --data 29/05/2026 --pasta IPTUs_Baixados_29_05_2026_RETRY</code>.</li>
    <li><strong>Vistoria final</strong> com <code>python vistoria_iptus.py</code> e
        <code>python analise_profunda.py --com-conteudo</code>.</li>
    <li><strong>Iniciar envios por e-mail</strong> com <code>python enviar_iptu.py --teste 5</code>
        antes da rodada completa.</li>
    <li><strong>Capturar provas</strong> dos envios com <code>python capturar_prova.py</code>
        (atalho F9 abre PDF organizado por cliente).</li>
    <li><strong>Disponibilizar relatorio por comprador</strong> com
        <code>python relatorio_baixados.py --abrir</code> (botoes copiam nome e abrem
        OneDrive direto na pasta correta).</li>
  </ol>
</div>

<footer>
  <strong>Natela Analytics</strong> &middot; Sintese Completa do Projeto &middot;
  Gerado em {agora}<br>
  <small>Sistema de automacao para a Natela Urbanismo &middot; IPTU 2026 Cidade Exemplo</small>
</footer>

</div>

<script>
new Chart(document.getElementById('donut'), {{
  type: 'doughnut',
  data: {{
    labels: {json.dumps(donut_labels)},
    datasets: [{{
      data: {json.dumps(donut_data)},
      backgroundColor: ['#22c55e', '#3b82f6', '#a3a3a3', '#f97316', '#eab308'],
      borderWidth: 2, borderColor: '#fff',
    }}]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ position: 'right', labels: {{ font: {{ size: 13 }} }} }},
      title: {{ display: true, text: 'Natela Analytics', color: '#9ca3af', font: {{ size: 10 }} }}
    }}
  }}
}});

new Chart(document.getElementById('bar'), {{
  type: 'bar',
  data: {{ labels: {json.dumps(etapas)}, datasets: {json.dumps(bar_datasets)} }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    scales: {{ x: {{ stacked: true }}, y: {{ stacked: true, beginAtZero: true }} }},
    plugins: {{
      legend: {{ position: 'top', labels: {{ font: {{ size: 12 }} }} }},
      title: {{ display: true, text: 'Natela Analytics', color: '#9ca3af', font: {{ size: 10 }} }}
    }}
  }}
}});
</script>
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser(description="Relatorio completo do projeto - Natela Analytics")
    parser.add_argument("--abrir", action="store_true",
                        help="Abre o HTML no navegador apos gerar")
    parser.add_argument("--saida", type=str, default=SAIDA_HTML,
                        help=f"Arquivo de saida (padrao: {SAIDA_HTML})")
    args = parser.parse_args()

    print("Coletando dados...")
    dados = coletar()
    print("Gerando HTML...")
    html = gerar_html(dados)
    saida = Path(args.saida)
    saida.write_text(html, encoding="utf-8")
    print(f"\n[OK] Relatorio gerado: {saida.resolve()}")
    print(f"     Tamanho: {saida.stat().st_size / 1024:.1f} KB")
    if args.abrir:
        webbrowser.open(saida.resolve().as_uri())
        print("     Aberto no navegador.")


if __name__ == "__main__":
    main()
