# -*- coding: utf-8 -*-
"""
RELATORIO DE IPTUs BAIXADOS - Natela Analytics

Le a planilha e gera um relatorio agrupado por COMPRADOR mostrando:
  - Quantos PDFs ja baixados
  - Quantos sem debito
  - Quantos indisponiveis
  - Quantos ainda pendentes
  - Detalhe de cada lote (CCI, etapa, quadra/lote, status)

Saidas:
  - Terminal (resumo)
  - CSV detalhado
  - HTML estilizado (Natela Analytics)

Uso:
  python relatorio_baixados.py
      -> terminal + CSV + HTML
  python relatorio_baixados.py --abrir
      -> tambem abre o HTML no navegador
  python relatorio_baixados.py --so-baixados
      -> filtra so quem tem PDF baixado (exclui sem_debito / pendente)
"""
import argparse
import re
import sys
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

COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13

# Mapeamento das pastas de PDF -> data de pagamento que o boleto foi gerado.
# As pastas listadas tem prioridade nessa ordem. Alem destas, qualquer
# outra pasta IPTUs_*/ tambem eh scaneada automaticamente.
PASTAS_E_DATAS = [
    ("IPTUs_Baixados_29_06_2026",       "29/06/2026"),
    ("IPTUs_Baixados_29_05_2026_RETRY", "29/05/2026"),
    ("IPTUs_Baixados",                  "29/05/2026"),
    ("IPTUs_Final_Consolidado",         ""),    # data sera deduzida via prioridade
    ("IPTUs_Ja_Enviados_Arquivo",       ""),    # idem - vem do consolidado
]

# URL canonica do SharePoint - abre DIRETO na subpasta do comprador.
# Click no botao -> nova aba abre em ETAPA{N}/{COMPRADOR}/ com os PDFs do cara.
URL_ONEDRIVE_BASE = (
    "https://exemplo-my.sharepoint.com/personal/"
    "usuario_exemplo_com_br/_layouts/15/onedrive.aspx"
)
PASTA_RAIZ_ONEDRIVE = (
    "/personal/usuario_exemplo_com_br/Documents/IPTUs_Baixados"
)


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


def extrair_etapa(s):
    m = re.search(r"(\d+)", str(s or ""))
    return m.group(1) if m else ""


def limpar_nome_pasta(nome):
    """Normaliza o nome do comprador no MESMO formato usado em IPTUs_Baixados/."""
    nome = (nome or "").upper().strip()
    nome = re.sub(r'[<>:"/\\|?*]', "", nome)
    nome = re.sub(r"\s+", "_", nome)
    return nome[:80] or "SEM_NOME"


def gerar_url_onedrive(etapa_str, comprador_pasta):
    """Monta URL canonica do SharePoint que abre direto a subpasta do comprador."""
    from urllib.parse import quote
    caminho = f"{PASTA_RAIZ_ONEDRIVE}/{etapa_str}/{comprador_pasta}"
    return f"{URL_ONEDRIVE_BASE}?id={quote(caminho, safe='/')}"


def _deduzir_data_de_nome(pasta_nome):
    """Detecta data no nome da pasta. Ex: '_29_06_2026' -> '29/06/2026'."""
    m = re.search(r"(\d{2})_(\d{2})_(\d{4})", pasta_nome)
    if m:
        return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
    return ""


def mapear_ccis_por_data():
    """
    Varre TODAS as pastas IPTUs_*/ (configuradas + auto-descobertas)
    e devolve dict {cci: 'DD/MM/YYYY'}.
    Primeira pasta que tem o CCI vence (ordem da lista PASTAS_E_DATAS).
    """
    # Comeca com a lista configurada (ordem de prioridade)
    pastas_visitadas = set()
    pastas_em_ordem = []
    for nome, data in PASTAS_E_DATAS:
        pastas_em_ordem.append((nome, data))
        pastas_visitadas.add(nome.lower())

    # Auto-descobre quaisquer outras pastas IPTUs_*/
    for d in sorted(Path(".").iterdir()):
        if not d.is_dir():
            continue
        nome = d.name
        if nome.lower() in pastas_visitadas:
            continue
        if nome.lower().startswith("iptus_") or nome.lower().startswith("iptu_"):
            pastas_em_ordem.append((nome, _deduzir_data_de_nome(nome)))
            pastas_visitadas.add(nome.lower())

    mapa = {}
    contagem = {}
    for pasta_nome, data in pastas_em_ordem:
        p = Path(pasta_nome)
        if not p.exists():
            continue
        cnt = 0
        for pdf in p.rglob("*.pdf"):
            m = re.search(r"CCI(\d+)", pdf.name, re.I)
            if not m:
                continue
            cci = m.group(1)
            cnt += 1
            if cci not in mapa:
                mapa[cci] = data or _deduzir_data_de_nome(pasta_nome) or "?"
        contagem[pasta_nome] = cnt
    return mapa, contagem


def classificar(obs_low):
    if "guia enviada e-mail cliente" in obs_low:
        return "enviado"
    if "sem_debito" in obs_low or "sem debito" in obs_low:
        return "sem_debito"
    if "indisponivel" in obs_low or "boleto_indisponivel" in obs_low:
        return "indisponivel"
    if "pdf baixado" in obs_low:
        return "baixado"
    return "pendente"


def main():
    parser = argparse.ArgumentParser(
        description="Relatorio de IPTUs baixados por comprador - Natela Analytics"
    )
    parser.add_argument("--abrir", action="store_true",
                        help="Abre o HTML no navegador padrao apos gerar")
    parser.add_argument("--so-baixados", action="store_true",
                        help="Inclui apenas compradores com pelo menos 1 PDF baixado/enviado")
    args = parser.parse_args()

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    print(f"Lendo planilha: {planilha}")
    # Retry: a planilha pode estar sendo salva pelo baixar_iptu nesse exato momento
    import time
    wb = None
    for tentativa in range(8):
        try:
            wb = openpyxl.load_workbook(planilha, read_only=True)
            break
        except Exception as e:
            print(f"  [.] tentativa {tentativa+1}/8 falhou ({type(e).__name__}); "
                  f"aguardando 2s...")
            time.sleep(2)
    if wb is None:
        print("[ERRO] Nao consegui ler a planilha apos 8 tentativas. "
              "Ela esta sendo escrita constantemente?")
        sys.exit(1)
    ws = wb.active

    # Mapeamento CCI -> data de pagamento (varrendo as pastas dos PDFs)
    print("Mapeando datas dos PDFs (29/05 vs 29/06)...")
    mapa_data_por_cci, contagem_por_pasta = mapear_ccis_por_data()
    print(f"  CCIs com data identificada: {len(mapa_data_por_cci)}")
    for pasta, cnt in contagem_por_pasta.items():
        if cnt > 0:
            print(f"    {pasta:<40} {cnt:>5} PDFs")

    # Estrutura: por_comprador[nome] = {baixado, enviado, sem_debito, indisponivel, pendente, lotes: []}
    por_comprador = defaultdict(lambda: {
        "baixado": 0, "enviado": 0, "sem_debito": 0,
        "indisponivel": 0, "pendente": 0, "lotes": [],
    })

    total_geral = defaultdict(int)
    status_global = Counter()
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        status = str(row[COL_STATUS] or "").strip()
        cci = row[COL_CCI]
        # Aceita TODOS os status, exige apenas que tenha CCI
        if not cci:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue

        comprador = str(row[COL_COMPRADOR] or "").strip() or "(sem nome)"
        quadra    = str(row[COL_QUADRA] or "").strip()
        lote      = str(row[COL_LOTE] or "").strip()
        etapa     = extrair_etapa(row[COL_ETAPA])
        obs       = str(row[COL_OBS_IPTU] or "").strip()
        cat       = classificar(obs.lower())

        status_global[status or "(sem status)"] += 1
        por_comprador[comprador][cat] += 1
        por_comprador[comprador].setdefault("status_set", set()).add(status or "?")
        por_comprador[comprador]["lotes"].append({
            "cci": cci_str,
            "quadra": quadra,
            "lote": lote,
            "etapa": etapa,
            "categoria": cat,
            "status": status or "?",
            "obs": obs[:80],
            "data_pagamento": mapa_data_por_cci.get(cci_str, ""),
        })
        total_geral[cat] += 1
        total_geral["total"] += 1

    wb.close()

    # Filtra so baixados se pedido
    if args.so_baixados:
        por_comprador = {
            c: d for c, d in por_comprador.items()
            if d["baixado"] + d["enviado"] > 0
        }

    # Ordena por quantidade total de baixados (decrescente)
    ordem = sorted(
        por_comprador.items(),
        key=lambda kv: -(kv[1]["baixado"] + kv[1]["enviado"]),
    )

    # ---------- Terminal ----------
    print("\n" + "=" * 72)
    print("NATELA ANALYTICS - RELATORIO DE BAIXADOS POR COMPRADOR")
    print("=" * 72)
    print(f"  Total de lotes (com CCI) : {total_geral['total']}")
    print(f"  PDFs Baixados            : {total_geral['baixado']}")
    print(f"  E-mails Enviados         : {total_geral['enviado']}")
    print(f"  Sem Debito               : {total_geral['sem_debito']}")
    print(f"  Indisponivel             : {total_geral['indisponivel']}")
    print(f"  Pendentes                : {total_geral['pendente']}")
    print(f"  Compradores no filtro    : {len(por_comprador)}")
    print()
    print("DISTRIBUICAO POR STATUS DO LOTE:")
    for st, cnt in status_global.most_common():
        print(f"  {st:<30} {cnt:>5}")
    print("=" * 72)

    print(f"\n{'#':<4} {'COMPRADOR':<40} {'BAIX':>5} {'ENV':>4} {'S/D':>4} {'IND':>4} {'PEND':>5}")
    print("-" * 72)
    for idx, (comprador, d) in enumerate(ordem[:30], 1):
        print(f"{idx:<4} {comprador[:40]:<40} "
              f"{d['baixado']:>5} {d['enviado']:>4} "
              f"{d['sem_debito']:>4} {d['indisponivel']:>4} "
              f"{d['pendente']:>5}")
    if len(ordem) > 30:
        print(f"... e mais {len(ordem) - 30} compradores no CSV/HTML.")
    print("=" * 72)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # ---------- CSV ----------
    csv_path = Path(f"relatorio_baixados_{ts}.csv")
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write("comprador;total_lotes;baixados;enviados;sem_debito;indisponivel;pendentes;ccis_baixados\n")
        for comprador, d in ordem:
            total = d["baixado"] + d["enviado"] + d["sem_debito"] + d["indisponivel"] + d["pendente"]
            ccis_b = ";".join(
                l["cci"] for l in d["lotes"] if l["categoria"] in ("baixado", "enviado")
            )
            f.write(f"{comprador};{total};{d['baixado']};{d['enviado']};"
                    f"{d['sem_debito']};{d['indisponivel']};{d['pendente']};{ccis_b}\n")
    print(f"\n[OK] CSV: {csv_path.resolve()}")

    # ---------- HTML ----------
    html_path = Path(f"relatorio_baixados_{ts}.html")
    html_path.write_text(gerar_html(total_geral, ordem, ts), encoding="utf-8")
    print(f"[OK] HTML: {html_path.resolve()}")

    if args.abrir:
        webbrowser.open(html_path.resolve().as_uri())
        print("Aberto no navegador.")


def gerar_html(total_geral, ordem, ts):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    rows_compradores = []
    for idx, (comprador, d) in enumerate(ordem, 1):
        total = (d["baixado"] + d["enviado"] + d["sem_debito"]
                 + d["indisponivel"] + d["pendente"])
        nome_pasta = limpar_nome_pasta(comprador)

        # Etapas onde o comprador tem QUALQUER lote (baixado/enviado/pendente/etc).
        # Mostra botao pra TODAS, mesmo pendentes - usuario pode navegar
        # pra conferir a pasta no OneDrive.
        etapas_com_lotes = sorted(set(
            l["etapa"] for l in d["lotes"] if l["etapa"]
        ), key=lambda e: int(e) if e.isdigit() else 999)

        # Quais etapas tem PDF baixado/enviado de fato (pra estilizar)
        etapas_com_pdf = set(
            l["etapa"] for l in d["lotes"]
            if l["categoria"] in ("baixado", "enviado") and l["etapa"]
        )

        # Botoes: 1 por etapa onde o comprador tem lote (independente de PDF).
        # Click: abre DIRETO a subpasta ETAPA{N}/COMPRADOR/ no OneDrive.
        botoes = []
        for etapa_n in etapas_com_lotes:
            etapa_str = f"ETAPA{etapa_n}"
            url = gerar_url_onedrive(etapa_str, nome_pasta)
            classe_extra = "" if etapa_n in etapas_com_pdf else " btn-pendente"
            tit = (f"Abre ETAPA{etapa_n}/{nome_pasta} no OneDrive"
                   if etapa_n in etapas_com_pdf else
                   f"ETAPA{etapa_n}/{nome_pasta} (pendente - pasta pode nao existir ainda)")
            botoes.append(
                f"<button class='btn-pasta{classe_extra}' "
                f"data-url=\"{escape(url)}\" "
                f"title=\"{escape(tit)}\">"
                f"ETAPA{etapa_n} <span class='arrow'>&#x2197;</span>"
                f"</button>"
            )
        if not botoes:
            botoes_html = "<span class='sem-pasta'>(sem etapa)</span>"
        else:
            botoes_html = "<div class='pastas-wrap'>" + " ".join(botoes) + "</div>"

        # CCIs detalhados (com a data de pagamento de cada um)
        ccis_baixados = [l for l in d["lotes"] if l["categoria"] in ("baixado", "enviado")]

        def _fmt_cci(l):
            dt = l.get("data_pagamento", "")
            badge_data = ""
            if dt == "29/05/2026":
                badge_data = "<span class='badge-data badge-29-05'>29/05</span> "
            elif dt == "29/06/2026":
                badge_data = "<span class='badge-data badge-29-06'>29/06</span> "
            else:
                badge_data = "<span class='badge-data badge-sem'>?</span> "
            # Badge extra: enviado
            badge_enviado = ""
            if l.get("categoria") == "enviado":
                badge_enviado = "<span class='badge-enviado'>ENVIADO</span> "
            return (f"{badge_enviado}{badge_data}CCI {escape(l['cci'])} "
                    f"(Et{escape(l['etapa'])} Qd{escape(l['quadra'])} Lt{escape(l['lote'])})")

        ccis_html_list = [_fmt_cci(l) for l in ccis_baixados[:20]]
        ccis_html_raw = ", ".join(ccis_html_list)
        if len(ccis_baixados) > 20:
            ccis_html_raw += f" ... +{len(ccis_baixados) - 20}"

        # Resumo por data (X com 29/05, Y com 29/06)
        ccis_29_05 = sum(1 for l in ccis_baixados if l.get("data_pagamento") == "29/05/2026")
        ccis_29_06 = sum(1 for l in ccis_baixados if l.get("data_pagamento") == "29/06/2026")

        # bloco antigo de ccis_html sera substituido abaixo
        ccis_html = ", ".join(
            f"CCI {l['cci']} (Et{l['etapa']} Qd{l['quadra']} Lt{l['lote']})"
            for l in ccis_baixados[:20]
        )
        if len(ccis_baixados) > 20:
            ccis_html += f" ... +{len(ccis_baixados) - 20}"

        # Resumo de datas para a coluna "Datas"
        partes_data = []
        if ccis_29_05:
            partes_data.append(f"<span class='badge-data badge-29-05'>{ccis_29_05}x 29/05</span>")
        if ccis_29_06:
            partes_data.append(f"<span class='badge-data badge-29-06'>{ccis_29_06}x 29/06</span>")
        datas_resumo = " ".join(partes_data) if partes_data else "<span class='sem-pasta'>-</span>"

        # Status do(s) lote(s) desse comprador
        status_lista = sorted(d.get("status_set", set()))
        status_badges = " ".join(
            f"<span class='badge-status'>{escape(s)}</span>" for s in status_lista
        ) or "<span class='sem-pasta'>?</span>"

        rows_compradores.append(
            f"<tr>"
            f"<td>{idx}</td>"
            f"<td><div class='cmp-celula'>"
            f"<strong>{escape(comprador)}</strong>"
            f"<code class='cmp-pasta-code'>{escape(nome_pasta)}</code>"
            f"</div></td>"
            f"<td>{status_badges}</td>"
            f"<td>{botoes_html}</td>"
            f"<td>{datas_resumo}</td>"
            f"<td class='num total'>{total}</td>"
            f"<td class='num baix'>{d['baixado']}</td>"
            f"<td class='num env'>{d['enviado']}</td>"
            f"<td class='num sd'>{d['sem_debito']}</td>"
            f"<td class='num ind'>{d['indisponivel']}</td>"
            f"<td class='num pend'>{d['pendente']}</td>"
            f"<td class='ccis'>{ccis_html_raw if ccis_html_raw else '-'}</td>"
            f"</tr>"
        )
    rows_html = "\n".join(rows_compradores) or "<tr><td colspan='12' class='vazio'>Nenhum dado.</td></tr>"

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Natela Analytics - Relatorio de Baixados por Comprador</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, "Segoe UI", Roboto, sans-serif;
         background: #f4f6fa; color: #1f2937; padding: 24px; }}
  .container {{ max-width: 1400px; margin: 0 auto; }}
  header {{
    background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
    color: white; padding: 28px 32px; border-radius: 16px;
    box-shadow: 0 10px 30px rgba(30,58,138,0.2); margin-bottom: 24px;
  }}
  .brand {{ display: flex; align-items: center; gap: 16px; }}
  .logo {{
    width: 56px; height: 56px; background: white; color: #1e3a8a;
    border-radius: 14px; display:flex; align-items:center; justify-content:center;
    font-weight: 900; font-size: 20px;
  }}
  header h1 {{ font-size: 26px; }}
  header .sub {{ font-size: 14px; opacity: 0.85; margin-top: 4px; }}
  header .meta {{ font-size: 13px; opacity: 0.75; margin-top: 12px; }}
  .cards {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 14px; margin-bottom: 22px;
  }}
  .card {{ background: white; border-radius: 12px; padding: 16px 18px;
           box-shadow: 0 2px 8px rgba(0,0,0,0.04); border-left: 4px solid #6366f1; }}
  .card .label {{ font-size: 11px; color: #6b7280; text-transform: uppercase;
                  font-weight: 600; letter-spacing: 0.5px; }}
  .card .val {{ font-size: 28px; font-weight: 800; color: #111827; margin-top: 4px; }}
  .card.b {{ border-left-color: #22c55e; }} .card.b .val {{ color: #16a34a; }}
  .card.e {{ border-left-color: #3b82f6; }} .card.e .val {{ color: #2563eb; }}
  .card.s {{ border-left-color: #9ca3af; }} .card.s .val {{ color: #6b7280; }}
  .card.i {{ border-left-color: #f97316; }} .card.i .val {{ color: #ea580c; }}
  .card.p {{ border-left-color: #eab308; }} .card.p .val {{ color: #ca8a04; }}
  .panel {{ background: white; border-radius: 12px; padding: 22px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04); }}
  .panel h2 {{ font-size: 16px; font-weight: 700; margin-bottom: 14px;
              padding-bottom: 10px; border-bottom: 1px solid #e5e7eb;
              display: flex; align-items: center; gap: 8px; }}
  .panel h2::before {{ content: ""; display: inline-block;
    width: 4px; height: 16px; background: #3b82f6; border-radius: 2px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th, td {{ padding: 8px 10px; text-align: left; border-bottom: 1px solid #f3f4f6; }}
  th {{ background: #f9fafb; font-size: 11px; color: #374151;
        text-transform: uppercase; letter-spacing: 0.3px; font-weight: 700; }}
  tr:hover td {{ background: #f9fafb; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }}
  .num.total {{ color: #111827; }}
  .num.baix {{ color: #16a34a; }}
  .num.env  {{ color: #2563eb; }}
  .num.sd   {{ color: #6b7280; }}
  .num.ind  {{ color: #ea580c; }}
  .num.pend {{ color: #ca8a04; }}
  .ccis {{ color: #6b7280; font-size: 11px; max-width: 380px; }}
  .vazio {{ text-align: center; color: #9ca3af; padding: 20px; font-style: italic; }}
  footer {{ margin-top: 32px; text-align: center; color: #9ca3af;
            font-size: 12px; padding-top: 20px; }}
  footer strong {{ color: #1e3a8a; }}
  .cmp-celula {{ display: flex; flex-direction: column; gap: 2px; }}
  .cmp-pasta-code {{
    font-size: 10px; font-family: Consolas, monospace;
    color: #6b7280; background: #f3f4f6; padding: 1px 5px;
    border-radius: 3px; display: inline-block; max-width: fit-content;
  }}
  .pastas-wrap {{ display: flex; flex-wrap: wrap; gap: 4px; }}
  .btn-pasta {{
    background: #eff6ff; border: 1px solid #bfdbfe; color: #1e3a8a;
    padding: 4px 10px; border-radius: 5px; cursor: pointer;
    font-family: inherit; font-size: 12px; font-weight: 600;
    transition: all 0.15s ease; white-space: nowrap;
  }}
  .btn-pasta:hover {{
    background: #3b82f6; color: white; border-color: #2563eb;
    transform: translateY(-1px);
    box-shadow: 0 2px 6px rgba(59, 130, 246, 0.3);
  }}
  .btn-pasta:active {{ transform: translateY(0); }}
  .btn-pasta .arrow {{ font-size: 14px; margin-left: 2px; }}
  .btn-pasta.copiado {{
    background: #22c55e; color: white; border-color: #16a34a;
  }}
  .btn-pasta.copiado::after {{ content: " OK"; }}
  /* Botao pra comprador SEM PDF baixado ainda (pendente).
     Visual mais discreto pra distinguir dos com PDF. */
  .btn-pasta.btn-pendente {{
    background: #fef3c7; border-color: #fcd34d; color: #92400e;
  }}
  .btn-pasta.btn-pendente:hover {{
    background: #f59e0b; color: white; border-color: #d97706;
  }}
  .sem-pasta {{ color: #9ca3af; font-size: 11px; font-style: italic; }}
  .badge-data {{
    display: inline-block; font-size: 10px; font-weight: 700;
    padding: 2px 6px; border-radius: 4px; margin-right: 3px;
    font-family: Consolas, monospace; letter-spacing: 0.3px;
  }}
  .badge-29-05 {{ background: #fef3c7; color: #92400e; }}
  .badge-29-06 {{ background: #dbeafe; color: #1e40af; }}
  .badge-sem   {{ background: #f3f4f6; color: #6b7280; }}
  .badge-enviado {{
    display: inline-block; background: #16a34a; color: white;
    font-size: 9px; font-weight: 800; padding: 2px 6px;
    border-radius: 4px; margin-right: 3px; letter-spacing: 0.5px;
  }}
  .badge-status {{
    display: inline-block; background: #e0e7ff; color: #3730a3;
    font-size: 10px; font-weight: 700; padding: 2px 6px;
    border-radius: 4px; margin-right: 3px;
  }}
  .info-box {{
    background: #fef3c7; border-left: 4px solid #f59e0b;
    padding: 12px 16px; border-radius: 8px; margin-bottom: 16px;
    font-size: 13px; color: #78350f;
  }}
  .info-box strong {{ color: #92400e; }}
  .info-box code {{
    background: #fde68a; padding: 1px 5px; border-radius: 3px;
    font-family: Consolas, monospace; font-size: 12px;
  }}
</style>
</head>
<body>
<div class="container">

<header>
  <div class="brand">
    <div class="logo">NATELA</div>
    <div>
      <h1>Natela Analytics</h1>
      <div class="sub">Relatorio de Baixados por Comprador - IPTU 2026</div>
    </div>
  </div>
  <div class="meta">Gerado em {agora}</div>
</header>

<div class="cards">
  <div class="card"><div class="label">Total Vendido</div><div class="val">{total_geral['total']:,}</div></div>
  <div class="card b"><div class="label">Baixados</div><div class="val">{total_geral['baixado']:,}</div></div>
  <div class="card e"><div class="label">Enviados</div><div class="val">{total_geral['enviado']:,}</div></div>
  <div class="card s"><div class="label">Sem Debito</div><div class="val">{total_geral['sem_debito']:,}</div></div>
  <div class="card i"><div class="label">Indisponivel</div><div class="val">{total_geral['indisponivel']:,}</div></div>
  <div class="card p"><div class="label">Pendentes</div><div class="val">{total_geral['pendente']:,}</div></div>
  <div class="card b"><div class="label">Compradores</div><div class="val">{len(ordem):,}</div></div>
</div>

<div class="panel">
  <h2>Compradores ordenados por PDFs baixados</h2>
  <div class="info-box">
    <strong>Localizador de IPTUs no OneDrive.</strong>
    Click no botao <code>ETAPA{{N}}&#x2197;</code> abre <strong>direto a pasta
    do comprador</strong> no OneDrive em nova aba - voce baixa o PDF dali.
    <br><br>
    <strong>Badge verde "ENVIADO"</strong> indica que o IPTU ja foi enviado por
    e-mail (nao precisa enviar de novo).
  </div>
  <table>
    <thead>
      <tr>
        <th>#</th>
        <th>Comprador</th>
        <th>Status</th>
        <th>Pastas</th>
        <th>Datas</th>
        <th class="num">Total</th>
        <th class="num">Baix</th>
        <th class="num">Env</th>
        <th class="num">S/D</th>
        <th class="num">Ind</th>
        <th class="num">Pend</th>
        <th>CCIs baixados (Etapa/Qd/Lt) - <span class="badge-data badge-29-05">29/05</span>/<span class="badge-data badge-29-06">29/06</span></th>
      </tr>
    </thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>

<footer>
  <strong>Natela Analytics</strong> &middot; Relatorio de IPTU 2026 &middot;
  Gerado por <strong>Natela Analytics</strong> em {agora}
</footer>

</div>

<script>
document.querySelectorAll('.btn-pasta').forEach(function(btn) {{
    btn.addEventListener('click', function(ev) {{
        ev.preventDefault();
        const url = btn.getAttribute('data-url') || '';
        if (!url) return;
        btn.classList.add('copiado');
        setTimeout(function() {{ btn.classList.remove('copiado'); }}, 1500);
        window.open(url, '_blank', 'noopener');
    }});
}});
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
