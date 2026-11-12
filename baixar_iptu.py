# -*- coding: utf-8 -*-
"""
AUTOMACAO - DOWNLOAD DE IPTUs (CIDADE EXEMPLO)

Reescrita limpa. Sem acentos no codigo (escapes \\uXXXX onde necessario).
Framework: Playwright async. Navegador visivel por padrao.

Uso:
  python baixar_iptu.py --debug          -> inspeciona campos da pagina e sai
  python baixar_iptu.py --teste 1        -> baixa apenas 1 lote (smoke)
  python baixar_iptu.py --teste 5        -> baixa 5 lotes
  python baixar_iptu.py                  -> baixa todos pendentes (headed)
  python baixar_iptu.py --headless       -> roda sem janela
  python baixar_iptu.py --inicio 500     -> retoma a partir da linha 500
"""

import asyncio
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

# ----------------------------- CONSTANTES -----------------------------------

NOMES_PLANILHA = [
    "IPTU - CIDADE EXEMPLO (2).xlsx",
    "IPTU - CIDADE EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
    "IPTU_-_CIDADE_EXEMPLO.xlsx",
]

PASTA_SAIDA       = "IPTUs_Baixados_29_06_2026"
URL_SITE          = "https://sig.prefeitura-exemplo.gov.br/sig/app.html#/servicosonline/debito-imovel"
DATA_PAGAMENTO    = "29/06/2026"
PAUSA_ENTRE_LOTES   = 1.5
PAUSA_ENTRE_CLIQUES = 2.2   # entre 1.5 (mais rapido) e 3 (mais seguro)
ARQ_LOG_FALHAS      = "falhas.log"

# Colunas (0-indexed)
COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13

JA_ENVIADO = "guia enviada e-mail cliente"

# Tudo que indica que o input NAO e o campo CCI (chat de suporte, login etc.)
TERMOS_PROIBIDOS = ("usuario", "chat", "atendimento", "support",
                    "senha", "login", "nomeusuario")

# Termos que sugerem que o input E o campo CCI
TERMOS_POSITIVOS = ("codigo", "imovel", "cci", "matricula", "inscricao")


# ----------------------------- UTILITARIOS ----------------------------------

def achar_planilha():
    script_dir = Path(__file__).resolve().parent
    for base in (Path.cwd(), script_dir):
        for nome in NOMES_PLANILHA:
            p = base / nome
            if p.exists():
                return p
    # fallback: qualquer xlsx com "iptu" no nome
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


def extrair_etapa(etapa_raw):
    if not etapa_raw:
        return ""
    m = re.search(r"(\d+)", str(etapa_raw))
    return m.group(1) if m else ""


def montar_nome_arquivo(lote):
    etapa = extrair_etapa(lote["etapa"])
    prefixo = f"ETAPA{etapa}_" if etapa else ""
    return (
        f"{prefixo}QD{lote['quadra']}_LT{lote['lote']}_"
        f"CCI{lote['cci']}_{limpar_nome(lote['comprador'])}.pdf"
    )


def montar_caminho(lote, pasta_saida):
    etapa = extrair_etapa(lote["etapa"])
    pasta_etapa = f"ETAPA{etapa}" if etapa else "SEM_ETAPA"
    subpasta = pasta_saida / pasta_etapa / limpar_nome(lote["comprador"])
    return subpasta / montar_nome_arquivo(lote)


def carregar_pendentes(planilha_path):
    print(f"Lendo planilha: {planilha_path}")
    wb = openpyxl.load_workbook(planilha_path)
    ws = wb.active
    pendentes = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        status = str(row[COL_STATUS] or "").strip().lower()
        cci    = row[COL_CCI]
        obs    = str(row[COL_OBS_IPTU] or "").strip().lower()
        if status != "vendido":
            continue
        if not cci:
            continue
        if JA_ENVIADO in obs:
            continue
        # Ja processado em rodada anterior
        if ("pdf baixado" in obs or "sem_debito" in obs
                or "sem_unica" in obs or "indisponivel" in obs):
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue
        pendentes.append({
            "linha":     i,
            "quadra":    str(row[COL_QUADRA]    or "").strip(),
            "lote":      str(row[COL_LOTE]      or "").strip(),
            "comprador": str(row[COL_COMPRADOR] or "").strip(),
            "cci":       cci_str,
            "etapa":     str(row[COL_ETAPA]     or "").strip(),
        })
    print(f"[OK] {len(pendentes)} lotes pendentes encontrados.")
    return pendentes, wb, ws


def marcar_baixado(ws, linha, nome_arq):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    ws.cell(row=linha, column=COL_OBS_IPTU + 1).value = f"PDF baixado {agora} - {nome_arq}"


def registrar_falha(cci, motivo):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ARQ_LOG_FALHAS, "a", encoding="utf-8") as f:
        f.write(f"{ts}\tCCI={cci}\t{motivo}\n")


def salvar_planilha(wb, planilha_path):
    try:
        wb.save(planilha_path)
        return True
    except PermissionError:
        print("  [AVISO] Planilha aberta no Excel. Feche para salvar progresso.")
        return False


# ----------------------------- PLAYWRIGHT JS ---------------------------------

JS_LISTAR_INPUTS = """
() => {
    function textoProximo(el) {
        // Procura por label associado, label irmao, ou texto de container proximo
        if (el.id) {
            const lbl = document.querySelector('label[for="' + el.id + '"]');
            if (lbl && lbl.innerText) return lbl.innerText.trim();
        }
        // Label dentro do mesmo wrapper (pd-input, .form-group, etc.)
        const wrap = el.closest('.pd-input, .form-group, .col, .field, div');
        if (wrap) {
            const lbl = wrap.querySelector('label');
            if (lbl && lbl.innerText) return lbl.innerText.trim();
        }
        // Caption Prodata: pd-caption-input, span.titulo, irmaos com texto
        const sib = el.previousElementSibling;
        if (sib && sib.innerText) return sib.innerText.trim().slice(0, 60);
        return '';
    }
    return Array.from(document.querySelectorAll('input')).map((el, i) => {
        const rect = el.getBoundingClientRect();
        const parent = el.closest('[ng-controller], .servico-online, form, .modal, .pd-form') || el.parentElement;
        return {
            idx: i,
            ngModel: el.getAttribute('ng-model') || '',
            ngClick: el.getAttribute('ng-click') || '',
            placeholder: el.placeholder || '',
            id: el.id || '',
            name: el.name || '',
            type: el.type || '',
            visible: el.offsetParent !== null && rect.width > 0 && rect.height > 0,
            x: Math.round(rect.x), y: Math.round(rect.y),
            w: Math.round(rect.width), h: Math.round(rect.height),
            disabled: el.disabled,
            parentClasses: (parent && parent.className) || '',
            parentNgCtrl: (parent && parent.getAttribute && (parent.getAttribute('ng-controller') || '')) || '',
            label: textoProximo(el),
            inPdForm: !!el.closest('.pd-form'),
        };
    });
}
"""

JS_TEM_CAMPO_VALIDO = """
() => {
    const PROIB = ['usuario','chat','senha','login','atendimento','nomeusuario'];
    return Array.from(document.querySelectorAll('.pd-form input, form input')).some(el => {
        const ng = (el.getAttribute('ng-model') || '').toLowerCase();
        if (ng === 'colfilter.term') return false;
        const ph = (el.placeholder || '').toLowerCase();
        const id = (el.id || '').toLowerCase();
        const blob = ng + ' ' + ph + ' ' + id;
        if (PROIB.some(t => blob.includes(t))) return false;
        const r = el.getBoundingClientRect();
        return (el.type === 'text' || el.type === '')
            && r.width > 50 && r.height > 10
            && el.offsetParent !== null
            && !el.disabled;
    });
}
"""


# ----------------------------- LOGICA DE PAGINA ------------------------------

async def esperar_pagina_pronta(page, timeout=25000):
    try:
        await page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        pass
    await page.wait_for_function(JS_TEM_CAMPO_VALIDO, timeout=timeout)


async def localizar_campo_cci(page):
    """
    Retorna (Locator, info) do campo CCI. Estrategia em camadas:
      1) input dentro de .pd-form que NAO seja filtro de coluna (colFilter.term)
         nem campo de chat. Entre os candidatos, prefere o que tem label
         contendo 'codigo' ou 'imovel'; senao o de menor X (mais a esquerda).
      2) fallback: primeiro input visivel sem termos proibidos.
    """
    inputs_info = await page.evaluate(JS_LISTAR_INPUTS)

    def blob(i):
        return " ".join([
            i["ngModel"], i["placeholder"], i["id"], i["name"],
            i["parentClasses"], i["parentNgCtrl"], i.get("label", ""),
        ]).lower()

    def proibido(i):
        return any(t in blob(i) for t in TERMOS_PROIBIDOS)

    def is_filtro_tabela(i):
        return i["ngModel"].lower() == "colfilter.term"

    def positivo(i):
        return any(t in blob(i) for t in TERMOS_POSITIVOS)

    # Camada 1: dentro do pd-form, fora dos filtros de coluna
    no_form = [
        i for i in inputs_info
        if i["visible"] and not i["disabled"]
        and i["type"] in ("", "text")
        and not proibido(i)
        and not is_filtro_tabela(i)
        and i.get("inPdForm", False)
    ]
    if no_form:
        # Prefere por label/positivo; senao o de menor X (mais a esquerda)
        com_match = [i for i in no_form if positivo(i)]
        if com_match:
            com_match.sort(key=lambda i: (i["y"], i["x"]))
            escolhido = com_match[0]
        else:
            no_form.sort(key=lambda i: (i["y"], i["x"]))
            escolhido = no_form[0]
        locator = page.locator("input").nth(escolhido["idx"])
        return locator, escolhido

    # Camada 2: fallback amplo
    candidatos = [
        i for i in inputs_info
        if i["visible"] and not i["disabled"]
        and i["type"] in ("", "text")
        and not proibido(i)
        and not is_filtro_tabela(i)
    ]
    if not candidatos:
        raise RuntimeError("nenhum input candidato apos filtros")
    com_match = [i for i in candidatos if positivo(i)]
    candidatos.sort(key=lambda i: (i["y"], i["x"]))
    escolhido = (sorted(com_match, key=lambda i: (i["y"], i["x"]))[0]
                 if com_match else candidatos[0])
    locator = page.locator("input").nth(escolhido["idx"])
    return locator, escolhido


class BoletoIndisponivel(Exception):
    """O Jasper retornou pagina de erro (sem dados de boleto). Pular o lote."""
    pass


class SemDebito(Exception):
    """O CCI nao tem debitos pendentes (modal de Debitos vazia). Pular o lote."""
    pass


class SemUnica(Exception):
    """O CCI tem debitos mas nao tem opcao de parcela Unica. Pular o lote."""
    pass


async def capturar_pdf_da_nova_aba(page, nova_aba, timeout=25000):
    """
    A nova aba do Jasper carrega um PDF. Pegamos a URL e baixamos via request
    do contexto (compartilha cookies). Funciona em modo headed e headless.
    Se a aba carregar a tela de erro "Error creating barcode", levanta
    BoletoIndisponivel para o chamador marcar o lote como skipped.
    """
    try:
        await nova_aba.wait_for_load_state("domcontentloaded", timeout=timeout)
    except Exception:
        pass

    url_pdf = nova_aba.url
    for _ in range(20):
        if url_pdf and "about:blank" not in url_pdf:
            break
        await asyncio.sleep(0.5)
        url_pdf = nova_aba.url

    if not url_pdf or "about:blank" in url_pdf:
        raise RuntimeError("URL do PDF nao detectada na nova aba")

    # Detecta a tela de erro do Jasper antes de tentar baixar como PDF
    try:
        titulo = await nova_aba.title()
        if _eh_erro_barcode(titulo):
            raise BoletoIndisponivel(f"Jasper retornou erro: {titulo}")
        texto_pagina = await nova_aba.evaluate("() => document.body && document.body.innerText || ''")
        if _eh_erro_barcode(texto_pagina):
            raise BoletoIndisponivel("Jasper retornou tela 'Error creating barcode'")
    except BoletoIndisponivel:
        raise
    except Exception:
        pass

    response = await page.context.request.get(url_pdf)
    if not response.ok:
        raise RuntimeError(f"HTTP {response.status} ao baixar PDF")

    data = await response.body()
    content_type = (response.headers.get("content-type") or "").lower()
    if b"%PDF" not in data[:1024] and "pdf" not in content_type:
        # Pode ser a tela de erro retornada como HTML
        try:
            texto = data[:4096].decode("utf-8", errors="ignore")
            if _eh_erro_barcode(texto):
                raise BoletoIndisponivel("Jasper retornou HTML de erro em vez de PDF")
        except BoletoIndisponivel:
            raise
        except Exception:
            pass
        raise RuntimeError(f"Resposta nao e PDF (content-type={content_type})")
    return data


async def _diag_screenshot(page, tag):
    """Salva screenshot para diagnostico quando algo falha."""
    try:
        ts = datetime.now().strftime("%H%M%S")
        nome = f"diag_{ts}_{tag}.png"
        await page.screenshot(path=nome, full_page=True)
        print(f"  [DIAG] screenshot salvo: {nome}")
    except Exception:
        pass


async def _mouse_click_no_centro(page, locator):
    """Clica no centro do elemento via mouse.move + click (dispara mousedown/up)."""
    box = await locator.bounding_box()
    if not box:
        raise RuntimeError("elemento sem bounding box")
    x = box["x"] + box["width"] / 2
    y = box["y"] + box["height"] / 2
    await page.mouse.move(x, y)
    await page.mouse.down()
    await asyncio.sleep(0.1)
    await page.mouse.up()


def _eh_erro_barcode(html_ou_texto):
    """Detecta a pagina de erro do Jasper quando o boleto nao tem dados."""
    if not html_ou_texto:
        return False
    s = html_ou_texto.lower() if isinstance(html_ou_texto, str) else ""
    return ("error creating barcode" in s
            or "data to encode cannot be empty" in s
            or "erro ao gerar relat" in s)


async def baixar_lote(page, lote, pasta_saida):
    """Baixa o boleto de IPTU de 1 lote.

    Fluxo baseado nos seletores capturados via Playwright Codegen:
      1. Fill CCI -> click Pesquisar
      2. Dblclick na linha do resultado (habilita Exibir Debitos)
      3. Click Exibir Debitos -> modal de Debitos
      4. Click no icone de impressora da linha "Unica"
      5. Modal "Informar a data de pagamento" abre na aba Detalhes
      6. Click no icone do calendario -> click no header mes/ano ->
         click no nome do mes alvo (lowercase) -> click no dia alvo
      7. Click "Previa do valor"
      8. Click "Imprimir" -> abre nova aba com PDF do Jasper
      9. Captura PDF via cookies do contexto
    """
    cci = lote["cci"]
    comprador = lote["comprador"]
    caminho_pdf = montar_caminho(lote, pasta_saida)
    nome_arquivo = caminho_pdf.name
    caminho_pdf.parent.mkdir(parents=True, exist_ok=True)

    if caminho_pdf.exists():
        print(f"  [PULA] Ja existe: {nome_arquivo}")
        return True, nome_arquivo

    print(f"  [>>] CCI {cci} | {comprador[:35]}")

    # Parse data alvo
    dia_target, mes_target_num, ano_target = DATA_PAGAMENTO.split("/")
    MESES_LOWER = ["janeiro", "fevereiro", "mar\u00e7o", "abril", "maio", "junho",
                   "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    mes_lower = MESES_LOWER[int(mes_target_num) - 1]
    dia_label = str(int(dia_target))  # remove zero a esquerda

    nova_aba = None
    try:
        # Forca reload completo (a URL eh hash-routed, page.goto() na mesma URL
        # nao recarrega o SPA - modais e estado do lote anterior persistiriam).
        try:
            await page.goto("about:blank", timeout=10000)
            await asyncio.sleep(0.3)
        except Exception:
            pass
        await page.goto(URL_SITE, wait_until="domcontentloaded", timeout=30000)
        await esperar_pagina_pronta(page)

        # 1) Fill CCI
        campo_cci = page.get_by_role("textbox", name="C\u00f3digo do im\u00f3vel (CCI)")
        await campo_cci.wait_for(state="visible", timeout=15000)
        await campo_cci.click()
        await campo_cci.fill(cci)
        await asyncio.sleep(0.4)

        # 2) Click Pesquisar
        await page.get_by_role("button", name="Pesquisar").first.click()
        print(f"  [.] clicou Pesquisar, aguardando {PAUSA_ENTRE_CLIQUES}s...")
        try:
            await page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        await asyncio.sleep(PAUSA_ENTRE_CLIQUES)

        # 3) Dblclick na linha do resultado (habilita Exibir + as vezes abre modal).
        # Combina o seletor do codegen (get_by_role) com fallback CSS que ja funcionou.
        linha_res = page.get_by_role("row").filter(has_text=cci).first
        try:
            await linha_res.wait_for(state="visible", timeout=8000)
        except Exception:
            # Fallback: CSS-based (a tabela de resultados pode nao expor role=row)
            linha_res = page.locator(
                f'tr:has-text("{cci}"), [role="row"]:has-text("{cci}"), '
                f'.ui-grid-row:has-text("{cci}")'
            ).first
            try:
                await linha_res.wait_for(state="visible", timeout=15000)
            except Exception:
                await _diag_screenshot(page, f"sem_resultado_CCI{cci}")
                raise RuntimeError(f"linha do imovel CCI {cci} nao apareceu")

        await linha_res.scroll_into_view_if_needed()
        await linha_res.dblclick(force=True)
        print(f"  [.] dblclick na linha do resultado")
        await asyncio.sleep(1)

        # 4) Click Exibir Debitos (se ainda nao abriu a modal automaticamente)
        btn_exibir = page.get_by_role("button", name="Exibir D\u00e9bitos")
        try:
            if await btn_exibir.count() > 0 and await btn_exibir.is_enabled():
                await btn_exibir.click()
                print(f"  [.] clicou Exibir Debitos")
        except Exception:
            pass
        await asyncio.sleep(PAUSA_ENTRE_CLIQUES)

        # 5) Modal de Debitos aberta. Pode estar em 3 estados:
        #    (a) Tem debitos com opcao Unica -> baixa boleto a vista
        #    (b) Tem debitos sem opcao Unica -> so parcelado, pula (SEM_UNICA)
        #    (c) Nenhum debito -> "Nenhum registro" no grid, pula (SEM_DEBITO)
        try:
            estado = await page.wait_for_function(
                """() => {
                    const txt = (document.body.innerText || '').toLowerCase();
                    if (txt.includes('unica') || txt.includes('única')) return 'com_debito';
                    if (txt.includes('nenhum registro')) return 'sem_debito';
                    // Modal carregou (tem total/cabecalho) mas sem Unica nem 'Nenhum registro'
                    // -> tem so parcelas individuais
                    if (txt.includes('total d\\u00e9bito')
                        || txt.includes('total debito')
                        || txt.includes('total a pagar')) {
                        return 'sem_unica';
                    }
                    return false;
                }""",
                timeout=20000,
            )
            estado_str = await estado.json_value()
        except Exception:
            await _diag_screenshot(page, f"modal_debitos_indeterminado_CCI{cci}")
            raise RuntimeError("modal de Debitos abriu mas estado indeterminado (nem Unica nem 'Nenhum registro' nem 'Total')")

        if estado_str == "sem_debito":
            print(f"  [SEM DEBITO] CCI {cci} nao possui debitos pendentes (skip)")
            registrar_falha(cci, "SemDebito: CCI sem debitos na tabela")
            return True, f"SEM_DEBITO_CCI{cci}"

        if estado_str == "sem_unica":
            print(f"  [SEM UNICA] CCI {cci} tem debitos mas sem parcela Unica (skip)")
            registrar_falha(cci, "SemUnica: CCI sem opcao de parcela unica")
            return True, f"SEM_UNICA_CCI{cci}"

        # Tem debitos -> localiza a linha Unica
        linha_unica = page.get_by_role("row").filter(has_text="Unica").first
        try:
            await linha_unica.wait_for(state="visible", timeout=8000)
        except Exception:
            linha_unica = page.locator(
                'tr:has-text("Unica"), [role="row"]:has-text("Unica"), '
                '.ui-grid-row:has-text("Unica")'
            ).first
            try:
                await linha_unica.wait_for(state="visible", timeout=12000)
            except Exception:
                await _diag_screenshot(page, f"sem_unica_CCI{cci}")
                raise RuntimeError("linha 'Unica' nao apareceu na tabela de Debitos")

        await linha_unica.scroll_into_view_if_needed()
        # Click no primeiro botao/link da linha (icone de impressora)
        btn_print = linha_unica.locator("button, a").first
        try:
            await btn_print.click(force=True, timeout=8000)
        except Exception:
            # Fallback JS: clica no primeiro elemento clicavel da linha
            await linha_unica.evaluate("""
                row => {
                    const cand = row.querySelector('[ng-click], button, a, i.fa-print');
                    if (cand) cand.click();
                }
            """)
        print(f"  [.] clicou icone impressora da Unica")
        await asyncio.sleep(PAUSA_ENTRE_CLIQUES)

        # 6) Modal "Informar a data de pagamento" abre.
        #    Click no icone do calendario (esta dentro de um <p>).
        try:
            await page.wait_for_function(
                "() => /data de pagamento/i.test(document.body.innerText)",
                timeout=15000,
            )
        except Exception:
            await _diag_screenshot(page, f"sem_modal_data_CCI{cci}")
            raise RuntimeError("modal 'Informar a data de pagamento' nao abriu")
        await asyncio.sleep(1)

        await page.get_by_role("paragraph").get_by_role("button").first.click(force=True)
        print(f"  [.] clicou icone calendario")
        await asyncio.sleep(2.5)   # popup do calendario demora ~2s pra renderizar

        # 7) Click no heading mes/ano DENTRO do datepicker
        # (a primeira tentativa pega qualquer heading - pode estar pegando o title
        # da modal, por isso escopamos ao container do uib-datepicker).
        container_dp = page.locator('[id^="datepicker-"]').first
        try:
            await container_dp.wait_for(state="visible", timeout=15000)
        except Exception:
            pass

        # Estrategia 1: button.uib-title dentro do datepicker
        # Estrategia 2: heading dentro do datepicker
        # Estrategia 3: heading com nome de algum mes (lowercase)
        title_btn = None
        for sel in ('button.uib-title', 'button.btn-default'):
            cand = container_dp.locator(sel).first
            if await cand.count() > 0:
                title_btn = cand
                break
        if title_btn is None:
            cand = container_dp.get_by_role("heading").first
            if await cand.count() > 0:
                title_btn = cand
        if title_btn is None:
            # Por nome do mes (busca em toda a pagina)
            for m in MESES_LOWER:
                h = page.get_by_role("heading", name=m).first
                if await h.count() > 0:
                    title_btn = h
                    break

        if title_btn is None:
            await _diag_screenshot(page, f"sem_heading_cal_CCI{cci}")
            raise RuntimeError("heading mes-ano do calendario nao encontrado")

        await title_btn.click(force=True, timeout=8000)
        print(f"  [.] clicou heading mes-ano")
        await asyncio.sleep(1)

        # 8) Click no mes alvo (ex: "maio" lowercase) DENTRO do datepicker
        btn_mes = container_dp.get_by_role("button", name=mes_lower).first
        if await btn_mes.count() == 0:
            # Fallback: na pagina inteira
            btn_mes = page.get_by_role("button", name=mes_lower).first
        try:
            await btn_mes.click(force=True, timeout=10000)
            print(f"  [.] clicou mes '{mes_lower}'")
        except Exception as e:
            await _diag_screenshot(page, f"sem_mes_CCI{cci}")
            raise RuntimeError(f"mes '{mes_lower}' nao clicavel: {e}")
        await asyncio.sleep(0.8)

        # 9) Click no dia alvo. Estrategia robusta via JS:
        # No calendario Prodata, os dias do mes anterior (muted/cinza) aparecem
        # na linha de cima e os do proximo mes na linha de baixo. O "29" do
        # mes ALVO esta sempre entre eles. Como nem sempre conseguimos detectar
        # 'text-muted' pelo nome da classe, usamos a regra:
        #   - dentre os "29" visiveis no datepicker corrente, pegar o ULTIMO
        #     que nao seja muted/disabled (ou o ultimo total como fallback).
        # April-29 sempre vem antes de May-29 no DOM, entao pegar o LAST.
        resultado_dia = await page.evaluate(
            """(dia) => {
                const dps = document.querySelectorAll('[id^="datepicker-"], .uib-datepicker, [uib-datepicker]');
                let alvo_dp = null;
                for (const d of dps) {
                    if (d.offsetParent !== null) { alvo_dp = d; break; }
                }
                if (!alvo_dp) return {ok: false, motivo: 'nenhum datepicker visivel'};

                const buttons = alvo_dp.querySelectorAll('button');
                const candidatos = [];
                for (const b of buttons) {
                    const t = (b.innerText || '').trim();
                    if (t !== dia) continue;
                    const td = b.closest('td');
                    const cls = ((b.className||'') + ' ' + (td?.className||'')).toLowerCase();
                    const style = getComputedStyle(b);
                    const styleSpan = b.querySelector('span') ? getComputedStyle(b.querySelector('span')) : null;
                    const colorGray = (style.color && /128|153|170|187|204|221/.test(style.color));
                    const fontItalic = style.fontStyle === 'italic' || (styleSpan && styleSpan.fontStyle === 'italic');
                    const isMuted = cls.includes('muted') || cls.includes('outside')
                                  || cls.includes('off') || b.disabled
                                  || colorGray || fontItalic;
                    candidatos.push({text: t, muted: isMuted, cls: cls.slice(0,50),
                                     color: style.color, italic: fontItalic, y: b.getBoundingClientRect().y});
                }
                // 1a tentativa: pega o ULTIMO nao-muted (mes corrente vem depois dos
                // dias do mes anterior no DOM).
                let alvo = null;
                for (let i = candidatos.length - 1; i >= 0; i--) {
                    if (!candidatos[i].muted) {
                        // localiza o botao real correspondente (na mesma ordem)
                        let idx = 0;
                        for (const b of buttons) {
                            const t = (b.innerText || '').trim();
                            if (t === dia) {
                                if (idx === i) { alvo = b; break; }
                                idx++;
                            }
                        }
                        break;
                    }
                }
                // Fallback: pega o ULTIMO em geral
                if (!alvo) {
                    const todos29 = Array.from(buttons).filter(b => (b.innerText||'').trim() === dia);
                    if (todos29.length > 0) alvo = todos29[todos29.length - 1];
                }
                if (!alvo) return {ok: false, motivo: 'nenhum "' + dia + '" encontrado', candidatos};
                alvo.click();
                return {ok: true, clicado: dia, candidatos};
            }""",
            dia_label,
        )
        print(f"  [.] click dia '{dia_label}': {resultado_dia}")
        if not resultado_dia or not resultado_dia.get("ok"):
            await _diag_screenshot(page, f"sem_dia_CCI{cci}")
            raise RuntimeError(f"dia '{dia_label}' nao clicado: {resultado_dia}")
        await asyncio.sleep(PAUSA_ENTRE_CLIQUES)

        # 10) Click "Pr\u00e9via do valor"
        await page.get_by_role("button", name="Pr\u00e9via do valor").first.click(force=True)
        print(f"  [.] clicou Pr\u00e9via do valor, aguardando {PAUSA_ENTRE_CLIQUES}s...")
        await asyncio.sleep(PAUSA_ENTRE_CLIQUES)

        # 11) Click "Imprimir" -> nova aba (popup) ou download direto.
        # Registra listeners ANTES do click para capturar:
        #  - URLs Jasper (caso a aba feche antes de ler)
        #  - Body da RESPONSE direto do navegador (evita 2a requisicao que o
        #    servidor recusa pois ja consumiu o timestamp)
        #  - Downloads diretos (Content-Disposition: attachment)
        urls_jasper = []
        downloads_capt = []
        pdf_body_capt = []  # bytes do PDF capturados na response da nova aba

        def _on_req(request):
            u = request.url or ""
            if ("jasper" in u.lower() or "imprimir" in u.lower()
                    or u.lower().endswith(".pdf")):
                urls_jasper.append(u)

        async def _capt_response_body(response):
            try:
                u = response.url or ""
                if not ("jasper" in u.lower() or u.lower().endswith(".pdf")):
                    return
                if pdf_body_capt:
                    return  # ja capturou
                body = await response.body()
                if body and body[:4] == b"%PDF":
                    pdf_body_capt.append(body)
            except Exception:
                pass

        def _on_resp(response):
            # dispara captura async sem bloquear
            asyncio.create_task(_capt_response_body(response))

        def _on_download(download):
            downloads_capt.append(download)

        page.context.on("request", _on_req)
        page.context.on("response", _on_resp)
        page.context.on("download", _on_download)

        nova_aba = None
        try:
            async with page.expect_popup(timeout=20000) as popup_info:
                await page.get_by_role("button", name="Imprimir").first.click(force=True)
                print(f"  [.] clicou Imprimir, aguardando popup (ate 20s)...")
            nova_aba = await popup_info.value
            # Espera breve por load + url
            try:
                await nova_aba.wait_for_load_state("domcontentloaded", timeout=8000)
            except Exception:
                pass
        except Exception:
            print("  [.] popup nao apareceu (pode ter sido download direto)")

        # Da tempo do request fluir e da nova aba carregar
        await asyncio.sleep(3)

        # Captura URL da aba se ainda estiver aberta
        url_da_aba = None
        if nova_aba is not None:
            try:
                if not nova_aba.is_closed():
                    url_da_aba = nova_aba.url
            except Exception:
                pass

        # Da tempo extra pra response listener capturar o body
        for _ in range(20):
            if pdf_body_capt or downloads_capt:
                break
            await asyncio.sleep(0.5)

        # Remove os listeners
        try:
            page.context.remove_listener("request", _on_req)
            page.context.remove_listener("response", _on_resp)
            page.context.remove_listener("download", _on_download)
        except Exception:
            pass

        print(f"  [.] URLs Jasper capturadas: {len(urls_jasper)}")
        for u in urls_jasper[:5]:
            print(f"      {u[:120]}")
        print(f"  [.] URL da nova aba: {url_da_aba}")
        print(f"  [.] downloads capturados: {len(downloads_capt)}")
        print(f"  [.] PDF body capturado via response listener: "
              f"{len(pdf_body_capt[0]) if pdf_body_capt else 0} bytes")

        pdf_bytes = None

        # 1) Response listener (caminho mais confiavel - body direto do navegador)
        if pdf_body_capt:
            pdf_bytes = pdf_body_capt[0]
            print(f"  [.] PDF capturado via response listener")

        # 2) Download event (Content-Disposition: attachment)
        if not pdf_bytes and downloads_capt:
            try:
                download = downloads_capt[-1]
                path = await download.path()
                if path:
                    pdf_bytes = Path(path).read_bytes()
                    print(f"  [.] PDF capturado via download event")
            except Exception as e:
                print(f"  [.] erro lendo download: {e}")

        # 3) Refetch via context.request (so se nada mais funcionou - servidor
        #    pode recusar pois ja consumiu o timestamp, mas vale tentar)
        if not pdf_bytes:
            url_para_baixar = None
            if url_da_aba and "about:blank" not in url_da_aba:
                url_para_baixar = url_da_aba
            elif urls_jasper:
                url_para_baixar = urls_jasper[-1]
            if url_para_baixar:
                try:
                    response = await page.context.request.get(url_para_baixar)
                    if response.ok:
                        body = await response.body()
                        ct = (response.headers.get("content-type") or "").lower()
                        if b"%PDF" in body[:1024] or "pdf" in ct:
                            pdf_bytes = body
                            print(f"  [.] PDF baixado via refetch URL")
                        else:
                            try:
                                t = body[:4096].decode("utf-8", errors="ignore")
                                if _eh_erro_barcode(t):
                                    raise BoletoIndisponivel(
                                        "Jasper retornou HTML de erro em vez de PDF"
                                    )
                            except BoletoIndisponivel:
                                raise
                            except Exception:
                                pass
                except BoletoIndisponivel:
                    if nova_aba is not None:
                        try:
                            await nova_aba.close()
                        except Exception:
                            pass
                    print(f"  [PULADO] CCI {cci}: boleto indisponivel")
                    registrar_falha(cci, "BoletoIndisponivel: refetch retornou HTML de erro")
                    return True, f"BOLETO_INDISPONIVEL_CCI{cci}"
                except Exception as e:
                    print(f"  [.] erro refetch URL: {e}")

        if not pdf_bytes:
            await _diag_screenshot(page, f"sem_pdf_CCI{cci}")
            raise RuntimeError("nao consegui capturar o PDF (sem listener, download, aba ou URL)")

        # Fecha a aba se ainda estiver aberta
        if nova_aba is not None:
            try:
                if not nova_aba.is_closed():
                    await nova_aba.close()
            except Exception:
                pass

        caminho_pdf.write_bytes(pdf_bytes)
        print(f"  [OK] Salvo: {nome_arquivo}")
        return True, nome_arquivo

    except BoletoIndisponivel as e:
        print(f"  [PULADO] CCI {cci}: {e}")
        registrar_falha(cci, f"BoletoIndisponivel: {e}")
        if nova_aba is not None:
            try:
                await nova_aba.close()
            except Exception:
                pass
        return True, f"BOLETO_INDISPONIVEL_CCI{cci}"

    except Exception as e:
        motivo = f"{type(e).__name__}: {str(e)[:200]}"
        print(f"  [FALHA] CCI {cci}: {motivo}")
        registrar_falha(cci, motivo)
        if nova_aba is not None:
            try:
                await nova_aba.close()
            except Exception:
                pass
        return False, None



# ----------------------------- DEBUG ----------------------------------------

async def debug_inspecionar():
    from playwright.async_api import async_playwright
    print("Abrindo navegador para inspecionar a pagina...")
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width": 1366, "height": 900})
        page = await context.new_page()
        try:
            await page.goto(URL_SITE, wait_until="domcontentloaded", timeout=30000)
            try:
                await page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass

            # Espera ate o botao PESQUISAR aparecer (sinal de que o form renderizou)
            print("Aguardando o formulario renderizar (PESQUISAR)...")
            try:
                await page.wait_for_selector(
                    'button:has-text("PESQUISAR"), button:has-text("Pesquisar")',
                    timeout=25000,
                )
                print("  -> botao PESQUISAR detectado.")
            except Exception:
                print("  -> botao PESQUISAR nao apareceu na pagina principal em 25s.")
            await asyncio.sleep(2)

            print(f"\nURL atual : {page.url}")
            print(f"Titulo    : {await page.title()}")

            # Lista os frames
            frames = page.frames
            print(f"\n=== {len(frames)} FRAME(S) NA PAGINA ===")
            for i, fr in enumerate(frames):
                print(f"  frame[{i}]: name={fr.name!r:<20} url={fr.url[:90]}")

            # Para cada frame, lista inputs e botoes
            for i, fr in enumerate(frames):
                try:
                    inputs = await fr.evaluate(JS_LISTAR_INPUTS)
                except Exception as e:
                    print(f"\n[frame {i}] erro avaliando inputs: {e}")
                    continue
                print(f"\n=== FRAME {i}: {len(inputs)} INPUTS ===")
                for inp in inputs:
                    vis = "VIS" if inp["visible"] else "esc"
                    pf = "pdF" if inp.get("inPdForm") else "   "
                    print(
                        f"[{inp['idx']:02d}] {vis} {pf} type={(inp['type'] or '-'):<6} "
                        f"id={inp['id'][:16]:<16} "
                        f"ng-model={inp['ngModel'][:24]:<24} "
                        f"pos=({inp['x']:>4},{inp['y']:>4}) {inp['w']:>3}x{inp['h']:>2}"
                    )
                    label = inp.get("label", "")
                    if label:
                        print(f"      label='{label[:60]}'")
                    if inp["placeholder"]:
                        print(f"      placeholder='{inp['placeholder'][:60]}'")
                # Botoes do frame
                try:
                    botoes = await fr.evaluate("""
                        () => Array.from(document.querySelectorAll('button')).map(b => ({
                            text: (b.innerText || b.textContent || '').trim().slice(0,40),
                            id: b.id || '',
                            ngClick: b.getAttribute('ng-click') || '',
                            visible: b.offsetParent !== null,
                        })).filter(b => b.text)
                    """)
                    print(f"  -- {len(botoes)} botoes:")
                    for b in botoes[:20]:
                        v = "V" if b["visible"] else "e"
                        print(f"    [{v}] '{b['text']}' id={b['id']} ng-click={b['ngClick'][:30]}")
                except Exception:
                    pass

            # O que o script escolheria como campo CCI
            print("\n=== O QUE O SCRIPT ESCOLHERIA COMO CAMPO CCI ===")
            try:
                _, escolhido = await localizar_campo_cci(page)
                print(f"INDICE no DOM: {escolhido['idx']}")
                print(f"  id          : {escolhido['id']}")
                print(f"  ng-model    : {escolhido['ngModel']}")
                print(f"  placeholder : {escolhido['placeholder']}")
                print(f"  label       : {escolhido.get('label','')}")
                print(f"  posicao     : x={escolhido['x']} y={escolhido['y']}")
            except Exception as e:
                print(f"  [ERRO] {e}")

            # Salva screenshot
            screenshot_path = Path("debug_screenshot.png").resolve()
            try:
                await page.screenshot(path=str(screenshot_path), full_page=True)
                print(f"\nScreenshot salva: {screenshot_path}")
            except Exception as e:
                print(f"\n[AVISO] Nao conseguiu salvar screenshot: {e}")

            print("\nJanela ficara aberta por 30s. Inspecione manualmente se quiser.")
            await asyncio.sleep(30)
        finally:
            await browser.close()


# ----------------------------- MAIN -----------------------------------------

async def executar(planilha_path, limite, inicio, headless, cci_filtro=None):
    from playwright.async_api import async_playwright

    pendentes, wb, ws = carregar_pendentes(planilha_path)

    if cci_filtro:
        pendentes = [p for p in pendentes if p["cci"] == str(cci_filtro)]
        print(f"  (Filtrando por CCI={cci_filtro}: {len(pendentes)} lote(s))")
    if inicio > 0:
        pendentes = [p for p in pendentes if p["linha"] >= inicio]
        print(f"  (Iniciando a partir da linha {inicio})")
    if limite:
        pendentes = pendentes[:limite]
        print(f"  (Modo teste: processando {len(pendentes)} lotes)")

    if not pendentes:
        print("[OK] Nada pendente. Tudo ja foi baixado.")
        return

    pasta_saida = Path(PASTA_SAIDA)
    pasta_saida.mkdir(exist_ok=True)
    print(f"PDFs serao salvos em: {pasta_saida.resolve()}")

    sucesso = falhou = 0
    total = len(pendentes)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=headless,
            args=[
                "--disable-popup-blocking",
                "--disable-features=IsolateOrigins,site-per-process",
            ],
        )
        context = await browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 900},
        )
        page = await context.new_page()

        print(f"\n[INICIO] Baixando {total} IPTUs...\n")

        for idx, lote in enumerate(pendentes, 1):
            print(f"[{idx}/{total}] linha {lote['linha']} ", end="")
            ok, nome = await baixar_lote(page, lote, pasta_saida)
            if ok:
                sucesso += 1
                marcar_baixado(ws, lote["linha"], nome or montar_nome_arquivo(lote))
                salvar_planilha(wb, planilha_path)
            else:
                falhou += 1
            if idx < total:
                await asyncio.sleep(PAUSA_ENTRE_LOTES)

        salvar_planilha(wb, planilha_path)
        await browser.close()

    print("\n" + "=" * 50)
    print("Concluido!")
    print(f"  Sucessos : {sucesso}")
    print(f"  Falhas   : {falhou}")
    print(f"  PDFs em  : {Path(PASTA_SAIDA).resolve()}")
    if falhou:
        print(f"  Log falhas: {Path(ARQ_LOG_FALHAS).resolve()}")
    print("=" * 50)


def main():
    parser = argparse.ArgumentParser(
        description="Download automatico de IPTUs - Cidade Exemplo"
    )
    parser.add_argument("--debug", action="store_true",
                        help="Inspeciona campos da pagina e sai")
    parser.add_argument("--teste", type=int, default=0, metavar="N",
                        help="Baixa apenas os primeiros N lotes pendentes")
    parser.add_argument("--inicio", type=int, default=0, metavar="LINHA",
                        help="Linha da planilha para comecar (skip ate la)")
    parser.add_argument("--cci", type=str, default="", metavar="CCI",
                        help="Roda apenas o lote com este CCI (ex: --cci 276739)")
    parser.add_argument("--data", type=str, default="", metavar="DD/MM/YYYY",
                        help="Sobrescreve DATA_PAGAMENTO (ex: --data 29/05/2026)")
    parser.add_argument("--pasta", type=str, default="", metavar="NOME",
                        help="Sobrescreve pasta de saida (ex: --pasta IPTUs_29_05)")
    parser.add_argument("--headless", action="store_true",
                        help="Roda sem janela (padrao: visivel)")
    args = parser.parse_args()

    if args.debug:
        asyncio.run(debug_inspecionar())
        return

    # Override de constantes via CLI
    if args.data:
        global DATA_PAGAMENTO
        DATA_PAGAMENTO = args.data.strip()
        print(f"[OVERRIDE] DATA_PAGAMENTO = {DATA_PAGAMENTO}")
    if args.pasta:
        global PASTA_SAIDA
        PASTA_SAIDA = args.pasta.strip()
        print(f"[OVERRIDE] PASTA_SAIDA = {PASTA_SAIDA}")

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        print("Procurei nestes nomes:")
        for n in NOMES_PLANILHA:
            print(f"  - {n}")
        print("XLSX presentes na pasta atual:")
        for f in Path(".").glob("*.xlsx"):
            print(f"  - {f.name}")
        sys.exit(1)

    asyncio.run(executar(
        planilha_path=str(planilha),
        limite=args.teste if args.teste > 0 else None,
        inicio=args.inicio,
        headless=args.headless,
        cci_filtro=args.cci.strip() or None,
    ))


if __name__ == "__main__":
    main()
