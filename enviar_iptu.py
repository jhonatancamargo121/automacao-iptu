# -*- coding: utf-8 -*-
"""
AUTOMACAO - ENVIO DE BOLETOS IPTU POR E-MAIL (Outlook Web)

Le a planilha, encontra os lotes com PDF ja baixado, e envia cada um por
e-mail via Outlook Web. Tira screenshot do envio como prova.

Uso:
  python enviar_iptu.py --teste 1
      -> Envia 1 boleto (rotaciona destino entre 5 e-mails de teste, adiciona
         marcador "email test" no rodape do corpo).
  python enviar_iptu.py --teste 5
      -> Envia 5 boletos em modo teste.
  python enviar_iptu.py --cci 276794
      -> Envia o boleto de 1 CCI especifico (modo teste se sem --producao).
  python enviar_iptu.py --producao
      -> Modo producao: usa e-mail do cliente como destinatario (TBD onde vem).
  python enviar_iptu.py --headless
      -> Roda sem janela (NAO usar na primeira execucao por causa do 2FA).

Primeira execucao: o navegador abre, voce faz login + 2FA manualmente, e a
sessao fica salva em outlook_session.json para reutilizar.
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

PASTA_PDFS         = "IPTUs_Baixados"
PASTA_PROVAS       = "Provas_Envio"
ARQ_SESSAO_OUTLOOK = "outlook_session.json"
ARQ_LOG_FALHAS     = "falhas_envio.log"

URL_OUTLOOK = "https://outlook.office.com/mail/"

REMETENTE = "usuario@exemplo.com.br"
ASSUNTO   = "Boleto IPTU 2026 - \Cidade Exemplo"

# Em PRODUCAO esses 5 vao sempre em CC. Em TESTE NAO vao.
EMAILS_FIXOS = [
    "contato@exemplo.com.br",
    "adm-cidadeexemplo@exemplo.com.br",
    "adm-natela@exemplo.com.br",
    "teste2@exemplo.com.br",
    "teste1@exemplo.com.br",
]

# Em TESTE: Para rotaciona entre os 5 acima, CC sao os outros 4.
# (Pode sobrescrever o destinatario via --email-destino se precisar.)

TEMPLATE_CORPO = (
    "Prezado cliente {COMPRADOR};\n\n"
    "Segue em anexo o(s) boleto(s) do seu IPTU 2026 referente ao(s) seu(s) "
    "lote(s). Aproveite o desconto de 30% pagando ate o dia 29/05/2026.\n\n"
    "QUADRA/LOTE: {QUADRA}/{LOTE}\n"
    "CCI: {CCI}\n"
    "Link: https://sig.prefeitura-exemplo.gov.br/sig/app.html"
    "#/servicosonline/debito-imovel\n\n"
    "As CCIs de seu(s) lote(s) esta(ao) tambem informada nos boletos em anexo, "
    "caso necessite emitir em outra data de vencimento.\n\n"
    "Lembramos que manter o IPTU em dia e fundamental para a valorizacao "
    "e regularidade do seu imovel."
)

# Colunas da planilha (mesmas do baixar_iptu.py)
COL_STATUS    = 0
COL_QUADRA    = 1
COL_LOTE      = 2
COL_COMPRADOR = 8
COL_CCI       = 10
COL_OBS_IPTU  = 12
COL_ETAPA     = 13

PAUSA_ENTRE_ENVIOS = 8   # segundos entre cada envio (anti-rate-limit)

JA_ENVIADO_MARKER = "Guia enviada e-mail cliente"


# ----------------------------- UTILITARIOS ----------------------------------

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


def montar_caminho_pdf(lote, pasta_pdfs):
    etapa = extrair_etapa(lote["etapa"])
    pasta_etapa = f"ETAPA{etapa}" if etapa else "SEM_ETAPA"
    subpasta = Path(pasta_pdfs) / pasta_etapa / limpar_nome(lote["comprador"])
    return subpasta / montar_nome_arquivo(lote)


def achar_pdf_por_cci(cci, pasta_pdfs):
    """
    Procura recursivamente em pasta_pdfs um PDF cujo nome contenha 'CCI{cci}_'
    ou 'CCI{cci}.'. Retorna Path ou None.
    """
    base = Path(pasta_pdfs)
    if not base.exists():
        return None
    marcadores = (f"CCI{cci}_", f"CCI{cci}.")
    for pdf in base.rglob("*.pdf"):
        nome = pdf.name
        if any(m in nome for m in marcadores):
            return pdf
    return None


def carregar_lotes_para_envio(planilha_path):
    """Retorna lotes que tem PDF baixado e ainda nao foram enviados por email."""
    print(f"Lendo planilha: {planilha_path}")
    wb = openpyxl.load_workbook(planilha_path)
    ws = wb.active
    lotes = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row or len(row) <= COL_ETAPA:
            continue
        status = str(row[COL_STATUS] or "").strip().lower()
        cci    = row[COL_CCI]
        obs    = str(row[COL_OBS_IPTU] or "").strip().lower()
        if status != "vendido" or not cci:
            continue
        # Precisa ter PDF baixado (mas nao um marker SEM_DEBITO/SEM_UNICA/INDISPONIVEL)
        if "pdf baixado" not in obs:
            continue
        if any(t in obs for t in ("sem_debito", "sem_unica", "indisponivel", "boleto_indisponivel")):
            continue
        # E nao pode ter sido enviado ainda
        if "guia enviada e-mail cliente" in obs:
            continue
        try:
            cci_str = str(int(float(cci)))
        except (TypeError, ValueError):
            continue
        lotes.append({
            "linha":     i,
            "quadra":    str(row[COL_QUADRA]    or "").strip(),
            "lote":      str(row[COL_LOTE]      or "").strip(),
            "comprador": str(row[COL_COMPRADOR] or "").strip(),
            "cci":       cci_str,
            "etapa":     str(row[COL_ETAPA]     or "").strip(),
        })
    print(f"[OK] {len(lotes)} lotes com PDF baixado aguardando envio.")
    return lotes, wb, ws


def marcar_enviado(ws, linha, info_envio):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    val_atual = ws.cell(row=linha, column=COL_OBS_IPTU + 1).value or ""
    novo = f"{JA_ENVIADO_MARKER} {agora} - {info_envio}"
    # Mantem o marcador "PDF baixado" antigo se houver, separado por '|'
    if "PDF baixado" in str(val_atual) and JA_ENVIADO_MARKER not in str(val_atual):
        novo = f"{val_atual} | {novo}"
    ws.cell(row=linha, column=COL_OBS_IPTU + 1).value = novo


def registrar_falha(cci, motivo):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ARQ_LOG_FALHAS, "a", encoding="utf-8") as f:
        f.write(f"{ts}\tCCI={cci}\t{motivo}\n")


def salvar_planilha(wb, planilha_path):
    try:
        wb.save(planilha_path)
        return True
    except PermissionError:
        print("  [AVISO] Planilha aberta no Excel. Feche para salvar.")
        return False


def substituir_placeholders(template, lote, modo_teste=False):
    corpo = template.format(
        COMPRADOR=lote["comprador"],
        QUADRA=lote["quadra"],
        LOTE=lote["lote"],
        CCI=lote["cci"],
    )
    if modo_teste:
        corpo += "\n\n--- email test ---"
    return corpo


# ----------------------------- OUTLOOK WEB -----------------------------------

async def garantir_login(playwright, headless, storage_path):
    """
    Abre o navegador. Se ja existe outlook_session.json com cookies validos,
    reusa. Senao, abre janela e espera login manual + 2FA.
    Retorna (browser, context, page).
    """
    browser = await playwright.chromium.launch(
        headless=headless,
        args=["--disable-popup-blocking"],
    )
    storage_state = storage_path if Path(storage_path).exists() else None
    context = await browser.new_context(
        storage_state=storage_state,
        accept_downloads=True,
        viewport={"width": 1366, "height": 900},
    )
    page = await context.new_page()
    await page.goto(URL_OUTLOOK, wait_until="domcontentloaded", timeout=60000)

    # Detecta se esta logado: ha o botao "Novo email" / "New mail"
    print("Verificando login no Outlook...")
    botao_novo = page.get_by_role(
        "button", name=re.compile(r"^Novo email|^New mail", re.I)
    )
    try:
        await botao_novo.wait_for(state="visible", timeout=10000)
        print("[OK] Sessao do Outlook valida.")
    except Exception:
        print("\n>>> SESSAO EXPIRADA OU PRIMEIRO LOGIN. <<<")
        print("Faca login com:", REMETENTE)
        print("Aprove o 2FA no celular se solicitado.")
        print("Quando a caixa de entrada estiver visivel, pressione ENTER aqui.")
        # Espera o usuario logar manualmente
        await asyncio.get_event_loop().run_in_executor(None, input, "")
        # Salva a sessao
        await context.storage_state(path=storage_path)
        print(f"[OK] Sessao salva em {storage_path}")

    return browser, context, page


async def _commit_email(page, field_locator, email, timeout=3000):
    """Digita email num campo, tenta selecionar autocomplete; senao da Enter."""
    await field_locator.click()
    await field_locator.fill(email)
    await asyncio.sleep(0.8)
    # Estrategia 1: opcao com role="option"
    try:
        option = page.get_by_role(
            "option", name=re.compile(re.escape(email), re.I)
        ).first
        if await option.count() > 0:
            await option.click(timeout=timeout)
            return
    except Exception:
        pass
    # Estrategia 2: opcao com aria-label (visto no codegen do Outlook)
    try:
        option = page.get_by_label(
            re.compile(re.escape(email), re.I)
        ).first
        if await option.count() > 0:
            await option.click(timeout=timeout)
            return
    except Exception:
        pass
    # Sem autocomplete: confirma com Enter (Outlook aceita email novo)
    await page.keyboard.press("Enter")
    await asyncio.sleep(0.3)


async def compor_e_enviar(page, destinatario, cc_list, assunto, corpo, anexo_path):
    """Abre 'Novo email', preenche, anexa PDF, envia.
    Seletores capturados via Playwright Codegen na UI atual do Outlook (pt-BR).
    """
    # 1) Click "Novo email"
    await page.get_by_role("button", name="Novo email").click()
    await asyncio.sleep(2)

    # 2) Preenche "Para" (com tratamento de autocomplete)
    campo_para = page.get_by_label("Para", exact=True)
    await _commit_email(page, campo_para, destinatario)
    await asyncio.sleep(0.5)

    # 3) Preenche "Cc" (cada e-mail seguido de autocomplete ou Enter)
    if cc_list:
        campo_cc = page.get_by_label("Cc", exact=True)
        for email_cc in cc_list:
            await _commit_email(page, campo_cc, email_cc)
            await asyncio.sleep(0.3)

    # 4) Preenche Assunto
    campo_assunto = page.get_by_placeholder("Adicionar um assunto")
    await campo_assunto.click()
    await campo_assunto.fill(assunto)
    await asyncio.sleep(0.5)

    # 5) Preenche Corpo
    campo_corpo = page.get_by_label("Corpo da mensagem")
    await campo_corpo.click()
    await campo_corpo.fill(corpo)
    await asyncio.sleep(0.5)

    # 6) Anexa o PDF.
    # Estrategia: usa expect_file_chooser que captura o dialogo do SO
    # e injeta o PDF automaticamente (sem precisar abrir/navegar pelo Windows).
    caminho_absoluto = str(Path(anexo_path).resolve())
    print(f"  [.] anexando PDF: {caminho_absoluto}")
    try:
        async with page.expect_file_chooser(timeout=15000) as fc_info:
            await page.get_by_label("Anexar arquivo").click()
            await asyncio.sleep(0.8)
            await page.get_by_role("menuitem", name="Navegar neste computador").click()
        file_chooser = await fc_info.value
        await file_chooser.set_files(caminho_absoluto)
        print(f"  [.] arquivo injetado via file chooser")
    except Exception as e:
        # Fallback ao metodo direto via role=application
        print(f"  [.] file chooser nao capturou ({e}); tentando set_input_files direto")
        try:
            await page.get_by_role("application").set_input_files(caminho_absoluto)
        except Exception as e2:
            raise RuntimeError(f"falha ao anexar PDF: {e2}")

    # Espera o anexo aparecer na compose (cardzinho com nome do arquivo)
    nome_arq = Path(anexo_path).name
    print(f"  [.] aguardando '{nome_arq[:40]}' aparecer no e-mail...")
    try:
        await page.locator(f'text=/{re.escape(nome_arq[:40])}/i').first.wait_for(
            state="visible", timeout=30000
        )
        print(f"  [.] anexo confirmado")
    except Exception:
        print(f"  [.] timeout esperando confirmacao visual; aguardando fixo 8s")
        await asyncio.sleep(8)
    await asyncio.sleep(2)

    # 7) Click em Enviar
    await page.get_by_label("Enviar").click()
    print(f"  [.] clicou Enviar")


async def aguardar_envio_e_screenshot(page, screenshot_path, timeout=20000):
    """
    Espera ate o compose fechar (=mensagem enviada) e tira screenshot da inbox.
    """
    # Aguarda janela de compose desaparecer
    try:
        await page.wait_for_function(
            """() => {
                const btn = document.querySelector('[aria-label*="Enviar" i], [aria-label*="Send" i]');
                return !btn || btn.offsetParent === null;
            }""",
            timeout=timeout,
        )
    except Exception:
        pass
    await asyncio.sleep(2)
    Path(screenshot_path).parent.mkdir(parents=True, exist_ok=True)
    await page.screenshot(path=str(screenshot_path), full_page=True)
    print(f"  [.] screenshot salvo: {screenshot_path}")


async def enviar_lote(page, lote, idx_teste, modo_teste, modo_simular,
                       pasta_pdfs, pasta_provas, email_destino_override=None):
    cci = lote["cci"]
    comprador = lote["comprador"]

    # Procura o PDF pelo CCI (mais robusto que reconstruir o caminho).
    caminho_pdf = achar_pdf_por_cci(cci, pasta_pdfs)
    if caminho_pdf is None or not caminho_pdf.exists():
        # Fallback: tenta o caminho reconstruido pelo padrao
        caminho_pdf = montar_caminho_pdf(lote, pasta_pdfs)
    if not caminho_pdf.exists():
        # Diagnostico: lista o que existe na pasta da etapa
        etapa = extrair_etapa(lote["etapa"])
        pasta_etapa = Path(pasta_pdfs) / (f"ETAPA{etapa}" if etapa else "SEM_ETAPA")
        if pasta_etapa.exists():
            print(f"  [DIAG] PDFs na pasta {pasta_etapa.name}:")
            for f in list(pasta_etapa.rglob("*.pdf"))[:5]:
                print(f"      {f.relative_to(Path(pasta_pdfs))}")
        raise FileNotFoundError(
            f"PDF nao encontrado para CCI {cci}. "
            f"Procurei em: {Path(pasta_pdfs).resolve()}"
        )
    print(f"  [.] PDF encontrado: {caminho_pdf.relative_to(Path(pasta_pdfs))}")

    # Decide destinatario e CC
    if modo_simular:
        # Simula producao: pick deterministico do destinatario pelo CCI,
        # os 4 restantes viram CC.
        idx = int(cci) % len(EMAILS_FIXOS)
        destinatario = EMAILS_FIXOS[idx]
        cc_list = [e for e in EMAILS_FIXOS if e != destinatario]
    elif modo_teste:
        # Modo teste: rotaciona Para entre os 5. CC eh o PROXIMO da rotacao
        # (apenas 1 e-mail em CC).
        if email_destino_override:
            destinatario = email_destino_override
            # CC = primeiro fixo que nao seja o destinatario
            cc_list = [next((e for e in EMAILS_FIXOS if e != destinatario), EMAILS_FIXOS[0])]
        else:
            n = len(EMAILS_FIXOS)
            destinatario = EMAILS_FIXOS[idx_teste % n]
            cc_list = [EMAILS_FIXOS[(idx_teste + 1) % n]]
    else:
        # Em producao real: e-mail do cliente vem de outro arquivo (TBD).
        raise NotImplementedError(
            "Producao: o arquivo de e-mails dos clientes ainda nao foi configurado. "
            "Use --teste ou --simular por enquanto."
        )

    # Em modo simular tambem adiciona o marcador 'email test' pra ficar claro
    corpo = substituir_placeholders(
        TEMPLATE_CORPO, lote, modo_teste=(modo_teste or modo_simular)
    )

    print(f"  [.] destinatario: {destinatario}")
    if cc_list:
        print(f"  [.] CC: {', '.join(cc_list)}")
    print(f"  [.] anexo: {caminho_pdf.name}")

    await compor_e_enviar(
        page,
        destinatario=destinatario,
        cc_list=cc_list,
        assunto=ASSUNTO,
        corpo=corpo,
        anexo_path=caminho_pdf,
    )

    # Tira screenshot da prova com o MESMO padrao do PDF
    # (etapa, quadra, lote, CCI, nome do cliente).
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_nome = montar_nome_arquivo(lote).replace(".pdf", "")
    screenshot_path = Path(pasta_provas) / f"prova_{base_nome}_{ts}.png"
    await aguardar_envio_e_screenshot(page, screenshot_path)

    return f"para={destinatario} prova={screenshot_path.name}"


# ----------------------------- MAIN -----------------------------------------

async def executar(planilha_path, limite, inicio, cci_filtro, headless,
                   modo_teste, modo_simular, pasta_pdfs, pasta_provas,
                   email_destino_override=None):
    from playwright.async_api import async_playwright

    lotes, wb, ws = carregar_lotes_para_envio(planilha_path)
    if cci_filtro:
        lotes = [l for l in lotes if l["cci"] == str(cci_filtro)]
        print(f"  (Filtro CCI={cci_filtro}: {len(lotes)} lote(s))")
    if inicio > 0:
        lotes = [l for l in lotes if l["linha"] >= inicio]
        print(f"  (Iniciando da linha {inicio})")
    if limite:
        lotes = lotes[:limite]
        print(f"  (Modo teste: processando {len(lotes)} envio(s))")

    if not lotes:
        print("[OK] Nenhum lote pendente de envio.")
        return

    Path(pasta_provas).mkdir(exist_ok=True)
    sucesso = falhou = 0
    total = len(lotes)

    async with async_playwright() as p:
        browser, context, page = await garantir_login(
            p, headless=headless, storage_path=ARQ_SESSAO_OUTLOOK
        )
        try:
            print(f"\n[INICIO] Enviando {total} e-mails...\n")
            for idx, lote in enumerate(lotes, 1):
                print(f"[{idx}/{total}] linha {lote['linha']} | CCI {lote['cci']} "
                      f"| {lote['comprador'][:35]}")
                try:
                    info = await enviar_lote(
                        page, lote, idx_teste=idx - 1,
                        modo_teste=modo_teste,
                        modo_simular=modo_simular,
                        pasta_pdfs=pasta_pdfs,
                        pasta_provas=pasta_provas,
                        email_destino_override=email_destino_override,
                    )
                    marcar_enviado(ws, lote["linha"], info)
                    salvar_planilha(wb, planilha_path)
                    sucesso += 1
                    print(f"  [OK] envio confirmado.")
                except Exception as e:
                    motivo = f"{type(e).__name__}: {str(e)[:200]}"
                    print(f"  [FALHA] CCI {lote['cci']}: {motivo}")
                    registrar_falha(lote["cci"], motivo)
                    falhou += 1
                    # Tira screenshot do estado de falha
                    try:
                        ts = datetime.now().strftime("%H%M%S")
                        await page.screenshot(
                            path=f"diag_envio_{ts}_CCI{lote['cci']}.png",
                            full_page=True,
                        )
                    except Exception:
                        pass

                if idx < total:
                    await asyncio.sleep(PAUSA_ENTRE_ENVIOS)

            salvar_planilha(wb, planilha_path)
        finally:
            await browser.close()

    print("\n" + "=" * 50)
    print("Concluido!")
    print(f"  Sucessos : {sucesso}")
    print(f"  Falhas   : {falhou}")
    print(f"  Provas em: {Path(pasta_provas).resolve()}")
    if falhou:
        print(f"  Log falhas: {Path(ARQ_LOG_FALHAS).resolve()}")
    print("=" * 50)


def main():
    parser = argparse.ArgumentParser(
        description="Envio automatico de boletos IPTU por e-mail (Outlook Web)"
    )
    parser.add_argument("--teste", type=int, default=0, metavar="N",
                        help="Envia N e-mails em modo teste (rotaciona destinos fixos)")
    parser.add_argument("--inicio", type=int, default=0, metavar="LINHA",
                        help="Linha da planilha para comecar")
    parser.add_argument("--cci", type=str, default="", metavar="CCI",
                        help="Envia apenas o lote com este CCI")
    parser.add_argument("--producao", action="store_true",
                        help="Modo producao (e-mail do cliente como destino)")
    parser.add_argument("--simular", action="store_true",
                        help="Simula producao usando os 5 e-mails fixos: "
                             "destinatario escolhido pelo CCI (hash), outros 4 em CC")
    parser.add_argument("--headless", action="store_true",
                        help="Sem janela. NAO usar na 1a execucao (2FA).")
    parser.add_argument("--email-destino", type=str, default="", metavar="EMAIL",
                        help="Sobrescreve destinatario em modo teste "
                             "(ex: --email-destino fulano@exemplo.com)")
    args = parser.parse_args()

    planilha = achar_planilha()
    if planilha is None:
        print("[ERRO] Planilha nao encontrada.")
        sys.exit(1)

    modo_simular = args.simular
    modo_teste = (not args.producao) and (not modo_simular)
    email_dest = args.email_destino.strip() or None

    if modo_simular:
        print(f"[MODO SIMULAR] Comportamento de producao com os 5 e-mails fixos:")
        print(f"               - Destinatario = EMAILS_FIXOS[hash(CCI) %% 5]")
        print(f"               - CC = os outros 4 e-mails")
        print(f"               - Marcador 'email test' no rodape")
    elif modo_teste:
        if email_dest:
            print(f"[MODO TESTE] Destinatario unico (override): {email_dest}")
            print(f"             CC: 1 e-mail fixo (proximo da rotacao)")
        else:
            print(f"[MODO TESTE] Rotacionando 'Para' entre {len(EMAILS_FIXOS)} e-mails:")
            for e in EMAILS_FIXOS:
                print(f"             - {e}")
            print(f"             CC: o proximo da rotacao (apenas 1 e-mail em cada envio)")
        print("             Marcador 'email test' adicionado no rodape do corpo.")
    else:
        print("[MODO PRODUCAO] Destinatario sera o e-mail real do cliente.")

    asyncio.run(executar(
        planilha_path=str(planilha),
        limite=args.teste if args.teste > 0 else None,
        inicio=args.inicio,
        cci_filtro=args.cci.strip() or None,
        headless=args.headless,
        modo_teste=modo_teste,
        modo_simular=modo_simular,
        pasta_pdfs=PASTA_PDFS,
        pasta_provas=PASTA_PROVAS,
        email_destino_override=email_dest,
    ))


if __name__ == "__main__":
    main()
