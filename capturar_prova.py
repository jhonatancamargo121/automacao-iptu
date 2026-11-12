# -*- coding: utf-8 -*-
"""
CAPTURAR PROVA DE ENVIO - Natela Analytics

Roda em background. Quando voce aperta o atalho (AltGr por padrao),
captura a tela inteira e salva como PDF em Provas_Envio_Manual/.

O nome do PDF e montado automaticamente a partir do conteudo da tela:
  QD{quadra}_LT{lote}_CCI{cci}_{NOME_COMPRADOR}_envio_iptu_2026.pdf

Tres caminhos para descobrir os dados (em cascata):
  1) OCR (le diretamente o texto da imagem) - precisa Tesseract + pytesseract
  2) Clipboard - se OCR falhar, le o que esta copiado
  3) Console - se ainda nao conseguir, pergunta no terminal

Requisitos minimos:
  pip install keyboard pillow pyperclip

OCR opcional (deixa tudo automatico):
  1. Baixar Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
  2. Instalar em C:\\Program Files\\Tesseract-OCR\\
  3. pip install pytesseract

IMPORTANTE: Em Windows, talvez precise rodar o terminal como ADMINISTRADOR
para o atalho global funcionar.

Uso:
  python capturar_prova.py
      -> abre o script, aguarde AltGr para capturar, ESC para sair
"""
import sys
import re
import datetime
from pathlib import Path

# -------- imports obrigatorios --------
try:
    import keyboard  # pip install keyboard
except ImportError:
    print("[ERRO] Instale: pip install keyboard")
    sys.exit(1)

try:
    from PIL import ImageGrab
except ImportError:
    print("[ERRO] Instale: pip install pillow")
    sys.exit(1)

# -------- imports opcionais --------
HAS_OCR = False
TESSERACT_PATH = None
try:
    import pytesseract
    # Procura tesseract.exe em paths comuns no Windows
    import shutil
    paths_tentar = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Tesseract-OCR\tesseract.exe",
        r"C:\Users\usuario\AppData\Local\Tesseract-OCR\tesseract.exe",
        r"C:\Users\usuario\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
    ]
    # tambem tenta achar via PATH do sistema
    no_path = shutil.which("tesseract")
    if no_path:
        paths_tentar.insert(0, no_path)
    for p in paths_tentar:
        if p and Path(p).exists():
            pytesseract.pytesseract.tesseract_cmd = p
            TESSERACT_PATH = p
            HAS_OCR = True
            break
    if not HAS_OCR:
        print("[AVISO] pytesseract instalado mas tesseract.exe nao encontrado.")
        print("        Baixe e instale em: https://github.com/UB-Mannheim/tesseract/wiki")
except ImportError:
    pass

HAS_CLIP = False
try:
    import pyperclip
    HAS_CLIP = True
except ImportError:
    pass

# -------- config --------
PASTA_SAIDA = Path("Provas_Envio_Manual")
HOTKEY = "f9"   # F9 do teclado
SUFIXO_NOME = "envio_iptu_2026"


def normalizar_texto(s):
    """UPPERCASE + tira pontos/acentos sutis + espaco vira _, limit 80 chars."""
    if not s:
        return ""
    s = s.strip().upper()
    s = re.sub(r"[<>:\"/\\|?*;,.()]", "", s)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s[:80].strip("_")


def cortar_painel_email(img):
    """
    Corta a screenshot para pegar SO o painel direito do Outlook
    (onde fica o corpo do e-mail aberto). Elimina:
      - Barra lateral de pastas (esquerda)
      - Lista de mensagens (meio) com previews que confundem o OCR
    """
    w, h = img.size
    # Painel direito do Outlook: comeca em ~45-48% da largura
    left = int(w * 0.45)
    return img.crop((left, 0, w, h))


def preprocessar_para_ocr(img):
    """Grayscale + upscaling 2x para melhorar leitura de texto pequeno."""
    from PIL import Image
    img_gray = img.convert("L")
    w, h = img_gray.size
    return img_gray.resize((w * 2, h * 2), Image.LANCZOS)


def extrair_info_via_ocr(img):
    """Extrai QUADRA, LOTE, CCI, NOME do texto OCR da imagem.

    Importante: a `img` recebida ja deve ter sido CROPPED para o painel
    direito do Outlook (corpo do e-mail), nao a tela inteira. Isso evita
    que OCR pegue nomes de previews de outros e-mails da lista lateral.
    """
    if not HAS_OCR:
        return {}
    info = {}

    # Config do tesseract: PSM 6 = assume bloco uniforme de texto (ideal pra
    # corpo de e-mail com paragrafos). lang=por+eng para tratar tanto PT
    # quanto qualquer texto em ingles no header.
    configs_tentar = [
        ("por+eng", "--psm 6"),
        ("por",     "--psm 6"),
        ("por",     "--psm 11"),  # texto esparso
        ("eng",     "--psm 6"),
        (None,      ""),
    ]

    texto = ""
    for lang, cfg in configs_tentar:
        try:
            if lang:
                texto = pytesseract.image_to_string(img, lang=lang, config=cfg)
            else:
                texto = pytesseract.image_to_string(img, config=cfg)
            if texto and len(texto.strip()) > 50:
                break
        except Exception:
            continue

    if not texto:
        print("  [.] OCR retornou texto vazio")
        return {}

    # Debug: mostra trecho do texto pra calibrar
    print(f"  [.] OCR (primeiros 300 chars): {texto[:300].replace(chr(10), ' / ')!r}")

    # QUADRA/LOTE: 52/09 (aceita variacoes com espacos e barras)
    m = re.search(
        r"QUADRA\s*[/\\]?\s*LOTE\s*[:=]?\s*(\d+)\s*[/\\]\s*(\d+)",
        texto, re.I,
    )
    if m:
        info["quadra"] = m.group(1)
        info["lote"] = m.group(2)
    # CCI: 278449
    m = re.search(r"CCI\s*[:=]?\s*(\d{4,8})", texto, re.I)
    if m:
        info["cci"] = m.group(1)
    # Prezado cliente NOME COMPLETO;
    # Restritivo: so aceita letras (incluindo acentos), espaco, hifen, apostrofo.
    # \s* (zero ou mais espacos) porque OCR as vezes lê 'clienteVAUDEIR' grudado.
    m = re.search(
        r"Prezado\s+cliente\s*([A-Za-zÀ-ÿ' \-]{3,80})",
        texto, re.I,
    )
    if m:
        nome = m.group(1)
        # Limpeza forte: tira chars nao-letra do inicio/fim, normaliza espacos
        nome = re.sub(r"^[^A-Za-zÀ-ÿ]+", "", nome)
        nome = re.sub(r"[^A-Za-zÀ-ÿ]+$", "", nome)
        nome = re.sub(r"\s+", " ", nome).strip()
        # Descarta nomes muito curtos ou com 1 palavra so (provavelmente lixo)
        if len(nome) >= 3 and " " in nome:
            info["nome"] = nome
    return info


def ler_clipboard():
    if not HAS_CLIP:
        return ""
    try:
        v = pyperclip.paste() or ""
        return v.strip()[:80]
    except Exception:
        return ""


def perguntar_console():
    print("\n  Nao consegui extrair automaticamente. Digite manualmente:")
    quadra = input("    QUADRA (ex: 52): ").strip()
    lote   = input("    LOTE   (ex: 10): ").strip()
    nome   = input("    NOME completo (ex: Paolo Carneiro De Souza): ").strip()
    return {"quadra": quadra, "lote": lote, "nome": nome}


def montar_nome_arquivo(info):
    partes = []
    if info.get("quadra"):
        partes.append(f"QD{info['quadra'].zfill(2)}")
    if info.get("lote"):
        partes.append(f"LT{str(info['lote']).zfill(3)}")
    if info.get("cci"):
        partes.append(f"CCI{info['cci']}")
    nome_norm = normalizar_texto(info.get("nome", ""))
    if nome_norm:
        partes.append(nome_norm)
    partes.append(SUFIXO_NOME)
    base = "_".join(p for p in partes if p)
    if not base or base == SUFIXO_NOME:
        # fallback total: so timestamp
        base = f"prova_{datetime.datetime.now():%Y%m%d_%H%M%S}_{SUFIXO_NOME}"
    return base + ".pdf"


def capturar(_event=None):
    """Callback chamado quando o hotkey eh pressionado."""
    print("\n" + "=" * 60)
    print(f"[{datetime.datetime.now():%H:%M:%S}] Capturando tela...")
    try:
        # all_screens=False = so monitor principal (mais focado pra OCR)
        img = ImageGrab.grab(all_screens=False)
        print(f"  [.] tela: {img.size[0]}x{img.size[1]} (monitor principal)")
    except Exception as e:
        print(f"  [ERRO] Falha ao capturar tela: {e}")
        return

    info = {}

    # OCR no painel direito do Outlook (corpo do email) ja pre-processado
    if HAS_OCR:
        img_painel = cortar_painel_email(img)
        print(f"  [.] painel cortado: {img_painel.size[0]}x{img_painel.size[1]}")
        img_ocr = preprocessar_para_ocr(img_painel)
        info = extrair_info_via_ocr(img_ocr)
        if info:
            print(f"  [.] OCR achou: {info}")
        else:
            print(f"  [.] OCR nao achou QUADRA/LOTE/Nome no painel direito")
    else:
        print("  [.] OCR indisponivel (Tesseract nao instalado)")

    nome_arq = montar_nome_arquivo(info)

    # Cria subpasta com o nome do comprador (mesmo padrao do baixar_iptu).
    # Se OCR nao achou o nome, salva em 'SEM_NOME/'.
    nome_pasta = normalizar_texto(info.get("nome", "")) or "SEM_NOME"
    subpasta = PASTA_SAIDA / nome_pasta
    subpasta.mkdir(parents=True, exist_ok=True)
    saida = subpasta / nome_arq

    # Se ja existe, adiciona timestamp pra nao sobrescrever
    if saida.exists():
        ts = datetime.datetime.now().strftime("%H%M%S")
        saida = subpasta / f"{saida.stem}_{ts}.pdf"

    try:
        # Pillow salva PDF direto via save(... "PDF", ...)
        img_rgb = img.convert("RGB") if img.mode != "RGB" else img
        img_rgb.save(saida, "PDF", resolution=100.0)
        print(f"  [OK] PDF salvo: {saida.resolve()}")
    except Exception as e:
        print(f"  [ERRO] Falha ao salvar PDF: {e}")

    print(f"\n  >>> Aperte {HOTKEY.upper()} novamente para outra captura <<<")
    print("=" * 60)


def main():
    print("=" * 60)
    print("  Natela Analytics - Captura de Prova de Envio")
    print("=" * 60)
    print(f"  Pasta de saida    : {PASTA_SAIDA.resolve()}")
    print(f"  Atalho de captura : {HOTKEY.upper()}")
    print(f"  Sair              : ESC")
    print("-" * 60)
    print(f"  OCR (Tesseract)   : {'OK' if HAS_OCR else 'NAO instalado'}")
    print(f"  Clipboard         : {'OK' if HAS_CLIP else 'NAO instalado'}")
    if not HAS_OCR:
        print("    -> Para automatizar 100%, instale Tesseract:")
        print("       https://github.com/UB-Mannheim/tesseract/wiki")
        print("       depois: pip install pytesseract")
    if not HAS_CLIP:
        print("    -> Para fallback de clipboard: pip install pyperclip")
    print("=" * 60)
    print(f"\nPronto. Abra o e-mail enviado no Outlook e aperte {HOTKEY.upper()}.")
    print(f"Para sair, aperte ESC.\n")

    try:
        keyboard.add_hotkey(HOTKEY, capturar)
        keyboard.wait("esc")
    except Exception as e:
        print(f"\n[ERRO] {e}")
        print("Talvez precise rodar o terminal como ADMINISTRADOR.")
        sys.exit(1)
    print("\nFinalizado.")


if __name__ == "__main__":
    main()
