# Automação WIMS

## Descrição
Automação desenvolvida em Python para coleta de VRIDs no WIMS e consolidação de dados em Excel.

## Funcionalidades
- Leitura automática de VRIDs via interface
- Extração de dados via automação de navegador (Playwright)
- Cruzamento com base externa
- Geração de relatório automatizado

## Tecnologias
- Python
- Playwright
- Regex
- Openpyxl

## Benefício
Redução de trabalho manual e aumento da eficiência operacional.# Leitura-wins_brlog
Automação em Python que coleta VRIDs do WIMS, cruza dados com base externa e gera relatório automatizado em Excel.


import re
import os
from typing import List, Dict
from playwright.sync_api import sync_playwright
from openpyxl import Workbook


CDP_URL = "http://127.0.0.1:9222"
ARQUIVO_SAIDA = r"C:\Users\jvilelad\Desktop\codigos\resultado_automacao_final.xlsx"
PASTA_DEBUG = r"C:\Users\jvilelad\Desktop\codigos\debug_brlog"


def garantir_pasta_debug():
    os.makedirs(PASTA_DEBUG, exist_ok=True)


def resetar_excel(caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        os.remove(caminho_arquivo)

    wb = Workbook()
    ws = wb.active
    ws.title = "Consulta VRIDs"
    ws.append([
        "VRID",
        "Previsão de entrega",
        "Fim recalculado",
        "Distância do cliente",
        "Status Automação"
    ])
    wb.save(caminho_arquivo)
    print("Excel reiniciado:", caminho_arquivo)


def adicionar_linha_excel(resultado: Dict[str, str], caminho_arquivo):
    from openpyxl import load_workbook

    wb = load_workbook(caminho_arquivo)
    ws = wb["Consulta VRIDs"]

    ws.append([
        str(resultado.get("VRID", "")),
        str(resultado.get("Previsão de entrega", "")),
        str(resultado.get("Fim recalculado", "")),
        str(resultado.get("Distância do cliente", "")),
        str(resultado.get("Status Automação", ""))
    ])

    wb.save(caminho_arquivo)


def encontrar_pagina(pages, trecho_url):
    for page in pages:
        try:
            if trecho_url in page.url:
                return page
        except Exception:
            pass
    return None


def extrair_vrids_do_wims(texto: str) -> List[str]:
    encontrados = re.findall(r"\bfor\s+([A-Z0-9]{9})\b", texto, flags=re.IGNORECASE)

    vrids = []
    vistos = set()

    for vrid in encontrados:
        vrid = vrid.upper().strip()
        if vrid not in vistos:
            vistos.add(vrid)
            vrids.append(vrid)

    return vrids


def ler_vrids_wims(wims_page) -> List[str]:
    print("Lendo VRIDs do WIMS...")
    wims_page.bring_to_front()
    wims_page.wait_for_timeout(3000)

    # espera a tabela aparecer minimamente
    wims_page.locator("body").wait_for(timeout=15000)
    texto = wims_page.locator("body").inner_text(timeout=15000)

    caminho_txt = os.path.join(PASTA_DEBUG, "wims_texto_lido.txt")
    with open(caminho_txt, "w", encoding="utf-8") as f:
        f.write(texto)

    vrids = extrair_vrids_do_wims(texto)

    print(f"VRIDs encontrados: {len(vrids)}")
    for i, vrid in enumerate(vrids, start=1):
        print(f"{i}. {vrid}")

    return vrids


def clicar_lupa_inicial(brlog_page):
    brlog_page.locator("#botaosearch").click()
    brlog_page.wait_for_timeout(800)


def preencher_vrid(brlog_page, vrid):
    campo = brlog_page.locator('input[placeholder="SML ou Vrid"]')
    campo.wait_for(state="visible", timeout=10000)
    campo.click()
    campo.press("Control+A")
    campo.press("Backspace")
    campo.type(vrid, delay=120)
    campo.press("Tab")
    brlog_page.wait_for_timeout(500)

    valor = campo.input_value().strip()
    print("Valor no campo:", valor)

    if valor != vrid:
        raise Exception(f"Campo preenchido incorretamente. Esperado={vrid} | Atual={valor}")


def clicar_seta(brlog_page):
    brlog_page.locator("#icone-btn-pesq").click(force=True)
    brlog_page.wait_for_timeout(5000)


def capturar_texto_visivel(brlog_page, vrid):
    texto = brlog_page.locator("body").inner_text()
    caminho_txt = os.path.join(PASTA_DEBUG, f"texto_pos_pesquisa_{vrid}.txt")
    with open(caminho_txt, "w", encoding="utf-8") as f:
        f.write(texto)
    return texto


def extrair_por_texto(texto):
    texto = texto.replace("\xa0", " ")
    texto = texto.replace("\r", "")
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n+", "\n", texto)

    bloco_match = re.search(
        r"Entregas para esta viagem(.*?)(VOLTAR|Voltar|$)",
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )

    # fallback: se não tiver o bloco, usa o texto inteiro
    bloco = bloco_match.group(1) if bloco_match else texto

    datas = re.findall(r"\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}", bloco)
    kms = re.findall(r"[\d.,]+\s*km", bloco, flags=re.IGNORECASE)

    previsao = datas[0].strip() if len(datas) >= 1 else ""
    fim = datas[1].strip() if len(datas) >= 2 else ""
    km = kms[-1].strip() if kms else ""

    if not previsao and not fim and not km:
        raise Exception("Não encontrei dados suficientes no retorno do BRLog.")

    return previsao, fim, km


def clicar_x(brlog_page):
    seletores = [
        "#LimparPesquisa",
        "span:has-text('×')",
        "i.material-icons:has-text('close')",
    ]

    for seletor in seletores:
        try:
            elemento = brlog_page.locator(seletor).first
            if elemento.count() > 0:
                elemento.click(timeout=5000)
                brlog_page.wait_for_timeout(1000)
                return
        except Exception:
            pass


def consultar_brlog(brlog_page, vrid: str) -> Dict[str, str]:
    try:
        brlog_page.bring_to_front()
        brlog_page.wait_for_timeout(1500)

        clicar_lupa_inicial(brlog_page)
        preencher_vrid(brlog_page, vrid)
        clicar_seta(brlog_page)

        texto = capturar_texto_visivel(brlog_page, vrid)
        previsao, fim_recalculado, km = extrair_por_texto(texto)

        clicar_x(brlog_page)

        return {
            "VRID": vrid,
            "Previsão de entrega": previsao,
            "Fim recalculado": fim_recalculado,
            "Distância do cliente": km,
            "Status Automação": "Sucesso"
        }

    except Exception as e:
        try:
            caminho_erro = os.path.join(PASTA_DEBUG, f"erro_{vrid}.png")
            brlog_page.screenshot(path=caminho_erro, full_page=True)
        except Exception:
            pass

        try:
            texto_erro = brlog_page.locator("body").inner_text(timeout=10000)
            caminho_txt_erro = os.path.join(PASTA_DEBUG, f"erro_{vrid}.txt")
            with open(caminho_txt_erro, "w", encoding="utf-8") as f:
                f.write(texto_erro)
        except Exception:
            pass

        try:
            clicar_x(brlog_page)
        except Exception:
            pass

        return {
            "VRID": vrid,
            "Previsão de entrega": "",
            "Fim recalculado": "",
            "Distância do cliente": "",
            "Status Automação": f"Erro: {str(e)[:120]}"
        }


def main():
    garantir_pasta_debug()
    resetar_excel(ARQUIVO_SAIDA)

    with sync_playwright() as p:
        print("Conectando ao Chrome já aberto...")
        browser = p.chromium.connect_over_cdp(CDP_URL)
        print("Conectado com sucesso!\n")

        if not browser.contexts:
            print("Nenhum contexto encontrado.")
            return

        context = browser.contexts[0]
        pages = context.pages

        wims_page = encontrar_pagina(pages, "optimus-internal.amazon.com/wims")
        brlog_page = encontrar_pagina(pages, "amazon.brasilrisk.com.br/Dashboard")

        if not wims_page:
            print("Não encontrei a aba do WIMS.")
            return

        if not brlog_page:
            print("Não encontrei a aba do BRLog.")
            return

        vrids = ler_vrids_wims(wims_page)

        if not vrids:
            print("Nenhum VRID encontrado.")
            return

        print("\nIniciando consultas no BRLog...\n")

        for i, vrid in enumerate(vrids, start=1):
            print(f"[{i}/{len(vrids)}] Consultando VRID {vrid}...")
            resultado = consultar_brlog(brlog_page, vrid)
            adicionar_linha_excel(resultado, ARQUIVO_SAIDA)
            print("Status:", resultado["Status Automação"])
            print("-" * 70)

        print("\nProcesso finalizado.")
        input("Pressione ENTER para encerrar...")


if __name__ == "__main__":
    main()

    
