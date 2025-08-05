import os
import time
from datetime import datetime
from playwright.sync_api import sync_playwright

# Diretório base dos downloads
BASE_DOWNLOAD_DIR = r"\\192.0.0.251\arquivos\XML PREFEITURA\GUIAS"

# Utilitário: monta o caminho final do PDF
def montar_caminho_download(nome_prestador, mes, ano, mes_extenso):
    nome_limpo = nome_prestador.split(" - ", 1)[1] if " - " in nome_prestador else nome_prestador
    pasta_cliente = os.path.join(BASE_DOWNLOAD_DIR, nome_limpo.strip())
    pasta_mes_ano = os.path.join(pasta_cliente, f"{str(mes).zfill(2)}.{ano}")
    os.makedirs(pasta_mes_ano, exist_ok=True)
    nome_arquivo_pdf = f"ISS_{nome_prestador}_{mes_extenso.lower()}_{ano}_emitido.pdf"
    caminho_final_pdf = os.path.join(pasta_mes_ano, nome_arquivo_pdf)
    return caminho_final_pdf

# Emite todas as guias disponíveis para um prestador
def emitir_guias(pagina, contexto, nome_prestador, mes, ano, mes_extenso):
    contador_guia = 1  # contador local de guias emitidas

    while True:
        links_emissao = pagina.locator('a[title="Emissão"]')
        total_links = links_emissao.count()

        if total_links == 0:
            print("⚠️ Nenhuma guia para emitir.")
            break

        try:
            link = links_emissao.nth(0)
            href = link.get_attribute("href")

            if href and "viewEditGuia" in href:
                # Etapa 1: clicar no ícone barcode
                link.locator("i.fa.fa-barcode").click()
                print("🔘 Clicando no botão de emissão...")

                # Etapa 2: clicar em "Emitir"
                pagina.wait_for_selector("input#emitir", timeout=5000)
                pagina.locator("input#emitir").scroll_into_view_if_needed()
                pagina.wait_for_timeout(300)

                # Etapa 3: clicar no botão, depois aceitar o popup, depois esperar nova aba
                with contexto.expect_page() as nova_guia_info:
                    pagina.locator("input#emitir").click()
                    pagina.once("dialog", lambda dialog: dialog.accept())

                # Etapa 4: nova aba é aberta após o popup
                nova_pagina = nova_guia_info.value
                nova_pagina.wait_for_url("**/nfsguiarecolhimento.imprimir.logic", timeout=10000)
                nova_pagina.wait_for_load_state("load")
                print("📄 PDF carregado. Tentando fazer download...")

                # Etapa 5: download do PDF no visualizador
                with nova_pagina.expect_download(timeout=10000) as download_info:
                    nova_pagina.locator('button[aria-label="Download"]').click()

                download_pdf = download_info.value

                caminho_final_pdf = montar_caminho_download(
                    nome_prestador,
                    mes,
                    ano,
                    mes_extenso
                )

                # Adiciona sufixo incremental
                base, ext = os.path.splitext(caminho_final_pdf)
                caminho_final_pdf = f"{base}_{contador_guia}{ext}"
                contador_guia += 1

                download_pdf.save_as(caminho_final_pdf)
                print(f"✅ PDF salvo em:\n{caminho_final_pdf}")

                # Etapa 6: fecha nova aba, volta para a principal
                nova_pagina.close()
                pagina.bring_to_front()

                # Etapa 7: clica em "Voltar", pesquisa de novo
                try:
                    pagina.click('input.botaoVoltar')  # botão Voltar
                    pagina.wait_for_load_state("networkidle")
                    time.sleep(1)
                    pagina.click("text=Pesquisar")
                    pagina.wait_for_timeout(1500)
                except Exception as e:
                    print("⚠️ Erro ao retornar para a pesquisa:", e)
                    break

                # Etapa 8: verifica se há mais guias
                if pagina.is_visible("text=Não há registros"):
                    print("✅ Nenhuma nova guia encontrada. Passando para o próximo prestador.")
                    break

        except Exception as e:
            print(f"❗ Erro inesperado ao emitir guia: {e}")
            break

# Processa todos os prestadores
def processar_prestadores(pagina, contexto):
    # Datas de referência
    hoje = datetime.today()
    mes_anterior = hoje.month - 1 if hoje.month > 1 else 12
    ano_ref = hoje.year if hoje.month > 1 else hoje.year - 1
    meses_ext = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_extenso = meses_ext[mes_anterior - 1]

    prestador_select = pagina.locator('select[name="pessoaModel.idPessoa"]')
    prestadores = prestador_select.locator("option").all()
    total_prestadores = len(prestadores)

    if total_prestadores < 2:
        raise Exception("❌ Nenhum prestador válido encontrado.")

    for index in range(64, total_prestadores):
        nome_prestador = prestadores[index].text_content().strip()
        print(f"\n🔍 Processando prestador: {nome_prestador}")

        prestador_select.select_option(index=index)
        pagina.wait_for_selector('select[name="formulario.nrExercicio"]')
        pagina.select_option('select[name="formulario.nrExercicio"]', str(ano_ref))
        pagina.click("text=Pesquisar")
        pagina.wait_for_timeout(1500)

        if pagina.is_visible("text=Não há registros"):
            print("❌ Nenhum registro encontrado.")
            continue

        try:
            links_emissao = pagina.locator('a[title="Emissão"]')
            total_links = links_emissao.count()

            if total_links > 0:
                print(f"✅ {total_links} guia(s) localizada(s)")
                emitir_guias(pagina, contexto, nome_prestador, mes_anterior, ano_ref, mes_extenso)
            else:
                print("⚠️ Nenhum link de emissão encontrado.")
        except Exception as e:
            print(f"⚠️ Erro ao processar registros: {e}")

# Fluxo principal
def main():
    with sync_playwright() as p:
        navegador = p.chromium.launch(channel="chrome", headless=False)
        contexto = navegador.new_context(accept_downloads=True)
        pagina = contexto.new_page()

        # Navegação inicial
        pagina.goto("https://www.esnfs.com.br/?e=35")
        time.sleep(2)

        pagina.wait_for_selector("text=Certificado digital")
        pagina.click("text=Certificado digital")
        pagina.wait_for_load_state("networkidle")
        time.sleep(1)

        pagina.wait_for_selector("text=Município de Francisco Beltrão")
        pagina.click("text=Município de Francisco Beltrão")
        pagina.wait_for_load_state("networkidle")
        time.sleep(1)

        pagina.click("text=GUIA DE RECOLHIMENTO")
        pagina.click("text=ISS devido / Consulta / Cancelamento")
        time.sleep(1)

        processar_prestadores(pagina, contexto)

        input("\n🛑 Pressione ENTER para encerrar manualmente...")

if __name__ == "__main__":
    main()
