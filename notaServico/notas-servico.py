import os
import time
from datetime import datetime
import subprocess
from playwright.sync_api import sync_playwright

# Caminho base dos downloads
base_download_dir = r"\\192.0.0.251\arquivos\XML PREFEITURA"

def salvar_captura_de_tela_declaracao(pagina, caminho, mes, ano):
    nome_arquivo = f"declaracao_sem_movimento_{str(mes).zfill(2)}.{ano}.png"
    caminho_arquivo = os.path.join(caminho, nome_arquivo)
    try:
        pagina.screenshot(path=caminho_arquivo, full_page=True)
        print(f"📸 Captura de tela salva em: {caminho_arquivo}")
    except Exception as e:
        print(f"❗ Erro ao salvar captura de tela: {e}")

def salvar_captura_de_tela(pagina, caminho, mes, ano, sufixo):
    nome_arquivo = f"sem_movimento_{str(mes).zfill(2)}.{ano}_{sufixo}.png"
    caminho_arquivo = os.path.join(caminho, nome_arquivo)
    try:
        pagina.screenshot(path=caminho_arquivo, full_page=True)
        print(f"📸 Captura de tela salva em: {caminho_arquivo}")
    except Exception as e:
        print(f"❗ Erro ao salvar captura de tela: {e}")

def emitir_declaracoes_disponiveis(pagina, nome_prestador, mes, ano, base_download_dir, origem_texto, modo_debug=False):
    while True:
        try:
            linhas = pagina.locator('table tbody tr')
            total_linhas = linhas.count()
            botoes_disponiveis = []

            for i in range(total_linhas):
                linha = linhas.nth(i)
                primeira_coluna = linha.locator("td").nth(0)
                link = primeira_coluna.locator("a")
                if link.count() > 0:
                    try:
                        href = link.get_attribute("href")
                        if href and "emitirDeclaracao" in href:
                            numero_mes = int(href.split("'")[1])
                            botoes_disponiveis.append(numero_mes)
                            if modo_debug:
                                print(f"🔍 Botão disponível para mês {numero_mes}: {href}")
                    except Exception as e:
                        if modo_debug:
                            print(f"⚠️ Erro ao interpretar href: {e}")
                        continue

                # ✅ Nome do cliente limpo (sem CNPJ)
            nome_limpo = nome_prestador.split(" - ", 1)[1] if " - " in nome_prestador else nome_prestador
            pasta_cliente = os.path.join(base_download_dir, nome_limpo.strip())
            pasta_mes_ano = os.path.join(pasta_cliente, f"{str(mes_anterior).zfill(2)}.{ano_ref}")
            os.makedirs(pasta_mes_ano, exist_ok=True)


            if not botoes_disponiveis:
                print("✅ Nenhum botão de declaração disponível. Salvando captura de tela.")
                salvar_captura_de_tela_declaracao(pagina, pasta_mes_ano, mes_anterior, ano_ref)
                time.sleep(3)
                break
            
            for numero_mes in sorted(set(botoes_disponiveis)):
                print(f"🟡 Emitindo declaração do mês {numero_mes}")
                try:
                    pagina.evaluate(f"emitirDeclaracao('{numero_mes}')")
                    print(f"✅ Função emitirDeclaracao('{numero_mes}') executada.")
                except Exception as e:
                    print(f"❗ Erro ao executar emitirDeclaracao('{numero_mes}'): {e}")
                    continue

                pagina.wait_for_timeout(2000)

                try:
                    pagina.click("text=Gravar", timeout=3000)
                    pagina.wait_for_timeout(2000)
                    print(f"✅ Declaração do mês {numero_mes} gravada com sucesso.")
                except TimeoutError:
                    print(f"❗ Botão 'Gravar' não encontrado após o mês {numero_mes}.")

                try:
                    pagina.click("text=Pesquisar")
                    pagina.wait_for_timeout(2000)
                except Exception as e:
                    print(f"❗ Erro ao clicar em 'Pesquisar': {e}")
                    break

                break  # volta ao while

        except Exception as e:
            print(f"❗ Erro inesperado: {e}")
            break

def baixar_arquivos(pagina, nome_prestador, mes_extenso, ano_ref, mes_anterior, origem_texto, tem_registro, index):
    pagina.click("text=Pesquisar")

    try:
        pagina.wait_for_timeout(1500)

        if pagina.is_visible("text=Não há registros"):
            tem_registro = False
        elif pagina.locator("table#tabelaDinamica i.fa.fa-search").count() > 0:
            tem_registro = True                
        else:
            tem_registro = False

    except Exception as e:
        print(f"⚠️ Erro ao verificar movimento: {e}")
        tem_registro = False

    sufixo = "emitido" if origem_texto.lower() == "emitida" else "recebido"

    # ✅ Nome do cliente limpo (sem CNPJ)
    nome_limpo = nome_prestador.split(" - ", 1)[1] if " - " in nome_prestador else nome_prestador
    pasta_cliente = os.path.join(base_download_dir, nome_limpo.strip())
    pasta_mes_ano = os.path.join(pasta_cliente, f"{str(mes_anterior).zfill(2)}.{ano_ref}")
    os.makedirs(pasta_mes_ano, exist_ok=True)

    try:
        if tem_registro:
            with pagina.expect_download(timeout=10000) as download_info:
                pagina.click("text=Exportar em XML")
            download = download_info.value
            nome_arquivo_xml = f"notas_{mes_extenso.lower()}_{ano_ref}_{sufixo}.xml"
        else:
                salvar_captura_de_tela(pagina, pasta_mes_ano, mes_anterior, ano_ref, sufixo)
                
                if sufixo == "emitido":
                    try:

                        print(f"⚠️ Sem registros para {nome_prestador}. Emitindo declaração sem movimento.")
                        pagina.click("text=DECLARAÇÃO")
                        pagina.click("text=Sem movimento")
                        time.sleep(3)
                        pagina.click("text=Pesquisar")
                        pagina.wait_for_timeout(2000)

                        emitir_declaracoes_disponiveis(
                            pagina,
                            nome_prestador,
                            mes_anterior,
                            ano_ref,
                            base_download_dir,
                            origem_texto,
                            modo_debug=True
                        )

                        pagina.click("text=NFS-E")
                        pagina.click("text=Consulta")
                        time.sleep(1)

                        prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
                        prestador_select.select_option(index=index)

                        pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]')
                        pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', label=mes_extenso)
                        pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))


                    except Exception as e:
                        print(f"❌ Erro ao executar declaracao.py: {e}")

                # Defina o nome do arquivo XML para manter a execução
                nome_arquivo_xml = f"sem_movimento_{mes_extenso.lower()}_{ano_ref}_{sufixo}.xml"
                download = None  # Nenhum download neste caso
        
        caminho_final_xml = os.path.join(pasta_mes_ano, nome_arquivo_xml)
        if download:
            download.save_as(caminho_final_xml)
            print(f"✅ XML salvo em:\n{caminho_final_xml}")
        else:
            print(f"ℹ️ Nenhum XML gerado para {sufixo.upper()} ({nome_prestador})")

        return tem_registro, pasta_mes_ano
        
    except Exception as e:
        print(f"⚠️ Falha ao exportar XML ({sufixo}): {e}")
        return False, pasta_mes_ano

def baixar_relatorio(pagina, nome_prestador, mes_extenso, ano_ref, mes_anterior, pasta_mes_ano, tem_registro_emitida, tem_registro_recebida):
    try:
        if tem_registro_emitida:
            pagina.select_option('select[name="formulario.tpOrigemNfs"]', label="Emitida")
            pagina.select_option('select[name="formulario.nrMesCompetencia"]', label=str(mes_anterior))
            pagina.select_option('select[name="formulario.nrAnoCompetencia"]', str(ano_ref))
            pagina.locator('input[value="Pesquisar"]').click()

            pagina.wait_for_timeout(2000)
            print("Mes anterior: ",mes_anterior)
            try:
                with pagina.expect_download(timeout=10000) as download_info:
                    pagina.locator('span.pdf.fa.fa-file-pdf-o').click()
                download_pdf = download_info.value
                nome_arquivo_pdf = f"notas_{mes_extenso.lower()}_{ano_ref}_emitido.pdf"
                caminho_final_pdf = os.path.join(pasta_mes_ano, nome_arquivo_pdf)
                download_pdf.save_as(caminho_final_pdf)
                print(f"✅ PDF EMITIDA salvo em:\n{caminho_final_pdf}")
                
            except Exception as e:
                print(f"⚠️ Falha ao exportar PDF Emitida: {e}")
                
            pagina.click("text=Limpar")
            time.sleep(0.5)
        
        print("Mes anterior: ",mes_anterior)

        if tem_registro_recebida:
            pagina.select_option('select[name="formulario.tpOrigemNfs"]', label="Recebida")
            pagina.select_option('select[name="formulario.nrMesCompetencia"]', label=str(mes_anterior))
            pagina.select_option('select[name="formulario.nrAnoCompetencia"]', str(ano_ref))
            pagina.locator('input[value="Pesquisar"]').click()

            pagina.wait_for_timeout(2000)
            try:
                with pagina.expect_download(timeout=10000) as download_info:
                    pagina.locator('span.pdf.fa.fa-file-pdf-o').click()
                download_pdf = download_info.value
                nome_arquivo_pdf = f"notas_{mes_extenso.lower()}_{ano_ref}_recebido.pdf"
                caminho_final_pdf = os.path.join(pasta_mes_ano, nome_arquivo_pdf)
                download_pdf.save_as(caminho_final_pdf)
                print(f"✅ PDF RECEBIDA salvo em:\n{caminho_final_pdf}")
            except Exception as e:
                print(f"⚠️ Falha ao exportar PDF Recebida: {e}")
                
            pagina.click("text=Limpar")
            time.sleep(0.5)

    except Exception as e:
        print(f"⚠️ Erro ao baixar relatórios: {e}")

    # 🔄 Volta para tela de consulta e aguarda corretamente
    try:
        pagina.click("text=NFS-E")
        pagina.click("text=Consulta")
        
        pagina.wait_for_function(
            """() => {
                const select = document.querySelector('select[name="parametrosTela.idPessoa"]');
                return select && select.options.length > 1;
            }""",
            timeout=15000
        )

        
        pagina.wait_for_selector('select[name="parametrosTela.idPessoa"]')
        # Após navegação, é necessário re-obter o seletor
        prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
        prestador_select.select_option(index=index)

        pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]')
        pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', label=mes_extenso)
        pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))

        time.sleep(1)

    except Exception as e:
        print(f"❗ Erro ao retornar para tela de consulta: {e}")

with sync_playwright() as p:
    navegador = p.chromium.launch(channel="chrome", headless=False)
    contexto = navegador.new_context(accept_downloads=True)
    pagina = contexto.new_page()

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

    pagina.click("text=NFS-E")
    pagina.click("text=Consulta")
    time.sleep(1)

    pagina.wait_for_selector('select[name="parametrosTela.idPessoa"]')
    prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
    prestadores = prestador_select.locator("option").all()
    total_prestadores = len(prestadores)

    if total_prestadores < 2:
        raise Exception("❌ Nenhum prestador válido encontrado.")

    hoje = datetime.today()
    mes_anterior = hoje.month - 1 or 12
    ano_ref = hoje.year if hoje.month > 1 else hoje.year - 1
    meses_ext = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_extenso = meses_ext[mes_anterior - 1]

    for index in range(1, total_prestadores):
        nome_prestador_completo = prestadores[index].text_content().strip()
        print(f"\n🔍 Processando prestador: {nome_prestador_completo}")

        # Limpar antes de buscar novo
        pagina.click("text=Limpar")
        time.sleep(0.5)

        prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
        prestador_select.select_option(index=index)

        pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]')
        pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', label=mes_extenso)
        pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))

        # Emitidas
        pagina.select_option('select[name="parametrosTela.origemEmissaoNfse"]', label="Emitida")
        tem_registro_emitida, pasta_mes_ano = baixar_arquivos(pagina, nome_prestador_completo, mes_extenso, ano_ref, mes_anterior, "Emitida", True, index)
        # Recebidas
        pagina.select_option('select[name="parametrosTela.origemEmissaoNfse"]', label="Recebida")
        tem_registro_recebida, pasta_mes_ano = baixar_arquivos(pagina, nome_prestador_completo, mes_extenso, ano_ref, mes_anterior, "Recebida", True, index)

        if tem_registro_emitida == True or tem_registro_recebida == True:
            pagina.click("text=RELATÓRIOS")
            pagina.click("text=Apuração do ISS")
            time.sleep(1)

            prestador_select = pagina.locator('select[name="formulario.idPessoa"]')
            prestador_select = pagina.locator('select[name="formulario.idPessoa"]')
            prestador_select.select_option(index=index)


            pagina.wait_for_selector('select[name="formulario.nrMesCompetencia"]')
            pagina.select_option('select[name="formulario.nrMesCompetencia"]', label=str(mes_anterior))
            pagina.select_option('select[name="formulario.nrAnoCompetencia"]', str(ano_ref))
            
            baixar_relatorio(
                pagina,
                nome_prestador_completo,
                mes_extenso,
                ano_ref,
                mes_anterior,
                pasta_mes_ano,
                tem_registro_emitida,
                tem_registro_recebida
            )


    # Caminho para o script principal.py
    CAMINHO_PRINCIPAL = r"C:\Users\Usuario\Documents\PYTHON\IMPORTADOR_NFSE\principal.py"

    print("\n🚀 Executando principal.py...")
    subprocess.run(["python", CAMINHO_PRINCIPAL], check=True)