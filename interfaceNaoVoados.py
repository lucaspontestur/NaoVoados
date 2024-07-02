import os
import sys
import time
import asyncio
from datetime import date
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from tkcalendar import DateEntry
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException
from win10toast import ToastNotifier

# Caminho para o ChromeDriver 
chrome_driver_path = "chromedriver.exe"

#  --- Caminho Absoluto para a Pasta de Downloads --- 
if getattr(sys, 'frozen', False):
    # Caminho quando executado como executável
    script_dir = os.path.dirname(sys.executable)
else:
    # Caminho quando executado como script .py
    script_dir = os.path.dirname(os.path.abspath(__file__))

pasta_downloads = os.path.join(script_dir, "relatorios_nao_voados")

# Cria a pasta "relatorios_nao_voados" se ela não existir
if not os.path.exists(pasta_downloads):
    os.makedirs(pasta_downloads)

async def baixar_relatorio_nao_voado(data_inicio, data_fim, tipo_arquivo):
    """Função para baixar o relatório Não Voados."""

    # Configurar as preferências do Chrome para downloads
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": pasta_downloads,
        "download.prompt_for_download": False,  # Desativa a caixa de diálogo de download
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    # Crie uma instância do serviço do ChromeDriver
    service = Service(chrome_driver_path)

    # Crie uma instância do navegador Chrome com as opções definidas
    driver = webdriver.Chrome(service=service, options=options)

    try:
        # Acesse o site
        driver.get("https://www.argoit.com.br/pontestur/")

        # Aguardar o carregamento da página de login (ajuste o tempo se necessário)
        driver.implicitly_wait(10)

        # --- Login ---
        username_field = driver.find_element(By.CSS_SELECTOR, "#LoginView1_Login1_User")
        username_field.send_keys("inovacaoptt")
        password_field = driver.find_element(By.CSS_SELECTOR, "#LoginView1_Login1_Password")
        password_field.send_keys("Pontes@2024")
        login_button = driver.find_element(By.CSS_SELECTOR, "#LoginView1_Login1_LoginButton")
        login_button.click()

        # Aguardar o carregamento da página principal (ajuste o tempo se necessário)
        driver.implicitly_wait(10)

        # --- Iterar sobre os clientes ---
        for i in range(len(driver.find_elements(By.CSS_SELECTOR, '#efeitoHeader > div.col-md-2.col-xs-2 > select > option'))):
            cliente_select = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#efeitoHeader > div.col-md-2.col-xs-2 > select'))
            )
            cliente_select.click()
            time.sleep(1)  # Aguarda um pouco para o menu dropdown abrir
            opcoes_cliente = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#efeitoHeader > div.col-md-2.col-xs-2 > select > option'))
            )
            opcoes_cliente[i].click()
            time.sleep(5)  # Aguarda o carregamento da página do cliente

            # --- Acessar Relatórios ---
            relatorios_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='conteudo-principal']//span[contains(text(),'Relatórios')]/parent::div/parent::a"))
            )
            relatorios_button.click()

            # Aguarde 2 segundos
            time.sleep(2)

            # Muda o contexto para o iframe
            iframe = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "iframePaginas"))
            )
            driver.switch_to.frame(iframe)

            # Aguarde 2 segundos
            time.sleep(2)
            
            abrir_agendas_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_imbAgenda"))
            )
            abrir_agendas_button.click()

            # Aguarde 2 segundos
            time.sleep(2)

            relatorios_nao_voados_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_gdvReports > tbody > tr:nth-child(8) > td:nth-child(1) > input[type=image]"))
            )
            relatorios_nao_voados_button.click()

            # Aguarde 2 segundos
            time.sleep(2)

            # --- Inserir Datas ---
            for tentativa in range(3):  # Tenta até 3 vezes
                try:
                    campo_data_inicio = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_txtDataIniA"))
                    )
                    campo_data_inicio.click()
                    time.sleep(2)
                    campo_data_inicio.send_keys(data_inicio.strftime("%d/%m/%Y"))
                    time.sleep(1)  # Aguarda um pouco para garantir que a data foi inserida
                    break  # Sai do loop se a interação for bem-sucedida
                except StaleElementReferenceException:
                    print("Elemento 'Data Início' desatualizado. Tentando novamente...")
                    time.sleep(1)  # Pequena pausa antes de tentar novamente

            for tentativa in range(3):  # Tenta até 3 vezes
                try:
                    campo_data_fim = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_txtDataFinA"))
                    )
                    campo_data_fim.click()
                    time.sleep(2)
                    campo_data_fim.send_keys(data_fim.strftime("%d/%m/%Y"))
                    time.sleep(1)  # Aguarda um pouco para garantir que a data foi inserida
                    break  # Sai do loop se a interação for bem-sucedida
                except StaleElementReferenceException:
                    print("Elemento 'Data Fim' desatualizado. Tentando novamente...")
                    time.sleep(1)  # Pequena pausa antes de tentar novamente


            # Aguarde 2 segundos
            time.sleep(2)

            # --- Selecionar Tipo de Arquivo ---
            tipo_saida = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_lstTipo"))
            )
            tipo_saida.click()
            
            if tipo_arquivo == "PDF":
                opcao_tipo = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_lstTipo > option:nth-child(1)"))
                )
            else:
                opcao_tipo = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_lstTipo > option:nth-child(2)"))
                )
            opcao_tipo.click()

            # --- Baixar Relatório ---
            # Encontrar todos os botões de download na tabela
            botoes_download = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#tblRepeater img[id*='_imgStatus']"))
            )

            # Clicar no último botão da lista
            if botoes_download:
                ultimo_botao = botoes_download[-1]  # Pega o último elemento da lista
                ultimo_botao.click()
            else:
                print("Nenhum botão de download encontrado.")

            # --- Lidar com o Alerta ---
            try:
                # Aguarda o alerta aparecer (ajuste o tempo limite se necessário)
                WebDriverWait(driver, 10).until(EC.alert_is_present()) 

                # Aceita o alerta
                alerta = driver.switch_to.alert
                alerta.accept()
                print("Alerta aceito.")

            except TimeoutException:
                print("Alerta não encontrado. Prosseguindo...")

            # Aguardar o download ser concluído (ajuste o tempo limite conforme necessário)
            aguardar_download(driver, 30)  # Aguarda até 30 segundos

            # --- Renomear Arquivo ---
            nome_cliente = opcoes_cliente[i].text
            renomear_arquivo(tipo_arquivo, nome_cliente, data_inicio, data_fim)  # Passa as datas para a função

            # --- Voltar para a página de Relatórios ---
            driver.back()  # Volta para a página anterior (Relatórios)
            time.sleep(3)  # Aguarda um pouco para a página carregar

        # --- Fim do loop dos clientes ---
        notificar_conclusao()

    except NoSuchElementException as e:
        print(f"Erro: Elemento não encontrado - {str(e)}")
        # Lógica adicional para lidar com o erro, se necessário
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        # Lógica adicional para lidar com o erro, se necessário
    finally:
        driver.quit()

def renomear_arquivo(tipo_arquivo, nome_cliente, data_inicio, data_fim):
    """Função para renomear o arquivo baixado."""
    extensao = ".pdf" if tipo_arquivo == "PDF" else ".xls"
    nome_arquivo_esperado = f"Relatorio{extensao}"  # Nome padrão do arquivo baixado

    # Aguarda o arquivo aparecer na pasta de downloads (ajuste o tempo limite se necessário)
    tempo_limite = 30
    tempo_inicio = time.time()
    while time.time() - tempo_inicio <= tempo_limite:
        for arquivo in os.listdir(pasta_downloads):
            if arquivo == nome_arquivo_esperado:
                # Formata as datas para o nome do arquivo
                data_inicio_formatada = data_inicio.strftime("%d-%m-%Y")
                data_fim_formatada = data_fim.strftime("%d-%m-%Y")

                nome_antigo = os.path.join(pasta_downloads, arquivo)
                nome_novo = os.path.join(pasta_downloads, f"{nome_cliente} {data_inicio_formatada} {data_fim_formatada}{extensao}")
                os.rename(nome_antigo, nome_novo)
                print(f"Arquivo renomeado para: {nome_novo}")
                return

        time.sleep(1)  # Aguarda 1 segundo antes de verificar novamente

    print(f"Erro: Arquivo '{nome_arquivo_esperado}' não encontrado na pasta de downloads.")


def notificar_conclusao(titulo="Automação Finalizada", mensagem="Download dos Relatórios Não Voados concluído!"):
    """Exibe uma notificação do Windows."""
    try:
        toaster = ToastNotifier()
        toaster.show_toast(titulo, mensagem, duration=10, threaded=True)
    except Exception as e:
        print(f"Erro ao exibir notificação: {e}")

def iniciar_download():
    """Função para obter as datas e tipo de arquivo da interface e iniciar o download."""
    data_inicio = calendario_inicio.get_date()
    data_fim = calendario_fim.get_date()
    tipo_arquivo = tipo_arquivo_var.get()

    asyncio.run(baixar_relatorio_nao_voado(data_inicio, data_fim, tipo_arquivo))

def aguardar_download(driver, timeout):
    """Aguarda o download ser concluído dentro do tempo limite."""
    driver.execute_script("window.open()")
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[1])
    driver.get("about:downloads")

    tempo_inicio = time.time()
    while time.time() - tempo_inicio <= timeout:
        try:
            # Obtém a referência ao elemento a cada iteração
            download_porcentagem = driver.execute_script(
                "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value"
            )
            if download_porcentagem == 100:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                return True
        except Exception as e:
            print(f"Erro na verificação do download: {e}")  # Imprime o erro para depuração
            pass  # Ignora o erro e continua verificando
        time.sleep(1)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    print("Download não concluído no tempo limite.")
    return False

# --- Interface Gráfica (CustomTkinter) ---
janela = ctk.CTk()
janela.title("Relatórios Não Voados")
janela.geometry("400x250")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# --- Widgets da Interface ---
label_data_inicio = ctk.CTkLabel(janela, text="Data de Início:")
label_data_inicio.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

calendario_inicio = DateEntry(janela, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
calendario_inicio.grid(row=0, column=1, padx=10, pady=10)

label_data_fim = ctk.CTkLabel(janela, text="Data de Fim:")
label_data_fim.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

calendario_fim = DateEntry(janela, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
calendario_fim.set_date(date.today())
calendario_fim.grid(row=1, column=1, padx=10, pady=10)

tipo_arquivo_var = ctk.StringVar(value="PDF")

radiobutton_pdf = ctk.CTkRadioButton(janela, text="PDF", variable=tipo_arquivo_var, value="PDF")
radiobutton_pdf.grid(row=2, column=0, padx=10, pady=5)

radiobutton_excel = ctk.CTkRadioButton(janela, text="EXCEL", variable=tipo_arquivo_var, value="EXCEL")
radiobutton_excel.grid(row=2, column=1, padx=10, pady=5)

botao_baixar = ctk.CTkButton(janela, text="Baixar Relatórios", command=iniciar_download)
botao_baixar.grid(row=3, column=0, columnspan=2, padx=10, pady=20)

janela.mainloop()