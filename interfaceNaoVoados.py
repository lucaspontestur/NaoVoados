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
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
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
    prefs = {"download.default_directory": pasta_downloads}
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

            # --- Acessar Relatórios Não Voados ---
            iframe = WebDriverWait(driver, 20).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "iframePaginas"))
            )
            
            abrir_agendas_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_imbAgenda"))
            )
            abrir_agendas_button.click()

            relatorios_nao_voados_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_gdvReports > tbody > tr:nth-child(8) > td:nth-child(1) > input[type=image]"))
            )
            relatorios_nao_voados_button.click()

            # --- Inserir Datas ---
            campo_data_inicio = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_txtDataIniA"))
            )
            campo_data_inicio.send_keys(data_inicio.strftime("%d/%m/%Y"))

            campo_data_fim = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_txtDataFinA"))
            )
            campo_data_fim.send_keys(data_fim.strftime("%d/%m/%Y"))

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
            baixar_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_rptRelatorios_ctl01_imgStatus"))
            )
            baixar_button.click()
            time.sleep(5)  # Aguarda o download

            # --- Renomear Arquivo ---
            nome_cliente = opcoes_cliente[i].text
            renomear_arquivo(tipo_arquivo, nome_cliente)

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

def renomear_arquivo(tipo_arquivo, nome_cliente):
    """Função para renomear o arquivo baixado."""
    extensao = ".pdf" if tipo_arquivo == "PDF" else ".xls"
    for arquivo in os.listdir(pasta_downloads):
        if arquivo.startswith("Relatorio") and arquivo.endswith(extensao):
            nome_antigo = os.path.join(pasta_downloads, arquivo)
            nome_novo = os.path.join(pasta_downloads, f"Relatório Não Voados ({nome_cliente}){extensao}")
            os.rename(nome_antigo, nome_novo)
            break

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