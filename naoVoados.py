import os
import sys
import time
from datetime import datetime
from win10toast import ToastNotifier
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import Select
from collections import defaultdict
import configparser

# Caminho para o ChromeDriver 
chrome_driver_path = "chromedriver.exe"

#  --- Caminho Absoluto para a Pasta de Downloads --- 
if getattr(sys, 'frozen', False):
    # Caminho quando executado como executável
    script_dir = os.path.dirname(sys.executable)
else:
    # Caminho quando executado como script .py
    script_dir = os.path.dirname(os.path.abspath(__file__))

pasta_downloads = os.path.join(script_dir, "relatorios")

# Cria a pasta "relatorios" se ela não existir
if not os.path.exists(pasta_downloads):
    os.makedirs(pasta_downloads)

# Dicionário para rastrear erros repetidos
erros_repetidos = defaultdict(lambda: {"contagem": 0, "cliente": ""})

async def executar_acao_por_cliente(data_inicio, data_fim, tipo_arquivo, apenas_agendar=False, apenas_baixar=False):
    """Executa a ação (agendar ou baixar) para cada cliente."""
    global erros_repetidos

    # Carregar configurações
    config = configparser.ConfigParser()
    config.read('configuracoes.ini')
    cliente_inicial = config.get('INICIO', 'nao_voados', fallback="")

    # Configurar as preferências do Chrome para downloads
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": pasta_downloads}
    options.add_experimental_option("prefs", prefs)

    # Crie uma instância do serviço do ChromeDriver
    service = Service(chrome_driver_path)

    # Crie uma instância do navegador Chrome com as opções definidas
    driver = webdriver.Chrome(service=service, options=options)

    # Definir o nome do arquivo de log com base na ação
    if apenas_agendar:
        nome_arquivo_log = 'log_agendamento_nao_voados.txt'
    elif apenas_baixar:
        nome_arquivo_log = 'log_download_nao_voados.txt'
    else:
        nome_arquivo_log = 'log_processo_completo_nao_voados.txt'

    # Redirecionar a saída padrão para o arquivo de log
    original_stdout = sys.stdout
    with open(nome_arquivo_log, 'w') as f:
        sys.stdout = f
        try:
            # Acesse o site
            print("Acessando o site...")
            driver.get("https://www.argoit.com.br/pontestur/")

            # Fazer login no sistema
            print("Inserindo usuário...")
            campo_usuario = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "#LoginView1_Login1_User"))
            )
            campo_usuario.send_keys("inovacaoptt")
            time.sleep(2)

            print("Inserindo senha...")
            campo_senha = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "#LoginView1_Login1_Password"))
            )
            campo_senha.send_keys("Pontes@2024")
            time.sleep(2)

            print("Logando...")
            botao_acessar = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#LoginView1_Login1_LoginButton"))
            )
            botao_acessar.click()
            time.sleep(2)

            print("Login realizado com sucesso!")
            # Aguardar o carregamento da página
            time.sleep(5)

            # Iterar sobre cada cliente
            print("Loop iniciado...")
            for i in range(len(driver.find_elements(By.CSS_SELECTOR, '#efeitoHeader > div.col-md-2.col-xs-2 > select > option'))):
                try:
                    # Selecionar o cliente
                    select_cliente = Select(driver.find_element(By.CSS_SELECTOR, '#efeitoHeader > div.col-md-2.col-xs-2 > select'))
                    nome_cliente = select_cliente.options[i].text

                    # Verificar se há um cliente inicial definido
                    if cliente_inicial and nome_cliente != cliente_inicial:
                        print(f"Pulando cliente: {nome_cliente} (aguardando cliente inicial: {cliente_inicial})")
                        continue

                    cliente_inicial = ""  # Reinicia a variável após encontrar o cliente inicial

                    select_cliente.select_by_index(i)
                    print(f"Processando cliente: {nome_cliente}...")
                    time.sleep(3)

                    # Clicar em "Relatórios"
                    print(f"Acessando relatórios de {nome_cliente}...")
                    botao_relatorios = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@class='conteudo-principal']//span[contains(text(),'Relatórios')]/parent::div/parent::a"))
                    )
                    botao_relatorios.click()
                    time.sleep(3)

                    # Mudar para o iframe
                    print(f"Mudando para o iframe de {nome_cliente}...")
                    iframe = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.ID, "iframePaginas"))
                    )
                    driver.switch_to.frame(iframe)
                    time.sleep(3)

                    # Clicar em "Abrir Agendas"
                    print(f"Abrindo agendas de {nome_cliente}...")
                    botao_abrir_agendas = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_imbAgenda"))
                    )
                    botao_abrir_agendas.click()
                    time.sleep(3)

                    # Clicar em "Relatórios Não Voados"
                    print(f"Acessando relatórios não voados de {nome_cliente}...")
                    botao_nao_voados = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_gdvReports > tbody > tr:nth-child(8) > td:nth-child(1) > input[type=image]"))
                    )
                    botao_nao_voados.click()
                    time.sleep(3)

                    # --- Inserir Datas ---
                    print(f"Inserindo datas para {nome_cliente}...")
                    for tentativa in range(3):  # Tenta até 3 vezes
                        try:
                            campo_data_inicio = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_txtDataIniA"))
                            )
                            campo_data_inicio.click()
                            time.sleep(2)
                            campo_data_inicio.send_keys(data_inicio.strftime("%d/%m/%Y"))
                            time.sleep(2)
                            break
                        except StaleElementReferenceException:
                            print("Elemento 'Data Início' desatualizado. Tentando novamente...")
                            time.sleep(2)

                    for tentativa in range(3):
                        try:
                            campo_data_fim = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_txtDataFinA"))
                            )
                            campo_data_fim.click()
                            time.sleep(2)
                            campo_data_fim.send_keys(data_fim.strftime("%d/%m/%Y"))
                            time.sleep(2)
                            break
                        except StaleElementReferenceException:
                            print("Elemento 'Data Fim' desatualizado. Tentando novamente...")
                            time.sleep(2)

                    time.sleep(2)

                    # Selecionar tipo de arquivo
                    print(f"Selecionando tipo de arquivo para {nome_cliente}...")
                    select_tipo_saida = Select(driver.find_element(By.CSS_SELECTOR, "#ctl00_cphContent_lstTipo"))
                    if tipo_arquivo == "PDF":
                        select_tipo_saida.select_by_index(0)
                    elif tipo_arquivo == "EXCEL":
                        select_tipo_saida.select_by_index(1)
                    time.sleep(3)

                    # --- Clicar em "Agendar" (se apenas_baixar for False) ---
                    if not apenas_baixar:
                        print(f"Agendando relatório para {nome_cliente}...")
                        botao_agendar = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "#ctl00_cphContent_btnSetAgenda"))
                        )
                        botao_agendar.click()
                        time.sleep(3)

                    # --- Baixar Relatório (se apenas_agendar for False) ---
                    if not apenas_agendar:
                        print(f"Baixando relatório de {nome_cliente}...")
                        # Rolar a tela para baixo
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(2)

                        # Obter a lista de arquivos antes do download
                        arquivos_antes = os.listdir(pasta_downloads)

                        # Encontrar todos os botões de download na tabela
                        botoes_download = WebDriverWait(driver, 20).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#tblRepeater img[id*='_imgStatus']"))
                        )

                        # Clicar no último botão da lista
                        if botoes_download:
                            ultimo_botao = botoes_download[-1]
                            ultimo_botao.click()
                            time.sleep(2)
                        else:
                            print("Nenhum botão de download encontrado.")

                        # --- Lidar com o Alerta ---
                        try:
                            WebDriverWait(driver, 10).until(EC.alert_is_present())
                            alerta = driver.switch_to.alert
                            alerta.accept()
                            print("Alerta aceito.")
                            time.sleep(2)

                        except TimeoutException:
                            print("Alerta não encontrado. Prosseguindo...")
                            time.sleep(2)

                        # Obter a lista de arquivos depois do download
                        arquivos_depois = os.listdir(pasta_downloads)

                        # Encontrar o novo arquivo baixado (se houver)
                        novo_arquivo = next((f for f in arquivos_depois if f not in arquivos_antes), None)

                        # Renomear o novo arquivo baixado (se encontrado)
                        if novo_arquivo:
                            data_inicio_formatada = data_inicio.strftime("%d-%m-%Y")
                            data_fim_formatada = data_fim.strftime("%d-%m-%Y")
                            nome_arquivo_esperado = f"Relatorio Nao Voados {nome_cliente} ({data_inicio_formatada} - {data_fim_formatada})."
                            extensao_arquivo = os.path.splitext(novo_arquivo)[1]

                            arquivo_antigo = os.path.join(pasta_downloads, novo_arquivo)
                            arquivo_novo = os.path.join(pasta_downloads, f"{nome_arquivo_esperado}{extensao_arquivo}")
                            try:
                                os.rename(arquivo_antigo, arquivo_novo)
                                print(f"Arquivo '{novo_arquivo}' renomeado para '{nome_arquivo_esperado}{extensao_arquivo}' com sucesso!")
                            except Exception as e:
                                print(f"Erro ao renomear arquivo '{novo_arquivo}': {e}")
                    
                        print(f"Download de {nome_cliente} concluído!")

                    # Voltar para a página inicial para o próximo cliente
                    print(f"Voltando para a página inicial...")
                    driver.get("https://www.argoit.com.br/pontestur/Home.aspx")
                    time.sleep(3)

                    print(f"{nome_cliente} processado com sucesso!")

                except Exception as e:
                    chave_erro = str(e)
                    if erros_repetidos[chave_erro]["contagem"] == 0:
                        erros_repetidos[chave_erro]["cliente"] = nome_cliente
                        print(f"Erro ao processar cliente {nome_cliente}: {e}")
                    erros_repetidos[chave_erro]["contagem"] += 1

                    # Voltar para a página inicial para o próximo cliente
                    print(f"Voltando para a página inicial...")
                    driver.get("https://www.argoit.com.br/pontestur/Home.aspx")
                    time.sleep(3)
                    continue

            print("Loop encerrado.")

            # Imprima erros repetidos no final do log
            print("\nErros Repetidos:")
            for erro, info in erros_repetidos.items():
                if info["contagem"] > 1:
                    print(f"Erro: {erro} (repetido {info['contagem']} vezes com o cliente {info['cliente']})")

        except Exception as e:
            print(f"Erro durante a automacao: {e}")

        # Aviso de fim da automacao
        print("Automacao finalizada!")

    # Restaurar a saída padrão
    sys.stdout = original_stdout

    # Fechar o navegador
    driver.quit()

    # Exibir notificação de conclusão
    notificar_conclusao()


async def executar_processo_completo(data_inicio, data_fim, tipo_arquivo):
    """Executa o processo completo (agendamento e download)."""
    await executar_acao_por_cliente(data_inicio, data_fim, tipo_arquivo, apenas_agendar=True)
    await executar_acao_por_cliente(data_inicio, data_fim, tipo_arquivo, apenas_baixar=True)


def notificar_conclusao(titulo="Automação Finalizada", mensagem="O processo de automação foi concluído com sucesso!"):
    """Exibe uma notificação do Windows."""
    try:
        toaster = ToastNotifier()
        toaster.show_toast(titulo, mensagem, duration=10, threaded=True)
    except Exception as e:
        print(f"Erro ao exibir notificação: {e}")