import tkinter as tk
from tkinter import ttk
import asyncio
from tkinter import Toplevel, Label, BooleanVar, Button, StringVar, Checkbutton
from tkcalendar import DateEntry
from datetime import date
import naoVoados  # Importa o módulo naoVoados.py
import customtkinter as ctk  # Importa o CustomTkinter
import configparser

async def executar_automacao(acao):
    """Coleta as informações da interface e executa a automação."""
    try:
        data_inicio = calendario_inicio.get_date()
        data_fim = calendario_fim.get_date()
        tipo_arquivo = variavel_tipo_arquivo.get()

        if tipo_arquivo not in ["PDF", "EXCEL"]:
            raise ValueError("Tipo de arquivo inválido. Escolha 'PDF' ou 'EXCEL'.")

        # Chama a função para executar a automação com os parâmetros
        if acao == "Agendar":
            await naoVoados.executar_acao_por_cliente(data_inicio, data_fim, tipo_arquivo, apenas_agendar=True)
        elif acao == "Baixar":
            await naoVoados.executar_acao_por_cliente(data_inicio, data_fim, tipo_arquivo, apenas_baixar=True)
        elif acao == "Processo Completo":
            await naoVoados.executar_processo_completo(data_inicio, data_fim, tipo_arquivo)

        print(f"Automação '{acao}' finalizada!")
    except Exception as e:
        print(f"Erro ao executar a automação: {e}")

def salvar_configuracoes():
    """Salva as configurações no arquivo configuracoes.ini."""
    config = configparser.ConfigParser()
    config.read('configuracoes.ini')

    # Verifica se a seção 'INICIO' existe, se não, cria
    if 'INICIO' not in config:
        config['INICIO'] = {}

    # Atualizar a seção 'INICIO' com base no estado da checkbox 'todos_clientes_var'
    if todos_clientes_var.get():
        config['INICIO']['nao_voados'] = ""  # Define como vazio para processar todos os clientes
    else:
        config['INICIO']['nao_voados'] = entry_cliente_nao_voados.get()

    with open('configuracoes.ini', 'w') as configfile:
        config.write(configfile)
    janela_configuracoes.destroy()

def atualizar_estado_entry(entry, var_todos):
    """Habilita/desabilita a entry com base na checkbox."""
    if var_todos.get():
        entry.configure(state="disabled")  # Correção: Usar .configure()
        entry.delete(0, tk.END)
    else:
        entry.configure(state="normal")  # Correção: Usar .configure()

def abrir_configuracoes():
    """Abre a janela de configurações."""
    global janela_configuracoes, entry_cliente_nao_voados, todos_clientes_var
    janela_configuracoes = ctk.CTkToplevel(janela)
    janela_configuracoes.title("Configurações")
    janela_configuracoes.geometry("350x150")  # Ajuste o tamanho conforme necessário

    # Para garantir que a janela de configurações fique acima da janela principal
    janela_configuracoes.transient(janela)
    janela_configuracoes.grab_set()

    # Carregar configurações existentes
    config = configparser.ConfigParser()
    config.read('configuracoes.ini')
    cliente_inicial = config.get('INICIO', 'nao_voados', fallback="")

    # Cliente Inicial para Não Voados
    label_cliente_nao_voados = ctk.CTkLabel(janela_configuracoes, text="Cliente Inicial (Não Voados):")
    label_cliente_nao_voados.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

    entry_cliente_nao_voados = ctk.CTkEntry(janela_configuracoes)
    entry_cliente_nao_voados.insert(0, cliente_inicial)
    entry_cliente_nao_voados.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    # Checkbox "Todos os Clientes"
    todos_clientes_var = ctk.BooleanVar(janela_configuracoes)
    checkbox_todos_clientes = ctk.CTkCheckBox(janela_configuracoes, text="(Todos os Clientes)",
                                            variable=todos_clientes_var,
                                            command=lambda: atualizar_estado_entry(entry_cliente_nao_voados, todos_clientes_var))
    checkbox_todos_clientes.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="w")

    # Botão para salvar configurações
    botao_salvar = ctk.CTkButton(janela_configuracoes, text="Salvar", command=salvar_configuracoes)
    botao_salvar.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Janela principal
janela = ctk.CTk()
janela.title("Relatório de Não Voados")
janela.geometry("350x250")  # Aumentar a largura e altura

# Define o tema da aplicação
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Data de Início
label_data_inicio = ctk.CTkLabel(janela, text="Data de Início:")
label_data_inicio.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
calendario_inicio = DateEntry(janela, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
calendario_inicio.grid(row=0, column=1, padx=10, pady=10)

# Data de Fim
label_data_fim = ctk.CTkLabel(janela, text="Data de Fim:")
label_data_fim.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
calendario_fim = DateEntry(janela, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
data_atual = date.today()
calendario_fim.set_date(data_atual)
calendario_fim.grid(row=1, column=1, padx=10, pady=10)

# Tipo de Arquivo
label_tipo_arquivo = ctk.CTkLabel(janela, text="Tipo de Arquivo:")
label_tipo_arquivo.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
opcoes_tipo_arquivo = ["PDF", "EXCEL"]
variavel_tipo_arquivo = StringVar(janela)
combobox_tipo_arquivo = ctk.CTkComboBox(janela, variable=variavel_tipo_arquivo, values=opcoes_tipo_arquivo)
combobox_tipo_arquivo.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

# Botões para cada ação ("Agendar" e "Baixar" na mesma linha)
botao_agendar = ctk.CTkButton(janela, text="Agendar", command=lambda: asyncio.run(executar_automacao("Agendar")))
botao_agendar.grid(row=3, column=0, padx=10, pady=5)

botao_baixar = ctk.CTkButton(janela, text="Baixar", command=lambda: asyncio.run(executar_automacao("Baixar")))
botao_baixar.grid(row=3, column=1, padx=10, pady=5)

# Botões "Processo Completo" e "Configurações" na mesma linha
botao_completo = ctk.CTkButton(janela, text="Processo Completo", command=lambda: asyncio.run(executar_automacao("Processo Completo")))
botao_completo.grid(row=4, column=0, padx=10, pady=5)

botao_configuracoes = ctk.CTkButton(janela, text="Configurações", command=abrir_configuracoes)
botao_configuracoes.grid(row=4, column=1, padx=10, pady=5)

janela.mainloop()