import tkinter as tk
from tkinter import ttk
import asyncio
from tkinter import Toplevel, Label, BooleanVar, Button, StringVar, Checkbutton
from tkcalendar import DateEntry
from datetime import date
import naoVoados  # Importa o módulo naoVoados.py
import customtkinter as ctk  # Importa o CustomTkinter

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

# Janela principal
janela = ctk.CTk()
janela.title("Relatório de Não Voados")
janela.geometry("290x280")  # Aumentar altura para acomodar botões

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

# Botões para cada ação
botao_agendar = ctk.CTkButton(janela, text="Agendar", command=lambda: asyncio.run(executar_automacao("Agendar")))
botao_agendar.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

botao_baixar = ctk.CTkButton(janela, text="Baixar", command=lambda: asyncio.run(executar_automacao("Baixar")))
botao_baixar.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

botao_completo = ctk.CTkButton(janela, text="Processo Completo", command=lambda: asyncio.run(executar_automacao("Processo Completo")))
botao_completo.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

janela.mainloop()