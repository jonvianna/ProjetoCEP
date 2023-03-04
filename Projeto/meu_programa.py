import pandas as pd
import tkinter as tk
from tkinter import messagebox

def verifica_cep_no_range(cep, df):
    for index, row in df.iterrows():
        if row['CEP Inicial'] <= cep <= row['CEP Final']:
            return True
    return False

def verificar_cep():
    cep = int(cep_entry.get())
    resultado = verifica_cep_no_range(cep, df)
    if resultado:
        messagebox.showinfo("Resultado", "Possui coleta")
    else:
        messagebox.showinfo("Resultado", "Não possui coleta")

# carrega o arquivo Excel
df = pd.read_excel('Planilhafonte.xlsx', header=None, names=['Faixas de CEP'])

# Divide a coluna em duas novas colunas
df[['CEP Inicial', 'CEP Final']] = df['Faixas de CEP'].str.split(' ao ', expand=True)

# Converte as colunas de string para inteiro
df['CEP Inicial'] = df['CEP Inicial'].astype(int)
df['CEP Final'] = df['CEP Final'].astype(int)

# cria a janela principal
janela = tk.Tk()
janela.title("Verificador de CEP")

# cria o rótulo e a entrada para o CEP
cep_label = tk.Label(janela, text="Digite um CEP:")
cep_label.pack()
cep_entry = tk.Entry(janela)
cep_entry.pack()

# cria o botão para verificar o CEP
verificar_button = tk.Button(janela, text="Verificar", command=verificar_cep)
verificar_button.pack()

# inicia o loop da janela
janela.mainloop()