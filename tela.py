import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


# Definir a janela principal da interface gráfica
root = tk.Tk()
root.title("Extrair Dados de Planilha")
root.geometry("400x300")

# Criar a caixa de entrada para o nome da planilha
lbl_planilha = tk.Label(root, text="Selecione a planilha:")
lbl_planilha.pack()
entrada_planilha = tk.Entry(root)
entrada_planilha.pack()

# Criar o botão para selecionar a planilha
def selecionar_planilha():
    caminho_planilha = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Excel Files", "*.xlsx")])
    entrada_planilha.delete(0, tk.END)
    entrada_planilha.insert(0, caminho_planilha)

btn_procurar = tk.Button(root, text="Procurar", command=selecionar_planilha)
btn_procurar.pack()

# Criar a caixa de entrada para as palavras-chave a serem buscadas
lbl_palavras_chave = tk.Label(root, text="Digite as palavras-chave separadas por vírgula:")
lbl_palavras_chave.pack()
entrada_palavras_chave = tk.Entry(root)
entrada_palavras_chave.pack()

# Função para buscar linhas que contêm as palavras-chave especificadas
def buscar_linhas():
    # Ler a planilha em um DataFrame
    nome_planilha = entrada_planilha.get()
    if nome_planilha == '':
        messagebox.showerror("Erro", "Nome da planilha não especificado!")
        return
    try:
        df = pd.read_excel(nome_planilha)
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Não foi possível encontrar a planilha {nome_planilha}!")
        return

    # Ler as palavras-chave a serem buscadas
    palavras_chave = entrada_palavras_chave.get().split(",")
    if not palavras_chave:
        messagebox.showerror("Erro", "Nenhuma palavra-chave especificada!")
        return

    # Função para verificar se as palavras-chave estão presentes em uma linha
    def buscar_palavras_chave(row):
        for palavra in palavras_chave:
            if palavra.lower() in str(row).lower():
                return True
        return False

    # Aplicar a função em cada linha do DataFrame
    filtro = df.apply(buscar_palavras_chave, axis=1)
    resultados = df[filtro]

    # Verificar se houve resultados encontrados
    if resultados.empty:
        messagebox.showinfo("Extrair Dados de Planilha", "Nenhum resultado encontrado.")
        return

    # Salvar os resultados em uma nova planilha
    nome_resultado = f"{nome_planilha[:-5]}_resultados.xlsx"
    arquivo_resultado = filedialog.asksaveasfilename(title="Salvar resultado da busca", filetypes=[("Excel Files", "*.xlsx")], defaultextension=".xlsx", initialfile=nome_resultado)
    if arquivo_resultado == '':
        messagebox.showerror("Erro", "Arquivo não selecionado!")
        return
    resultados.to_excel(arquivo_resultado, index=False)
    messagebox.showinfo("Extrair Dados de Planilha", f"{len(resultados)} resultados encontrados e salvos em {arquivo_resultado}.")

# Criar o botão para iniciar a busca de linhas
btn_buscar = tk.Button(root, text="Buscar Linhas", command=buscar_linhas)
btn_buscar.pack()

root.mainloop()