import win32com.client as win32
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para converter qualquer arquivo Excel para .xlsb
def converter_para_xlsb(arquivo_excel, arquivos_convertidos, arquivos_nao_encontrados, log_output):
    try:
        # Normalizar o caminho do arquivo
        arquivo_excel = os.path.normpath(arquivo_excel)
        
        # Verifica se o arquivo existe
        if not os.path.exists(arquivo_excel):
            log_output.insert(tk.END, f"Erro: O arquivo {arquivo_excel} não foi encontrado.\n")
            arquivos_nao_encontrados.append(arquivo_excel)
            return

        # Define o caminho do arquivo .xlsb
        arquivo_xlsb = os.path.splitext(arquivo_excel)[0] + '.xlsb'

        # Se o arquivo .xlsb já existe, apaga ele para evitar duplicatas
        if os.path.exists(arquivo_xlsb):
            os.remove(arquivo_xlsb)
            log_output.insert(tk.END, f"Arquivo {arquivo_xlsb} existente foi removido.\n")

        # Inicia o Excel em modo invisível e sem interação
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        # Abre o arquivo Excel (independente da extensão)
        workbook = excel.Workbooks.Open(arquivo_excel)

        # Salva o arquivo como .xlsb
        workbook.SaveAs(arquivo_xlsb, FileFormat=50)

        # Fecha o arquivo e o Excel
        workbook.Close(False)
        excel.Quit()

        # Remove o arquivo original (opcional)
        os.remove(arquivo_excel)
        log_output.insert(tk.END, f"Arquivo convertido com sucesso para {arquivo_xlsb} e o original foi removido.\n")

        # Adiciona à lista de arquivos convertidos
        arquivos_convertidos.append(arquivo_xlsb)

    except Exception as e:
        # Garante que o Excel será fechado em caso de erro
        log_output.insert(tk.END, f"Ocorreu um erro com o arquivo {arquivo_excel}: {e}\n")
        if 'excel' in locals():
            excel.Quit()

# Função para selecionar os arquivos
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos do Excel",
        filetypes=[("Arquivos Excel", "*.xls;*.xlsx;*.xlsm;*.xltx")]
    )
    for arquivo in arquivos:
        listbox_arquivos.insert(tk.END, arquivo)

# Função para converter os arquivos selecionados
def converter_arquivos():
    arquivos_convertidos = []
    arquivos_nao_encontrados = []

    arquivos = listbox_arquivos.get(0, tk.END)
    
    if not arquivos:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para conversão.")
        return

    for arquivo in arquivos:
        converter_para_xlsb(arquivo, arquivos_convertidos, arquivos_nao_encontrados, log_output)

    log_output.insert(tk.END, f"\nProcesso concluído!\nTotal de arquivos convertidos: {len(arquivos_convertidos)}\n")
    if arquivos_convertidos:
        log_output.insert(tk.END, "Arquivos convertidos:\n")
        for arquivo in arquivos_convertidos:
            log_output.insert(tk.END, f" - {arquivo}\n")

    log_output.insert(tk.END, f"Total de arquivos não encontrados: {len(arquivos_nao_encontrados)}\n")
    if arquivos_nao_encontrados:
        log_output.insert(tk.END, "Arquivos não encontrados:\n")
        for arquivo in arquivos_nao_encontrados:
            log_output.insert(tk.END, f" - {arquivo}\n")

# Função para limpar a lista de arquivos
def limpar_lista():
    listbox_arquivos.delete(0, tk.END)

# Configuração da janela principal
root = tk.Tk()
root.title("Conversor de Arquivos Excel para XLSB")
root.geometry("600x400")

# Frame para a seleção de arquivos
frame_arquivos = tk.Frame(root)
frame_arquivos.pack(pady=10)

listbox_arquivos = tk.Listbox(frame_arquivos, width=80, height=8)
listbox_arquivos.pack(side=tk.LEFT)

scrollbar = tk.Scrollbar(frame_arquivos)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

listbox_arquivos.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox_arquivos.yview)

# Botões de controle
frame_botoes = tk.Frame(root)
frame_botoes.pack(pady=10)

btn_selecionar = tk.Button(frame_botoes, text="Selecionar Arquivos", command=selecionar_arquivos)
btn_selecionar.pack(side=tk.LEFT, padx=5)

btn_converter = tk.Button(frame_botoes, text="Converter Arquivos", command=converter_arquivos)
btn_converter.pack(side=tk.LEFT, padx=5)

btn_limpar = tk.Button(frame_botoes, text="Limpar Lista", command=limpar_lista)
btn_limpar.pack(side=tk.LEFT, padx=5)

# Área de saída de log
log_output = tk.Text(root, width=80, height=10)
log_output.pack(pady=10)

# Loop da aplicação
root.mainloop()
