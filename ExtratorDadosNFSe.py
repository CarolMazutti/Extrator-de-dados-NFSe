##Quando a API for paga, alterar apenas a função analisar_texto_simulado

# Função para análise com ChatGPT
# def analisar_texto_chatgpt(texto_pdf):
#     resposta = openai.ChatCompletion.create(
#         model="gpt-3.5-turbo",  
#         messages=[
#             {"role": "system", "content": "Extraia as informações da NFSe como um dicionário JSON."},
#             {"role": "user", "content": texto_pdf}
#         ]
#     )
#     return resposta['choices'][0]['message']['content']

import os
import PyPDF2
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Função para extrair texto do PDF
def extrair_texto_pdf(caminho_pdf):
    texto = ''
    with open(caminho_pdf, 'rb') as arquivo:
        leitor = PyPDF2.PdfReader(arquivo)
        for pagina in leitor.pages:
            texto += pagina.extract_text()
    return texto

# Função para simular extração de dados usando API
def analisar_texto_simulado(texto_pdf):
    # Simulação do processamento de texto para extração
    dados_extraidos = {
        "CNPJ": "12.345.678/0001-99",
        "Número Nota": "12345",
        "Data Emissão": "2024-06-18",
        "Valor Total": "1500.00",
    }
    return dados_extraidos

# Função para salvar arquivos em CSV ou ODS
def salvar_arquivo(dados, formato, especie, root):
    # Adiciona as colunas ESPECIE, SERIE e MODELO
    dados["Espécie"] = especie
    dados["Serie"] = "1" if especie == "NFSE" else "U"
    dados["Modelo"] = "99" if especie == "NFSE" else "98"

    # Reorganiza as colunas conforme ordem exigida
    colunas_ordem = ["CNPJ", "Número Nota", "Espécie", "Serie", "Data Emissão", "Valor Total", "Modelo"]
    df = pd.DataFrame([dados])
    df = df[colunas_ordem]

    # Seleção do local de salvamento
    pasta_destino = filedialog.askdirectory(title="Selecione o local para salvar o arquivo")
    if not pasta_destino:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada. Processo cancelado.")
        return

    # Salva o arquivo no formato desejado
    caminho_arquivo = os.path.join(pasta_destino, f"notas.{formato.lower()}")
    if formato == 'CSV':
        df.to_csv(caminho_arquivo, index=False)
    elif formato == 'ODS':
        df.to_excel(caminho_arquivo, index=False, engine='odf')

    messagebox.showinfo("Sucesso", f"Arquivo salvo em: {caminho_arquivo}")
    root.destroy()  # Fecha o programa

# Função principal
def processar_notas(especie, root):
    root.withdraw()  # Esconde a janela principal temporariamente

    # Selecionar a pasta com os PDFs
    pasta_origem = filedialog.askdirectory(title="Selecione a pasta com os arquivos PDF")
    if not pasta_origem:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada. Processo cancelado.")
        root.destroy()
        return

    # Processa cada PDF da pasta
    dados_totais = []
    for arquivo in os.listdir(pasta_origem):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(pasta_origem, arquivo)
            texto_extraido = extrair_texto_pdf(caminho_pdf)
            dados_extraidos = analisar_texto_simulado(texto_extraido)
            dados_totais.append(dados_extraidos)

    # Tela para escolher o formato de salvamento
    janela_formato = tk.Toplevel()
    janela_formato.title("Selecione o Formato")
    janela_formato.geometry("300x100")

    lbl_instrucao = ttk.Label(janela_formato, text="Escolha o formato de salvamento:")
    lbl_instrucao.pack(pady=10)

    # Botões lado a lado
    btn_csv = ttk.Button(janela_formato, text="CSV", 
                         command=lambda: salvar_arquivo(dados_totais[0], "CSV", especie, root))
    btn_csv.pack(side=tk.LEFT, padx=20, pady=20)

    btn_ods = ttk.Button(janela_formato, text="ODS", 
                         command=lambda: salvar_arquivo(dados_totais[0], "ODS", especie, root))
    btn_ods.pack(side=tk.RIGHT, padx=20, pady=20)

# Função para criar a interface inicial
def interface_usuario():
    root = tk.Tk()
    root.title("Importação de Notas")
    root.geometry("300x150")

    # Label de instrução
    lbl_instrucao = ttk.Label(root, text="Selecione o tipo de nota:")
    lbl_instrucao.pack(pady=10)

    # Botões lado a lado
    btn_nfse = ttk.Button(root, text="NFSe", command=lambda: processar_notas("NFSE", root))
    btn_nfse.pack(side=tk.LEFT, padx=20, pady=20)

    btn_nd = ttk.Button(root, text="ND", command=lambda: processar_notas("ND", root))
    btn_nd.pack(side=tk.RIGHT, padx=20, pady=20)

    root.mainloop()

# Executar a interface inicial
if __name__ == "__main__":
    interface_usuario()
