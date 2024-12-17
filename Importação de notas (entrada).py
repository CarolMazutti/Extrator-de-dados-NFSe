import PyPDF2
import pandas as pd
import openai
from dotenv import load_dotenv
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# Carrega a chave da API
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

# Função para extrair texto de PDFs
def extrair_texto_pdf(caminho_pdf):
    texto = ''
    with open(caminho_pdf, 'rb') as arquivo:
        leitor = PyPDF2.PdfReader(arquivo)
        for pagina in leitor.pages:
            texto += pagina.extract_text() or ''
    return texto

# Função para limpar e filtrar o texto
def limpar_texto(texto):
    # Remove espaços extras e quebras de linha repetidas
    texto_limpo = re.sub(r'\n+', '\n', texto)  # Remove linhas vazias
    texto_limpo = re.sub(r'\s{2,}', ' ', texto_limpo)  # Substitui múltiplos espaços por um único
    
    # Remove rodapés e informações irrelevantes (exemplo com palavras-chave)
    texto_limpo = re.sub(r'Página \d+ de \d+|Software:.+|Impressão:.+', '', texto_limpo)
    
    # Mantém apenas linhas com palavras-chave importantes (exemplo de palavras usadas em NFSe)
    palavras_chave = ['CNPJ', 'Valor', 'Data', 'Nota Fiscal', 'Total', 'Prestador', 'Tomador']
    linhas = [linha for linha in texto_limpo.split('\n') if any(palavra in linha for palavra in palavras_chave)]
    
    return '\n'.join(linhas)  # Retorna apenas as linhas relevantes

# Função para análise com ChatGPT
def analisar_texto_chatgpt(texto_pdf):
    resposta = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",  
        messages=[
            {"role": "system", "content": "Extraia as informações da NFSe como um dicionário JSON."},
            {"role": "user", "content": texto_pdf}
        ]
    )
    return resposta['choices'][0]['message']['content']

# Função para salvar arquivos
def salvar_arquivo(dados, formato):
    df = pd.DataFrame([dados])
    pasta_destino = filedialog.askdirectory(title="Selecione o local para salvar o arquivo")
    if not pasta_destino:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada. Processo cancelado.")
        return
    
    if formato == 'CSV':
        caminho_arquivo = os.path.join(pasta_destino, 'notas.csv')
        df.to_csv(caminho_arquivo, index=False)
    elif formato == 'ODS':
        caminho_arquivo = os.path.join(pasta_destino, 'notas.ods')
        df.to_excel(caminho_arquivo, index=False, engine='odf')
    messagebox.showinfo("Sucesso", f"Arquivo salvo em: {caminho_arquivo}")

# Interface gráfica
def iniciar_processo():
    # Pergunta inicial - NFSe ou ND
    janela = tk.Tk()
    janela.withdraw()
    especie = messagebox.askquestion("Tipo de Nota", "A nota é NFSe?")
    especie = 'NFSe' if especie == 'yes' else 'ND'
    modelo = '99' if especie == 'NFSe' else '98'
    
    # Seleção da pasta com PDFs
    pasta_pdfs = filedialog.askdirectory(title="Selecione a pasta com os arquivos PDF")
    if not pasta_pdfs:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada. Processo cancelado.")
        return
    
    # Processamento dos arquivos PDF
    dados_extracao = []
    for arquivo in os.listdir(pasta_pdfs):
        if arquivo.endswith('.pdf'):
            caminho_pdf = os.path.join(pasta_pdfs, arquivo)
            print(f"Processando: {arquivo}")
            texto_extraido = extrair_texto_pdf(caminho_pdf)
            texto_filtrado = limpar_texto(texto_extraido)  # Limpa e filtra
            try:
                dados_extraidos = analisar_texto_chatgpt(texto_filtrado)
                dados = eval(dados_extraidos)  # Converte JSON/texto para dicionário
                dados['Espécie'] = especie
                dados['Modelo'] = modelo
                dados_extracao.append(dados)
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")
    
    # Escolha do formato de saída
    formato = messagebox.askquestion("Formato do Arquivo", "Deseja salvar em CSV?")
    formato = 'CSV' if formato == 'yes' else 'ODS'
    salvar_arquivo(dados_extracao, formato)

# Executa o programa
if __name__ == "__main__":
    iniciar_processo()
