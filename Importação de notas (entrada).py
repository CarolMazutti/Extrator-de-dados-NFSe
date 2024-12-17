import PyPDF2
import pandas as pd
from dotenv import load_dotenv
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Carrega variáveis de ambiente
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

# Verifica se a chave da API foi carregada
if api_key:
    import openai
    openai.api_key = api_key
else:
    messagebox.showerror("Erro", "A chave da API não foi encontrada. Verifique o arquivo .env.")

# Função para extrair texto do PDF
def extrair_texto_pdf(caminho_pdf):
    try:
        with open(caminho_pdf, 'rb') as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)
            texto = ''
            for pagina in leitor.pages:
                texto += pagina.extract_text() or ''
        return texto
    except Exception as e:
        print(f"Erro ao ler o arquivo PDF '{caminho_pdf}': {e}")
        return None

# Função para analisar texto usando ChatGPT
def analisar_texto_chatgpt(texto_pdf):
    if api_key:  # Se a chave da API estiver disponível
        try:
            resposta = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Extraia as informações essenciais da NFSe ou ND (CNPJ, Data, Valor, Fornecedor)."},
                    {"role": "user", "content": texto_pdf}
                ]
            )
            return resposta['choices'][0]['message']['content']
        except Exception as e:
            print(f"Erro ao usar a API OpenAI: {e}")
            return None
    return None

# Função para salvar dados em CSV ou ODS
def salvar_arquivo(dados, tipo_arquivo, pasta_destino):
    try:
        df = pd.DataFrame(dados)
        caminho_arquivo = os.path.join(pasta_destino, f"notas.{tipo_arquivo}")
        if tipo_arquivo == 'csv':
            df.to_csv(caminho_arquivo, index=False)
        elif tipo_arquivo == 'ods':
            df.to_excel(caminho_arquivo, index=False, engine='odf')
        print(f"Arquivo salvo com sucesso em: {caminho_arquivo}")
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em: {caminho_arquivo}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")

# Janela para seleção entre NFSe ou ND
def escolher_especie():
    janela = tk.Toplevel()
    janela.title("Escolha o tipo de nota")
    janela.geometry("300x150")
    janela.resizable(False, False)

    especie_selecionada = tk.StringVar()

    def selecionar(opcao):
        especie_selecionada.set(opcao)
        janela.destroy()

    tk.Label(janela, text="Selecione o tipo de nota:", font=("Arial", 12)).pack(pady=10)
    tk.Button(janela, text="NFSe", width=10, command=lambda: selecionar("NFSE")).pack(side=tk.LEFT, padx=20)
    tk.Button(janela, text="ND", width=10, command=lambda: selecionar("ND")).pack(side=tk.RIGHT, padx=20)
    janela.wait_window()
    return especie_selecionada.get()

# Janela para escolher formato de saída
def escolher_formato_saida():
    janela = tk.Toplevel()
    janela.title("Escolha o formato de saída")
    janela.geometry("300x150")
    janela.resizable(False, False)

    formato_selecionado = tk.StringVar()

    def selecionar(formato):
        formato_selecionado.set(formato)
        janela.destroy()

    tk.Label(janela, text="Selecione o formato do arquivo de saída:", font=("Arial", 12)).pack(pady=10)
    tk.Button(janela, text="CSV", width=10, command=lambda: selecionar("csv")).pack(side=tk.LEFT, padx=20)
    tk.Button(janela, text="ODS", width=10, command=lambda: selecionar("ods")).pack(side=tk.RIGHT, padx=20)
    janela.wait_window()
    return formato_selecionado.get()

# Função principal
def main():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    # Escolha do tipo de nota
    especie = escolher_especie()
    if not especie:
        print("Nenhuma opção selecionada. Encerrando...")
        return
    modelo = '99' if especie == "NFSE" else '98'

    # Seleciona a pasta com os arquivos PDF
    pasta_pdf = filedialog.askdirectory(title="Selecione a pasta onde estão os arquivos PDF")
    if not pasta_pdf:
        messagebox.showinfo("Encerrando", "Nenhuma pasta selecionada. Encerrando...")
        return

    # Escolhe o formato do arquivo de saída
    tipo_arquivo = escolher_formato_saida()
    if not tipo_arquivo:
        messagebox.showinfo("Encerrando", "Nenhum formato selecionado. Encerrando...")
        return

    # Seleciona a pasta para salvar o arquivo
    pasta_destino = filedialog.askdirectory(title="Selecione a pasta para salvar o arquivo")
    if not pasta_destino:
        messagebox.showinfo("Encerrando", "Nenhuma pasta de destino selecionada. Encerrando...")
        return

    # Processa os arquivos
    print(f"Processando arquivos na pasta: {pasta_pdf}")
    dados_coletados = []

    for arquivo in os.listdir(pasta_pdf):
        if arquivo.lower().endswith('.pdf'):
            caminho_pdf = os.path.join(pasta_pdf, arquivo)
            print(f"Extraindo dados do arquivo: {arquivo}")
            texto_extraido = extrair_texto_pdf(caminho_pdf)

            if texto_extraido:
                dados_extraidos = analisar_texto_chatgpt(texto_extraido)
                if dados_extraidos:
                    try:
                        dados = eval(dados_extraidos)  # Converte JSON/texto para dicionário
                        dados['Espécie'] = especie
                        dados['Modelo'] = modelo
                        dados_coletados.append(dados)
                    except Exception as e:
                        print(f"Erro ao processar '{arquivo}': {e}")

    # Salva os dados coletados
    if dados_coletados:
        salvar_arquivo(dados_coletados, tipo_arquivo, pasta_destino)
    else:
        messagebox.showwarning("Aviso", "Nenhum dado foi coletado.")

# Executa o programa
if __name__ == "__main__":
    main()
