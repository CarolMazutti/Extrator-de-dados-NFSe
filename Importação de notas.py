import re
import tkinter as tk
from tkinter import filedialog
import pdfplumber

def selecionar_arquivo():
    """Abre uma caixa de diálogo para o usuário selecionar um arquivo PDF."""
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal do Tkinter
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    return caminho_arquivo

def extrair_por_regex(texto, padrao, default="Não encontrado"):
    """Extrai um valor do texto com base em um padrão regex."""
    match = re.search(padrao, texto)
    return match.group(1).strip() if match else default

def extrair_dados_pdf(caminho_pdf):
    """Extrai dados relevantes do PDF baseado no layout fornecido."""
    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"
    
    # Exibe o texto completo para análise
    print("Texto extraído completo:")
    print(texto_completo[:2000])  # Limita a impressão para os primeiros 2000 caracteres

    # Limpando texto para evitar problemas com quebras
    texto_completo = ' '.join(texto_completo.split())

    # Extração com regex para Número da Nota
    numero_nota = extrair_por_regex(texto_completo, r"NFS-e Nº (\d+)")

    # Extrair nome do prestador
    prestador = "LUX CONTABILIDADE E AUDITORIA"
    
    # Ajuste para pegar o Tomador e CNPJ mais robusto
    tomador = extrair_por_regex(texto_completo, r"TOMADOR DO SERVIÇO.*?Nome/Razão Social\s*(.*?)\s*CPF/CNPJ")
    cnpj_tomador = extrair_por_regex(texto_completo, r"CPF/CNPJ\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})")

    # Ajustes nas expressões para capturar valores monetários corretamente
    valor_servico = extrair_por_regex(texto_completo, r"Valor Serviço\s*([\d\.,]+)")
    valor_deducao = extrair_por_regex(texto_completo, r"Valor Dedução\s*([\d\.,]+)")
    valor_iss = extrair_por_regex(texto_completo, r"Valor ISS\s*([\d\.,]+)")

    return {
        "Numero Nota": numero_nota,
        "Prestador": prestador,
        "Tomador": tomador,
        "CNPJ Tomador": cnpj_tomador,
        "Valor Serviço": valor_servico,
        "Valor Dedução": valor_deducao,
        "Valor ISS": valor_iss,
    }

# Função principal
def main():
    print("Por favor, selecione o arquivo PDF.")
    caminho_pdf = selecionar_arquivo()
    
    if not caminho_pdf:
        print("Nenhum arquivo selecionado. Saindo...")
        return
    
    print(f"Arquivo selecionado: {caminho_pdf}")
    dados = extrair_dados_pdf(caminho_pdf)
    print("Dados extraídos:")
    for chave, valor in dados.items():
        print(f"{chave}: {valor}")

if __name__ == "__main__":
    main()
