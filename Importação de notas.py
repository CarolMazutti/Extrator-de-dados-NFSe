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

def limpar_texto(texto):
    """Remove espaços extras e quebra de linhas indesejadas."""
    return ' '.join(texto.split())

def extrair_dados_pdf(caminho_pdf):
    """Extrai dados relevantes do PDF baseado no layout fornecido."""
    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"
    
    # Limpando texto para facilitar o parsing
    texto_completo = limpar_texto(texto_completo)

    # Parsing baseado nos padrões do PDF enviado
    numero_nota = "Não encontrado"
    prestador = "Não encontrado"
    tomador = "Não encontrado"
    cnpj = "Não encontrado"
    valor_servico = "Não encontrado"
    valor_deducao = "Não encontrado"
    valor_iss = "Não encontrado"

    # Localizando informações no texto
    if "NFS-e Nº" in texto_completo:
        numero_nota = texto_completo.split("NFS-e Nº")[1].split(" ")[0].strip()
    if "LUX CONTABILIDADE E AUDITORIA" in texto_completo:
        prestador = "LUX CONTABILIDADE E AUDITORIA"
    if "TOMADOR DO SERVIÇO" in texto_completo:
        try:
            tomador_inicio = texto_completo.index("TOMADOR DO SERVIÇO") + len("TOMADOR DO SERVIÇO")
            tomador_info = texto_completo[tomador_inicio:].split("CPF/CNPJ")[0].strip()
            cnpj = texto_completo.split("CPF/CNPJ")[1].split(" ")[0].strip()
            tomador = tomador_info
        except IndexError:
            tomador = "Não encontrado"
            cnpj = "Não encontrado"
    if "Valor Serviço" in texto_completo:
        try:
            valor_servico = texto_completo.split("Valor Serviço")[1].split(" ")[0].strip()
        except IndexError:
            valor_servico = "Não encontrado"
    if "Valor Dedução" in texto_completo:
        try:
            valor_deducao = texto_completo.split("Valor Dedução")[1].split(" ")[0].strip()
        except IndexError:
            valor_deducao = "Não encontrado"
    if "Valor ISS" in texto_completo:
        try:
            valor_iss = texto_completo.split("Valor ISS")[1].split(" ")[0].strip()
        except IndexError:
            valor_iss = "Não encontrado"

    return {
        "Numero Nota": numero_nota,
        "Prestador": prestador,
        "Tomador": tomador,
        "CNPJ": cnpj,
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
