import os
import PyPDF2
import pandas as pd
import re
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
    #######################################
    # linhas = texto.splitlines()  # Divide o texto em linhas

    # print("\n=== TEXTO EXTRAÍDO COM NUMERAÇÃO ===\n")
    # for i, linha in enumerate(linhas, start=1):  # Enumera as linhas a partir de 1
    #     print(f"Linha {i}: {linha}")
    #######################################
    return texto

# Função para extrair os dados do texto do PDF
def extrair_dados_nf(texto_pdf, especie, nome_arquivo):
    cnpj_prestador_match = re.search(r"CNPJ:\s*([\d./-]+)", texto_pdf)
    cnpj_prestador = re.sub(r'\D', '', cnpj_prestador_match.group(1)) if cnpj_prestador_match else "Não encontrado"

    numero_nota_match = re.search(r"Número da NFS-e\s+(\d+)", texto_pdf)
    
    datas_match = re.findall(r"\d{2}/\d{2}/\d{4}", texto_pdf)
    data_fato_gerador = datas_match[0] if len(datas_match) > 0 else "Não encontrado"
    data_emissao = datas_match[1] if len(datas_match) > 1 else "Não encontrado"
    
    if len(datas_match) < 1:
        data_fato_gerador = "Não encontrado"
    if len(datas_match) < 2:
        data_emissao = "Não encontrado"
    
    valor_total_match = re.search(r"Valor Serviço\s+([\d.,]+)", texto_pdf)
    base_calculo_match = re.search(r"Base de Cálculo\s+([\d.,]+)", texto_pdf)
    issqn_match = re.search(r"ISSQN\s+([\d.,]+)", texto_pdf)
    issrf_match = re.search(r"ISSRF\s+([\d.,]+)", texto_pdf)
    ir_match = re.search(r"IR\s+([\d.,]+)", texto_pdf)
    inss_match = re.search(r"INSS\s+([\d.,]+)", texto_pdf)
    csll_match = re.search(r"CSLL\s+([\d.,]+)", texto_pdf)
    cofins_match = re.search(r"COFINS\s+([\d.,]+)", texto_pdf)
    pis_match = re.search(r"PIS\s+([\d.,]+)", texto_pdf)
    outras_retencoes_match = re.search(r"Outras Retenções\s+([\d.,]+)", texto_pdf)
    
    descricao_servico = "Não encontrado"
    linhas = texto_pdf.splitlines()  # Divide o texto em linhas
    for i, linha in enumerate(linhas):
        if "Descrição dos subitens da Lista de Serviço em acordo com a Lei Complementar 116/03" in linha:
            if i + 1 < len(linhas):  # Verifica se há uma linha seguinte
                descricao_servico = linhas[i + 1].strip()  # Captura a linha seguinte
            break

    if descricao_servico == "Legenda do Local de Prestação do Serviço":
        descricao_servico = "Não encontrado"

    # Organiza os dados extraídos
    dados = {
        "CNPJ": cnpj_prestador,
        "Número Nota": numero_nota_match.group(1) if numero_nota_match else "Não encontrado",
        "Espécie": especie,
        "Serie": "1" if especie == "NFSE" else "U",
        "Data Fato Gerador": data_fato_gerador,
        "Data Emissão": data_emissao,
        "Valor Total": valor_total_match.group(1) if valor_total_match else "Não encontrado",
        "Base de Cálculo": base_calculo_match.group(1) if base_calculo_match else "Não encontrado",
        "ISSQN": issqn_match.group(1) if issqn_match else "Não encontrado",
        "ISSRF": issrf_match.group(1) if issrf_match else "Não encontrado",
        "IR": ir_match.group(1) if ir_match else "Não encontrado",
        "INSS": inss_match.group(1) if inss_match else "Não encontrado",
        "CSLL": csll_match.group(1) if csll_match else "Não encontrado",
        "COFINS": cofins_match.group(1) if cofins_match else "Não encontrado",
        "PIS": pis_match.group(1) if pis_match else "Não encontrado",
        "Outras Retenções": outras_retencoes_match.group(1) if outras_retencoes_match else "Não encontrado",
        "Modelo": "99" if especie == "NFSE" else "98",
        "Código e Descrição": descricao_servico,
    }

    # Adiciona o nome do arquivo na coluna "Arquivo com Erro" se houver informações ausentes
    dados["Arquivo com Erro"] = nome_arquivo if any(value == "Não encontrado" for value in dados.values()) else ""

    return dados

# Função para salvar arquivos em CSV ou ODS
def salvar_arquivo(dados, formato, root):
    # Reorganiza as colunas conforme ordem exigida
    colunas_ordem = [
	    "CNPJ",
        "Número Nota",
        "Espécie",
        "Serie",
        "Data Fato Gerador",
        "Data Emissão",
        "Valor Total",
        "Base de Cálculo",
        "ISSQN",
        "ISSRF",
        "IR",
        "INSS",
        "CSLL",
        "COFINS",
        "PIS",
        "Outras Retenções",
        "Modelo",
        "Código e Descrição",
        "Arquivo com Erro",
        "CFOP",
    ]
    df = pd.DataFrame(dados, columns=colunas_ordem)

    # Seleção do local de salvamento
    pasta_destino = filedialog.askdirectory(title="Selecione o local para salvar o arquivo")
    if not pasta_destino:
        messagebox.showwarning("Aviso", "Nenhuma pasta selecionada. Processo cancelado.")
        return

    # Salva o arquivo no formato desejado
    caminho_arquivo = os.path.join(pasta_destino, f"notas.{formato.lower()}")
    if formato == 'CSV':
        df.to_csv(caminho_arquivo, index=False, sep=';')
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
        messagebox.showwarning("Aviso," "Nenhuma pasta selecionada. Processo cancelado.")
        root.destroy()
        return

    # Processa cada PDF da pasta
    dados_totais = []
    for arquivo in os.listdir(pasta_origem):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(pasta_origem, arquivo)
            texto_extraido = extrair_texto_pdf(caminho_pdf)
            dados_extraidos = extrair_dados_nf(texto_extraido, especie, arquivo)
            dados_totais.append(dados_extraidos)

    # Tela para escolher o formato de salvamento
    janela_formato = tk.Toplevel()
    janela_formato.title("Selecione o Formato")
    janela_formato.geometry("300x100")

    lbl_instrucao = ttk.Label(janela_formato, text="Escolha o formato de salvamento:")
    lbl_instrucao.pack(pady=10)

    # Botões lado a lado
    btn_csv = ttk.Button(
        janela_formato,
        text="CSV",
        command=lambda: salvar_arquivo(dados_totais, "CSV", root),
    )
    btn_csv.pack(side=tk.LEFT, padx=20, pady=20)

    btn_ods = ttk.Button(
        janela_formato,
        text="ODS",
        command=lambda: salvar_arquivo(dados_totais, "ODS", root),
    )
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