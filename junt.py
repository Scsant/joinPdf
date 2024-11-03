import streamlit as st
from pypdf import PdfReader, PdfWriter
from io import BytesIO
import zipfile
import pyperclip  # Biblioteca para copiar para a área de transferência
import fitz  # PyMuPDF para leitura de PDF
import re
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

# Função para extrair informações do texto de cada página
def extract_info(page_text):
    pattern_nota = r"Nº\s*(\d+)"
    pattern_ordem_venda = r"Ordem de Venda:\s*(\d+)"
    pattern_fatura = r"Fatura:\s*-?(\d+)"
    pattern_remessa = r"Remessa:\s*(\d+)"
    pattern_chave_acesso = r"\b\d{2}\.\d{2}\.\d{8}\.\d{2}\.\d{8}\.\d{2}-\d+\b"
    pattern_chave_acesso_alt = r"\d{4} \d{4} \d{4} \d{4} \d{4} \d{4} \d{4} \d{4} \d{4} \d{4} \d{4}"
    pattern_transportadora = r"TRANSPORTADOR/VOLUMES TRANSPORTADOS\s+RAZÃO SOCIAL\s+([A-Z\s]+)"

    # Extração usando expressões regulares
    nota = re.search(pattern_nota, page_text)
    ordem_venda = re.search(pattern_ordem_venda, page_text)
    fatura = re.search(pattern_fatura, page_text)
    remessa = re.search(pattern_remessa, page_text)
    chave_acesso = re.search(pattern_chave_acesso, page_text) or re.search(pattern_chave_acesso_alt, page_text)
    transportadora = re.search(pattern_transportadora, page_text)

    return {
        "Data": pd.Timestamp.now().strftime("%d/%m/%Y"),
        "Nota": nota.group(1) if nota else None,
        "Ordem_de_Venda": ordem_venda.group(1) if ordem_venda else None,
        "Fatura": fatura.group(1) if fatura else None,
        "Remessa": remessa.group(1) if remessa else None,
        "CHAVE DE ACESSO": chave_acesso.group(0) if chave_acesso else None,
        "Transportadora": transportadora.group(1).strip() if transportadora else None
    }

# Função para processar o PDF e extrair as informações por página
def process_pdf(file):
    pdf_document = fitz.open(stream=file.read(), filetype="pdf")
    extracted_data = []
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        text = page.get_text()
        extracted_data.append(extract_info(text))
    pdf_document.close()
    return pd.DataFrame(extracted_data)

# Função para salvar os dados extraídos em um arquivo Excel
def export_to_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Notas Fiscais")
    output.seek(0)
    return output

# Função para copiar o DataFrame para a área de transferência no formato desejado
def copy_to_clipboard(dataframe):
    formatted_text = ""
    for _, row in dataframe.iterrows():
        formatted_row = "\t".join(map(str, row))  # Usa tabulação entre valores
        formatted_text += formatted_row + "\n"  # Nova linha para cada linha de dados
    pyperclip.copy(formatted_text.strip())  # Copia o texto formatado para a área de transferência



def merge_pdfs(pdf_files):
    writer = PdfWriter()
    for pdf_file in pdf_files:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            writer.add_page(page)
    
    output_pdf = BytesIO()
    writer.write(output_pdf)
    output_pdf.seek(0)
    return output_pdf

# Função para dividir PDF por página
def split_pdf_pages(pdf_file):
    reader = PdfReader(pdf_file)
    pages = []
    
    for i in range(len(reader.pages)):
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        
        page_file = BytesIO()
        writer.write(page_file)
        page_file.seek(0)
        
        pages.append((f"pagina_{i+1}.pdf", page_file))
    
    return pages

# Função para criar um arquivo ZIP com as páginas
def create_zip(pages):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for filename, filedata in pages:
            zip_file.writestr(filename, filedata.read())
    zip_buffer.seek(0)
    return zip_buffer

# Configuração da Interface do Streamlit
st.title("Manipulação de PDFs: Juntar ou Dividir")
st.write("Escolha se deseja juntar múltiplos PDFs ou dividir um PDF em páginas individuais.")

# Seleção da Operação
operation = st.radio("Escolha a operação:", ("Juntar PDFs", "Dividir PDF em Páginas Individuais"))

# Interface para Juntar PDFs
if operation == "Juntar PDFs":
    st.subheader("Juntar PDFs")
    uploaded_files = st.file_uploader("Escolha os arquivos PDF para juntar", accept_multiple_files=True, type="pdf")
    
    if uploaded_files:
        output_name = st.text_input("Nome do arquivo final (ex: documento_final.pdf)", "documento_final.pdf")
        if st.button("Juntar PDFs"):
            with st.spinner("Juntando arquivos..."):
                output_pdf = merge_pdfs(uploaded_files)
                st.success("Arquivos combinados com sucesso!")
                st.download_button(
                    label="Baixar PDF combinado",
                    data=output_pdf,
                    file_name=output_name,
                    mime="application/pdf"
                )

# Interface para Dividir PDF
elif operation == "Dividir PDF em Páginas Individuais":
    st.subheader("Dividir PDF em Páginas Individuais")
    pdf_file = st.file_uploader("Escolha o arquivo PDF para dividir", type="pdf")
    
    if pdf_file:
        if st.button("Dividir PDF"):
            with st.spinner("Dividindo páginas..."):
                pages = split_pdf_pages(pdf_file)
                
                # Criação de um arquivo ZIP com todas as páginas separadas
                zip_pages = create_zip(pages)
                
                st.success("PDF dividido com sucesso!")
                st.download_button(
                    label="Baixar páginas em arquivo ZIP",
                    data=zip_pages,
                    file_name="paginas_individuais.zip",
                    mime="application/zip"
                )


# Interface Streamlit
st.title("Extrair Informações de Notas Fiscais em PDF")
st.write("Carregue um arquivo PDF e extraia dados específicos para gerar um arquivo Excel.")

# Seção de Upload e Processamento
uploaded_pdf = st.file_uploader("Escolha o arquivo PDF", type="pdf")
if uploaded_pdf:
    if st.button("Extrair Informações"):
        with st.spinner("Extraindo informações..."):
            extracted_data = process_pdf(uploaded_pdf)
            st.success("Informações extraídas com sucesso!")
            st.dataframe(extracted_data)  # Exibe a tabela com dados extraídos

            # Botão para download do Excel
            excel_data = export_to_excel(extracted_data)
            st.download_button(
                label="Baixar informações em Excel",
                data=excel_data,
                file_name="notas_fiscais.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Botão para copiar dados no formato desejado para a área de transferência
            if st.button("Copiar para Área de Transferência"):
                copy_to_clipboard(extracted_data)
                st.success("Dados copiados para a área de transferência!")