import streamlit as st
from pypdf import PdfReader, PdfWriter
from io import BytesIO
import zipfile
import pyperclip  # Biblioteca para copiar para a área de transferência
import fitz  # PyMuPDF para leitura de PDF
import re
import pandas as pd
from openpyxl import Workbook

# Função para contar páginas e juntar PDFs
def merge_pdfs(pdf_files):
    writer = PdfWriter()
    total_pages = 0  # Variável para contar o total de páginas

    for pdf_file in pdf_files:
        reader = PdfReader(pdf_file)
        total_pages += len(reader.pages)  # Incrementa o contador de páginas
        for page in reader.pages:
            writer.add_page(page)
    
    output_pdf = BytesIO()
    writer.write(output_pdf)
    output_pdf.seek(0)
    
    return output_pdf, total_pages  # Retorna o PDF combinado e o total de páginas

# Interface do Streamlit
st.title("Manipulação de PDFs: Juntar ou Dividir")
st.write("Escolha se deseja juntar múltiplos PDFs ou dividir um PDF em páginas individuais.")

# Seleção da Operação
operation = st.radio("Escolha a operação:", ("Juntar PDFs", "Dividir PDF em Páginas Individuais"))

# Interface para Juntar PDFs
if operation == "Juntar PDFs":
    st.subheader("Juntar PDFs")
    uploaded_files = st.file_uploader("Escolha os arquivos PDF para juntar", accept_multiple_files=True, type="pdf")
    
    if uploaded_files:
        # Calcular o total de páginas antes de juntar
        total_pages = sum(len(PdfReader(file).pages) for file in uploaded_files)
        st.write(f"Total de páginas nos arquivos selecionados: {total_pages}")

        output_name = st.text_input("Nome do arquivo final (ex: documento_final.pdf)", "documento_final.pdf")
        if st.button("Juntar PDFs"):
            with st.spinner("Juntando arquivos..."):
                output_pdf, total_pages_combined = merge_pdfs(uploaded_files)
                st.success(f"Arquivos combinados com sucesso! Total de páginas no PDF combinado: {total_pages_combined}")
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
