# -*- coding: utf-8 -*-
"""
Created on Fri May 16 12:22:13 2025

@author: Admin
"""
import streamlit as st
from io import BytesIO
import os
import sys
import traceback

# Add directory to Python path to ensure module can be found
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from detector_de_anomalias_streamlit import process_file

# Configure page
st.set_page_config(page_title="Detector de Anomalias", layout="centered")

# Debug info at startup
st.title("📊 Detector de Anomalias Financeiras (último mês)")

# Show debug version & environment info
st.caption(f"Debug Version 1.1 | Python {sys.version} | Path: {os.path.abspath(__file__)}")
st.write(f"Diretório atual: {os.getcwd()}")
st.write(f"Arquivos no diretório: {os.listdir()}")

st.markdown(
"""
Envie sua planilha **.xlsx** (aba **TD Dados**).  
O algoritmo destacará:  
* 🟡 valores atípicos (LOF)  
* 🔴 ausências incomuns no último mês
"""
)

# File uploader
uploaded = st.file_uploader("Escolher arquivo Excel", type=["xlsx", "xls"])

# When a file is uploaded
if uploaded:
    st.write(f"File '{uploaded.name}' uploaded successfully")
    st.write(f"File type: {type(uploaded)}")
    
    # Create an expander for showing file info
    with st.expander("Informações do Arquivo"):
        file_bytes = BytesIO(uploaded.read())
        st.write(f"Tamanho do arquivo: {len(file_bytes.getvalue())} bytes")
        file_bytes.seek(0)  # Reset for future use
        
        # Try to check if it's a valid Excel file
        try:
            import openpyxl
            try:
                wb = openpyxl.load_workbook(file_bytes, read_only=True)
                st.write(f"Sheets disponíveis (openpyxl): {wb.sheetnames}")
                if "TD Dados" in wb.sheetnames:
                    st.success("✅ Encontrada aba 'TD Dados'")
                else:
                    st.error("❌ Aba 'TD Dados' não encontrada!")
            except Exception as e1:
                st.warning(f"Não foi possível ler com openpyxl: {str(e1)}")
                
                # Try alternative engine
                try:
                    import pandas as pd
                    file_bytes.seek(0)
                    xls = pd.ExcelFile(file_bytes, engine='xlrd')
                    st.write(f"Sheets disponíveis (xlrd): {xls.sheet_names}")
                    if "TD Dados" in xls.sheet_names:
                        st.success("✅ Encontrada aba 'TD Dados'")
                    else:
                        st.error("❌ Aba 'TD Dados' não encontrada!")
                except Exception as e2:
                    st.error(f"Não foi possível verificar o arquivo com nenhum engine Excel. Erro: {str(e2)}")
        except Exception as e:
            st.error(f"Erro ao verificar o arquivo Excel: {str(e)}")
        file_bytes.seek(0)  # Reset again
    
    # Process button
    if st.button("Processar"):
        st.write("Starting processing...")
        
        # Processing
        with st.spinner("Analisando..."):
            try:
                # Read file again to get fresh BytesIO
                file_bytes = BytesIO(uploaded.read())
                st.write(f"File size: {len(file_bytes.getvalue())} bytes")
                
                # Reset pointer and process
                file_bytes.seek(0)
                resultado = process_file(file_bytes)
                
                # Success path
                st.success("Pronto! Baixe o arquivo destacado:")
                st.download_button(
                    label="⬇️ Download Excel",
                    data=resultado,
                    file_name=uploaded.name.replace(".xlsx", "_highlighted.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.balloons()
                
            except Exception as e:
                # Error handling
                st.error(f"Erro ao processar o arquivo: {str(e)}")
                
                # Show detailed error info
                error_details = traceback.format_exc()
                with st.expander("Detalhes do Erro"):
                    st.code(error_details)
                
                # Check for common errors
                error_msg = str(e).lower()
                if "td dados" in error_msg:
                    st.warning("⚠️ Verifique se seu arquivo Excel contém uma aba chamada 'TD Dados'")
                elif "abel" in error_msg:
                    st.warning("⚠️ Verifique se a coluna A contém a palavra 'ABEL'")
                else:
                    st.warning("⚠️ Verifique o formato do arquivo de acordo com as instruções")