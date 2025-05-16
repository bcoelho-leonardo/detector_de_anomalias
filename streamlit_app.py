# -*- coding: utf-8 -*-
"""
Created on Fri May 16 11:54:28 2025
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

# Debug info at startup
st.set_page_config(page_title="Detector de Anomalias", layout="centered")
st.title("📊 Detector de Anomalias Financeiras (último mês)")

# Show debug version
st.caption("Debug Version 1.0")

st.markdown(
"""
Envie sua planilha **.xlsx** (aba **TD Dados**).  
O algoritmo destacará:  
* 🟡 valores atípicos (LOF)  
* 🔴 ausências incomuns no último mês
"""
)

# File uploader with debug info
uploaded = st.file_uploader("Escolher arquivo Excel", type=["xlsx", "xls"])

# When a file is uploaded
if uploaded:
    st.write(f"File '{uploaded.name}' uploaded successfully")
    st.write(f"File type: {type(uploaded)}")
    
    # Process button
    if st.button("Processar"):
        st.write("Starting processing...")
        
        # Processing with detailed debug info
        with st.spinner("Analisando..."):
            try:
                # Step 1: Read the file
                st.write("Reading file...")
                file_bytes = BytesIO(uploaded.read())
                file_size = len(file_bytes.getvalue())
                st.write(f"File size: {file_size} bytes")
                
                # Step 2: Reset file pointer
                st.write("Resetting file pointer...")
                file_bytes.seek(0)
                
                # Step 3: Process the file
                st.write("Calling process_file...")
                resultado = process_file(file_bytes)
                st.write(f"Processing complete! Result size: {len(resultado)} bytes")
                
            except Exception as e:
                # Detailed error handling
                st.write("⚠️ Error occurred during processing")
                st.error(f"Erro ao processar o arquivo: {str(e)}")
                
                # Show traceback in code block
                error_details = traceback.format_exc()
                st.code(error_details)
                
                # Exception type info
                st.write(f"Exception type: {type(e).__name__}")
                
                # Suggestions based on error type
                if "TD Dados" in str(e):
                    st.warning("Verifique se seu arquivo Excel contém uma aba chamada 'TD Dados'")
                elif "ABEL" in str(e):
                    st.warning("Verifique se a coluna A contém a palavra 'ABEL'")
                elif "header" in str(e) or "cabeçalho" in str(e):
                    st.warning("Verifique o formato do cabeçalho da planilha")
                
            else:
                # Success handling
                st.success("Pronto! Baixe o arquivo destacado:")
                st.download_button(
                    label="⬇️ Download Excel",
                    data=resultado,
                    file_name=uploaded.name.replace(".xlsx", "_highlighted.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                
                # Success details
                st.write("✅ Arquivo processado com sucesso.")
                st.balloons()