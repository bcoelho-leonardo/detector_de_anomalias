# -*- coding: utf-8 -*-
"""
Created on Fri May 16 12:22:13 2025

@author: Admin
"""
import streamlit as st
import traceback
from io import BytesIO
import os
import sys

# set_page_config MUST be the first Streamlit command
st.set_page_config(page_title="Detector de Anomalias", layout="centered")

# Global try-except block to catch any application-level errors
try:
    # Add directory to Python path to ensure module can be found
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from detector_de_anomalias_streamlit import process_file
    
    # Debug info at startup
    st.title("📊 Detector de Anomalias Financeiras (último mês)")
    
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
        st.write(f"Arquivo '{uploaded.name}' carregado com sucesso")
        
        # Store the uploaded file data
        uploaded_data = uploaded.read()
        file_bytes = BytesIO(uploaded_data)
        
        # Process button
        if st.button("Processar"):
            with st.spinner("Analisando..."):
                try:
                    # Process the file
                    file_bytes.seek(0)  # Reset pointer
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
                    st.error(f"Erro ao processar o arquivo: {str(e)}")
                    st.code(traceback.format_exc())
    
except Exception as main_err:
    st.error(f"Erro na aplicação: {str(main_err)}")
    st.code(traceback.format_exc())