# -*- coding: utf-8 -*-
"""
Created on Fri May 16 11:54:28 2025

@author: Admin
"""

import streamlit as st
from io import BytesIO
from detector_de_anomalias_streamlit import process_file

st.set_page_config(page_title="Detector de Anomalias", layout="centered")

st.title("📊 Detector de Anomalias Financeiras (último mês)")
st.markdown(
"""
Envie sua planilha **.xlsx** (aba **TD Dados**).  
O algoritmo destacará:  
* 🟡 valores atípicos (LOF)  
* 🔴 ausências incomuns no último mês
"""
)

uploaded = st.file_uploader("Escolher arquivo Excel", type=["xlsx", "xls"])

if uploaded:
    if st.button("Processar"):
        with st.spinner("Analisando..."):
            try:
                resultado = process_file(BytesIO(uploaded.read()))
            except Exception as e:
                import traceback
                st.error(f"Erro ao processar o arquivo: {e}")
                st.exception(e)  # Shows traceback in Streamlit UI
            else:
                st.success("Pronto! Baixe o arquivo destacado:")
                st.download_button(
                    label="⬇️ Download Excel",
                    data=resultado,
                    file_name=uploaded.name.replace(".xlsx", "_highlighted.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
