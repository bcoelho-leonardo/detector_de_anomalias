# streamlit_app.py
# -*- coding: utf-8 -*-
import streamlit as st
from io import BytesIO
from detector_de_anomalias_Claude import process_file

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
            resultado = process_file(BytesIO(uploaded.read()))

        st.success("Pronto! Baixe o arquivo destacado:")
        st.download_button(
            label="⬇️ Download Excel",
            data=resultado,
            file_name=uploaded.name.replace(".xlsx", "_highlighted.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
