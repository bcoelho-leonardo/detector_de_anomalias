# -*- coding: utf-8 -*-
import streamlit as st
from io import BytesIO
import os
import sys
import traceback

# set_page_config MUST be the first Streamlit command
st.set_page_config(page_title="Detector de Anomalias", layout="centered")

# Try to import the module
try:
    # Add directory to Python path to ensure module can be found
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from detector_de_anomalias_streamlit import process_file
    
    # App title and description
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
        
        # Basic file validation
        with st.expander("Informações do Arquivo"):
            try:
                # Create BytesIO and check file size
                file_bytes = BytesIO(uploaded.read())
                file_size = len(file_bytes.getvalue())
                st.write(f"Tamanho do arquivo: {file_size} bytes")
                
                # Check file header
                file_bytes.seek(0)
                file_header = file_bytes.read(8).hex() if file_size >= 8 else ""
                file_bytes.seek(0)
                
                # Identify file type
                if file_header.startswith("504b0304"):
                    st.write("✅ Arquivo Excel (.xlsx) válido")
                elif file_header.startswith("d0cf11e0"):
                    st.write("✅ Arquivo Excel (.xls) válido")
                else:
                    st.warning("⚠️ Formato de arquivo não reconhecido")
                
                # Try opening with openpyxl
                try:
                    import tempfile
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        tmp.write(file_bytes.getvalue())
                        temp_path = tmp.name
                    
                    import openpyxl
                    wb = openpyxl.load_workbook(temp_path, read_only=True)
                    st.write(f"Sheets disponíveis: {wb.sheetnames}")
                    
                    if "TD Dados" in wb.sheetnames:
                        st.success("✅ Encontrada aba 'TD Dados'")
                    else:
                        st.error("❌ Aba 'TD Dados' não encontrada!")
                    
                    os.unlink(temp_path)
                except Exception as e1:
                    st.warning(f"Não foi possível verificar estrutura: {str(e1)}")
                
                # Reset file pointer for later use
                file_bytes.seek(0)
            except Exception as e:
                st.error(f"Erro ao verificar arquivo: {str(e)}")
        
        # Process button
        if st.button("Processar"):
            with st.spinner("Analisando..."):
                try:
                    # Fresh copy of the file
                    file_bytes = BytesIO(uploaded.read())
                    
                    # Process the file
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
except Exception as e:
    st.error(f"Erro na aplicação: {str(e)}")
    st.code(traceback.format_exc())