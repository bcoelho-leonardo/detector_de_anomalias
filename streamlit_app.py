# -*- coding: utf-8 -*-
import streamlit as st
from io import BytesIO
import os
import sys
import traceback
import tempfile

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
        
        # Save file to disk first, then work with it
        if st.button("Processar"):
            with st.spinner("Analisando..."):
                try:
                    # Create a temporary file on disk
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                        tmp_file.write(uploaded.getvalue())
                        temp_path = tmp_file.name
                    
                    st.write(f"Arquivo salvo temporariamente em: {temp_path}")
                    st.write(f"Tamanho: {os.path.getsize(temp_path)} bytes")
                    
                    # Verify the file is accessible
                    try:
                        import pandas as pd
                        # Just list the sheet names to verify the file is valid
                        xls = pd.ExcelFile(temp_path)
                        st.write(f"Sheets disponíveis: {xls.sheet_names}")
                        
                        # If we can read the sheets, try to process the file directly from disk
                        with open(temp_path, "rb") as f:
                            file_data = f.read()
                            file_bytes = BytesIO(file_data)
                            
                            # Process the file with fresh BytesIO
                            resultado = process_file(file_bytes)
                        
                        # Display download button
                        st.success("Pronto! Baixe o arquivo destacado:")
                        st.download_button(
                            label="⬇️ Download Excel",
                            data=resultado,
                            file_name=uploaded.name.replace(".xlsx", "_highlighted.xlsx"),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                        st.balloons()
                        
                    except Exception as verify_err:
                        st.error(f"Erro ao verificar o arquivo: {str(verify_err)}")
                        st.code(traceback.format_exc())
                    
                    # Clean up temp file
                    os.unlink(temp_path)
                    
                except Exception as e:
                    st.error(f"Erro ao processar o arquivo: {str(e)}")
                    st.code(traceback.format_exc())
                    
                    # Try one more approach if all else fails
                    st.write("Tentando abordagem alternativa...")
                    try:
                        # Direct approach without temp file
                        import pandas as pd
                        
                        # Check if the file can be read with pandas first
                        uploaded.seek(0)
                        df = pd.read_excel(uploaded, sheet_name=None)
                        st.write(f"Sheets encontradas: {list(df.keys())}")
                        
                        if "TD Dados" in df:
                            st.write("✅ Sheet 'TD Dados' encontrada e pode ser lida")
                            
                            # Try processing again
                            uploaded.seek(0)
                            resultado = process_file(BytesIO(uploaded.read()))
                            
                            # Display download button
                            st.success("Processamento alternativo bem-sucedido!")
                            st.download_button(
                                label="⬇️ Download Excel",
                                data=resultado,
                                file_name=uploaded.name.replace(".xlsx", "_highlighted.xlsx"),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                        else:
                            st.error("❌ Sheet 'TD Dados' não encontrada")
                            
                    except Exception as alt_err:
                        st.error(f"Abordagem alternativa também falhou: {str(alt_err)}")
                        st.code(traceback.format_exc())
except Exception as e:
    st.error(f"Erro na aplicação: {str(e)}")
    st.code(traceback.format_exc())