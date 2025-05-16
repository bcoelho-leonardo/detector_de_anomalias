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
    
    # Create a copy of the uploaded file data immediately
    # Store the uploaded data in session state to preserve it between interactions
    if "file_data" not in st.session_state:
        uploaded_data = uploaded.read()
        st.session_state.file_data = uploaded_data
        st.session_state.file_size = len(uploaded_data)
    
    # Create an expander for showing file info
    with st.expander("Informações do Arquivo"):
        # Create a fresh BytesIO from the saved data
        file_bytes = BytesIO(st.session_state.file_data)
        file_size = st.session_state.file_size
        st.write(f"Tamanho do arquivo: {file_size} bytes")
        
        # Basic file signature check
        file_bytes.seek(0)
        file_header = file_bytes.read(8).hex() if file_size >= 8 else ""
        file_bytes.seek(0)
        
        # Check for common file signatures
        file_type = "Unknown"
        if file_header.startswith("504b0304"):
            file_type = "ZIP file (possibly .xlsx)"
        elif file_header.startswith("d0cf11e0"):
            file_type = "OLE Compound Document (.xls)"
        else:
            st.warning(f"Arquivo não parece ser um Excel. Assinatura do arquivo: {file_header}")
        
        st.write(f"Tipo de arquivo detectado: {file_type}")
        
        # Try to save and reload the file to ensure it's valid
        try:
            # Save file to temp location
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(file_bytes.getvalue())
                temp_path = tmp.name
            
            st.write(f"Arquivo salvo temporariamente em: {temp_path}")
            
            # Try to open with different methods
            try:
                st.write("Tentando abrir com openpyxl...")
                import openpyxl
                wb = openpyxl.load_workbook(temp_path, read_only=True)
                st.write(f"Sheets disponíveis (openpyxl): {wb.sheetnames}")
                if "TD Dados" in wb.sheetnames:
                    st.success("✅ Encontrada aba 'TD Dados'")
                else:
                    st.error("❌ Aba 'TD Dados' não encontrada!")
                    st.info("Sheets disponíveis: " + ", ".join(wb.sheetnames))
            except Exception as e1:
                st.warning(f"Não foi possível ler com openpyxl: {str(e1)}")
                
                # Additional reading attempts...
            
            # Clean up temp file
            import os
            os.unlink(temp_path)
            
        except Exception as e:
            st.error(f"Erro ao verificar o arquivo Excel: {str(e)}")
    
    # Provide help for preparing the file...
    
    # Process button
    if st.button("Processar"):
        st.write("Starting processing...")
        
        # Processing
        with st.spinner("Analisando..."):
            try:
                # Use the saved file data from session state
                if "file_data" in st.session_state and st.session_state.file_size > 0:
                    file_bytes = BytesIO(st.session_state.file_data)
                    file_size = st.session_state.file_size
                else:
                    st.error("Não foi possível acessar os dados do arquivo. Tente fazer o upload novamente.")
                    st.stop()
                
                st.write(f"File size for processing: {file_size} bytes")
                
                # For diagnostic - check the file header
                file_bytes.seek(0)
                header = file_bytes.read(16).hex() if file_size >= 16 else ""
                st.write(f"File header: {header}")
                file_bytes.seek(0)  # Reset again
                
                # Save to a temporary file
                import tempfile
                import os
                temp_path = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        # Write the file content
                        tmp.write(file_bytes.getvalue())
                        temp_path = tmp.name
                    
                    st.write(f"Temporary file saved at: {temp_path}")
                    
                    # Verify the file was written correctly
                    if os.path.exists(temp_path):
                        temp_size = os.path.getsize(temp_path)
                        st.write(f"Temporary file size: {temp_size} bytes")
                        
                        if temp_size == 0:
                            st.error("Erro: O arquivo temporário está vazio!")
                            raise ValueError("O arquivo temporário está vazio. Não foi possível copiar os dados do upload.")
                        
                        # Read the file directly from disk
                        with open(temp_path, "rb") as f:
                            temp_data = f.read()
                            st.write(f"Read {len(temp_data)} bytes from temp file")
                            
                            # Create NEW BytesIO from the disk read
                            process_bytes = BytesIO(temp_data)
                            
                            # Process the file
                            resultado = process_file(process_bytes)
                    
                    # Success path
                    st.success("Pronto! Baixe o arquivo destacado:")
                    st.download_button(
                        label="⬇️ Download Excel",
                        data=resultado,
                        file_name=uploaded.name.replace(".xlsx", "_highlighted.xlsx"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.balloons()
                    
                finally:
                    # Clean up temp file
                    if temp_path and os.path.exists(temp_path):
                        os.unlink(temp_path)
                        st.write("Temporary file cleaned up")
            
            except Exception as e:
                # Error handling
                st.error(f"Erro ao processar o arquivo: {str(e)}")
                
                # Show detailed error info
                error_details = traceback.format_exc()
                with st.expander("Detalhes do Erro"):
                    st.code(error_details)
                
                # Check for common errors
                error_msg = str(e).lower()
                if "not a zip file" in error_msg:
                    st.warning("⚠️ O arquivo não é um arquivo Excel (.xlsx) válido. Verifique se o arquivo não está corrompido.")
                    
                    # Additional debugging
                    st.write("Tentando diagnosticar o problema...")
                    try:
                        # Check if we can re-read the file
                        fresh_bytes = BytesIO(uploaded.read())
                        st.write(f"Upload pode ser lido novamente: {len(fresh_bytes.getvalue())} bytes")
                        
                        # Check file header
                        fresh_bytes.seek(0)
                        header = fresh_bytes.read(16).hex()
                        st.write(f"Primeiros 16 bytes: {header}")
                        
                        # Try different approach
                        import pandas as pd
                        st.write("Tentando ler com diferentes engines...")
                        
                        engines = ['openpyxl', 'xlrd', 'odf', 'pyxlsb']
                        for engine in engines:
                            try:
                                st.write(f"Tentando com engine '{engine}'...")
                                fresh_bytes.seek(0)
                                df = pd.read_excel(fresh_bytes, engine=engine)
                                st.write(f"✅ Sucesso com engine '{engine}'!")
                                break
                            except Exception as read_err:
                                st.write(f"❌ Falha com engine '{engine}': {str(read_err)}")
                    except Exception as diag_err:
                        st.write(f"Erro no diagnóstico: {str(diag_err)}")
                    
                elif "td dados" in error_msg:
                    st.warning("⚠️ Verifique se seu arquivo Excel contém uma aba chamada exatamente 'TD Dados'")
                elif "abel" in error_msg:
                    st.warning("⚠️ Verifique se a coluna A contém a palavra 'ABEL'")
                else:
                    st.warning("⚠️ Verifique o formato do arquivo de acordo com as instruções")