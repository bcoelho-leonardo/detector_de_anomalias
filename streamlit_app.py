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
# File uploader
uploaded = st.file_uploader("Escolher arquivo Excel", type=["xlsx", "xls"])

# When a file is uploaded
if uploaded:
    st.write(f"File '{uploaded.name}' uploaded successfully")
    st.write(f"File type: {type(uploaded)}")
    
    # Create an expander for showing file info
    with st.expander("Informações do Arquivo"):
        # Check file format and try to determine what it actually is
        file_bytes = BytesIO(uploaded.read())
        file_size = len(file_bytes.getvalue())
        st.write(f"Tamanho do arquivo: {file_size} bytes")
        
        # Basic file signature check
        file_bytes.seek(0)
        file_header = file_bytes.read(8).hex()
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
                
                try:
                    st.write("Tentando abrir com xlrd...")
                    import xlrd
                    xls = xlrd.open_workbook(temp_path)
                    st.write(f"Sheets disponíveis (xlrd): {xls.sheet_names()}")
                    if "TD Dados" in xls.sheet_names():
                        st.success("✅ Encontrada aba 'TD Dados'")
                    else:
                        st.error("❌ Aba 'TD Dados' não encontrada!")
                        st.info("Sheets disponíveis: " + ", ".join(xls.sheet_names()))
                except Exception as e2:
                    st.warning(f"Não foi possível ler com xlrd: {str(e2)}")
                    
                    # Try pandas directly
                    try:
                        st.write("Tentando abrir com pandas...")
                        import pandas as pd
                        xls = pd.ExcelFile(temp_path)
                        st.write(f"Sheets disponíveis (pandas): {xls.sheet_names}")
                        if "TD Dados" in xls.sheet_names:
                            st.success("✅ Encontrada aba 'TD Dados'")
                        else:
                            st.error("❌ Aba 'TD Dados' não encontrada!")
                            st.info("Sheets disponíveis: " + ", ".join(xls.sheet_names))
                    except Exception as e3:
                        st.error(f"Todos os métodos de leitura falharam: {str(e3)}")
            
            # Clean up temp file
            import os
            os.unlink(temp_path)
            
        except Exception as e:
            st.error(f"Erro ao verificar o arquivo Excel: {str(e)}")
        
        # Reset for future use
        file_bytes.seek(0)
    
    # Provide help for preparing the file
    with st.expander("Como preparar seu arquivo Excel?"):
        st.markdown("""
        ## Requisitos para o arquivo Excel:
        
        1. **Formato**: Arquivo Excel (.xlsx) válido e não corrompido
        2. **Aba**: Deve conter uma aba chamada exatamente "TD Dados"
        3. **Estrutura**:
            - Coluna A deve conter a palavra "ABEL" (identificador de início dos dados)
            - Uma linha de cabeçalho deve estar acima da linha com "ABEL"
            - Colunas com datas devem estar no formato "YYYY-MM" (exemplo: 2025-05)
        
        Se seu arquivo não atende a esses requisitos, por favor ajuste-o e tente novamente.
        """)
    
    # Process button
    if st.button("Processar"):
        st.write("Starting processing...")
        
        # Processing
        with st.spinner("Analisando..."):
            try:
                # Read file again to get fresh BytesIO
                file_bytes = BytesIO(uploaded.read())
                
                # Try to diagnose potential file issues
                file_bytes.seek(0)
                header = file_bytes.read(8).hex()
                if not (header.startswith("504b0304") or header.startswith("d0cf11e0")):
                    st.warning("⚠️ O arquivo não parece ser um Excel válido. Verifique o formato do arquivo.")
                
                # Reset pointer and process
                file_bytes.seek(0)
                
                # Write file to temp location for better reliability
                import tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(file_bytes.getvalue())
                    temp_path = tmp.name
                
                # Use the file path instead of BytesIO for more reliable processing
                with open(temp_path, "rb") as f:
                    resultado = process_file(BytesIO(f.read()))
                
                # Clean up temp file
                import os
                os.unlink(temp_path)
                
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
                if "not a zip file" in error_msg:
                    st.warning("⚠️ O arquivo não é um arquivo Excel (.xlsx) válido. Verifique se o arquivo não está corrompido.")
                elif "td dados" in error_msg:
                    st.warning("⚠️ Verifique se seu arquivo Excel contém uma aba chamada exatamente 'TD Dados'")
                elif "abel" in error_msg:
                    st.warning("⚠️ Verifique se a coluna A contém a palavra 'ABEL'")
                else:
                    st.warning("⚠️ Verifique o formato do arquivo de acordo com as instruções")