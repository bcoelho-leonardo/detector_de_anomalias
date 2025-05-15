# -*- coding: utf-8 -*-
"""
Created on Thu May 15 20:42:05 2025

@author: Admin
"""

# detector_de_anomalias_Claude.py
# -*- coding: utf-8 -*-
"""
Versão modular para uso em Streamlit:
------------------------------------
• process_file(file_like: BytesIO)  -> bytes
  - Recebe um objeto BytesIO contendo o Excel original
  - Executa toda a lógica de detecção (LOF + dados ausentes)
  - Devolve os bytes do novo arquivo *_highlighted.xlsx

Obs.: continua possível rodar em modo CLI (`python detector_de_anomalias_Claude.py`)
"""

import os
import numpy as np
import pandas as pd
import traceback
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from sklearn.neighbors import LocalOutlierFactor

###############################################################################
# 1) Funções auxiliares (inalteradas, exceto sem prints excessivos)
###############################################################################
def dynamic_n_neighbors(row_length, frac=0.15, min_n=2, max_n=50):
    raw_nn = int(frac * row_length)
    return max(min_n, min(raw_nn, max_n))


def detect_outliers(
    df,
    date_cols,
    last_month_idx=-1,
    frac=0.15,
    min_n=2,
    max_n=50,
    fixed_cont=0.05,
    missing_threshold=0.2
):
    lof_outliers, missing_outliers = [], []
    col_name_last = date_cols[-1]

    # % de faltantes por linha
    miss_pct = df[date_cols].isnull().mean(axis=1)

    for r_idx, (_, row) in enumerate(df.iterrows()):
        y_vals = pd.to_numeric(row[date_cols], errors="coerce").values

        # 1) faltantes incomuns
        if np.isnan(y_vals[-1]) and miss_pct.iloc[r_idx] < missing_threshold:
            missing_outliers.append(r_idx)
            continue

        # 2) LOF
        if not np.isnan(y_vals[-1]):
            x_num = np.arange(len(date_cols))
            data_2d = np.column_stack([x_num, y_vals])

            valid = ~np.isnan(y_vals)
            data_2d = data_2d[valid]
            if len(data_2d) < 2:
                continue

            lof = LocalOutlierFactor(
                n_neighbors=dynamic_n_neighbors(len(data_2d), frac, min_n, max_n),
                contamination=fixed_cont,
                metric="minkowski",
                p=1,
            )
            labels = lof.fit_predict(data_2d)
            if labels[-1] == -1:  # última posição é outlier
                lof_outliers.append(r_idx)

    return lof_outliers, missing_outliers


def highlight_workbook_in_memory(
    wb,
    sheet_name,
    lof_rows,
    miss_rows,
    excel_row_offset,
    excel_col_idx=2,  # B
):
    ws = wb[sheet_name]
    yellow = PatternFill("solid", fgColor="FFFF00")
    red = PatternFill("solid", fgColor="FF0000")

    for r in lof_rows:
        ws.cell(row=r + excel_row_offset, column=excel_col_idx).fill = yellow
    for r in miss_rows:
        ws.cell(row=r + excel_row_offset, column=excel_col_idx).fill = red

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


###############################################################################
# 2) Função principal para Streamlit
###############################################################################
def process_file(file_like):
    """
    Parameters
    ----------
    file_like : BytesIO
        Conteúdo do Excel original.

    Returns
    -------
    bytes
        Bytes do arquivo *_highlighted.xlsx.
    """
    try:
        # -------- 1. Leitura bruta (sem header) ----------
        raw_df = pd.read_excel(file_like, sheet_name="TD Dados", header=None)

        # -------- 2. Identificar cabeçalho via "ABEL" ----
        mask_abel = raw_df[0].astype(str).str.strip().str.upper() == "ABEL"
        if not mask_abel.any():
            raise ValueError("String 'ABEL' não encontrada na coluna A.")

        start_row = mask_abel.idxmax()
        if start_row == 0:
            raise ValueError("Não há linha de cabeçalho acima de 'ABEL'.")

        header_row = start_row - 1
        if str(raw_df.iloc[start_row - 1, 0]).strip().upper() == "UNIDADE 1":
            header_row = start_row - 2  # cabeçalho uma linha acima

        df = (
            raw_df.iloc[header_row:]
            .reset_index(drop=True)
        )
        df.columns = df.iloc[0]
        df = df.iloc[1:]
        df = df.set_index(df.columns[0])
        df = df.iloc[:, :-1]  # remove última coluna (total geral)

        # -------- 3. Filtrar colunas de data ------------
        valid_cols = [
            c for c in df.columns
            if "total geral" not in str(c).lower()
            and "total général" not in str(c).lower()
        ]
        dates = pd.to_datetime(valid_cols, format="%Y-%m", errors="coerce")
        col_sorted = [
            c for c, d in sorted(zip(valid_cols, dates), key=lambda x: x[1] or pd.Timestamp.min)
        ]
        if len(col_sorted) > 30:
            col_sorted = col_sorted[-30:]
        df = df[col_sorted]

        # -------- 4. Detectar outliers ------------------
        lof_rows, miss_rows = detect_outliers(
            df=df,
            date_cols=col_sorted,
            fixed_cont=0.05,
            missing_threshold=0.2,
        )

        # -------- 5. Destacar no Excel em memória -------
        file_like.seek(0)
        wb = load_workbook(file_like, data_only=True)
        offset = header_row + 2  # header + 1 (0-based → Excel 1-based)
        output_bytes = highlight_workbook_in_memory(
            wb=wb,
            sheet_name="TD Dados",
            lof_rows=lof_rows,
            miss_rows=miss_rows,
            excel_row_offset=offset,
            excel_col_idx=2,
        )
        return output_bytes

    except Exception as exc:
        print("Erro em process_file:", exc)
        traceback.print_exc()
        raise


###############################################################################
# 3) Modo CLI opcional
###############################################################################
if __name__ == "__main__":
    path = input("Caminho do .xlsx: ").strip('"').strip("'")
    if not (os.path.isfile(path) and path.lower().endswith(".xlsx")):
        print("Arquivo inválido.")
        exit(1)
    with open(path, "rb") as f:
        out_bytes = process_file(BytesIO(f.read()))
    out_path = os.path.splitext(path)[0] + "_highlighted.xlsx"
    with open(out_path, "wb") as f:
        f.write(out_bytes)
    print("Arquivo gerado:", out_path)
