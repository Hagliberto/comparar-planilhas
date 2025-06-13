import pandas as pd
import numpy as np
import io
from io import BytesIO

# Tolerância para comparar floats (até 0.01 de diferença em 2 casas decimais)
TOLERANCE = 0.01

def normalize_id_column(series_id):
    """Remove espaços e zeros à esquerda de uma coluna de ID (Matrícula)."""
    if series_id.empty:
        return series_id.astype(str)
    s = series_id.astype(str).str.strip()
    s = s.str.lstrip("0")
    s = s.replace("", "0")  # Se sobra string vazia, transforma em "0"
    return s

def load_data_xlsx(uploaded_file_content, has_header):
    """Carrega dados de um arquivo .xlsx (conteúdo em bytes) para um DataFrame."""
    try:
        if has_header:
            df = pd.read_excel(io.BytesIO(uploaded_file_content), dtype=str, header=0)
        else:
            df = pd.read_excel(io.BytesIO(uploaded_file_content), header=None, dtype=str)
        return df
    except Exception as e:
        raise ValueError(f"Erro ao ler arquivo XLSX: {e}")

def preprocess_dataframe(df, has_header, filename=""):
    """
    Recebe um DataFrame bruto, extrai e limpa as colunas Matrícula, Verba e Valor.
    Ajustado para sempre pegar as 3 primeiras colunas, independente do nome do cabeçalho.
    """
    if df is None:
        return None

    if df.shape[1] < 3:
        raise ValueError(
            f"No arquivo '{filename}' a planilha precisa de pelo menos 3 colunas: Matrícula, Verba e Valor."
        )
    
    df3 = df.iloc[:, :3].copy()
    df3.columns = ['Matricula', 'Verba', 'Valor_Str']

    df3.dropna(how='all', inplace=True)
    if df3.empty:
        return df3 # Retorna DataFrame vazio, sem erro
    
    df3['Matricula'] = normalize_id_column(df3['Matricula'])
    df3['Verba'] = df3['Verba'].astype(str).str.strip()
    df3['Valor_temp'] = df3['Valor_Str'].astype(str).str.replace(',', '.', regex=False)
    df3['Valor'] = pd.to_numeric(df3['Valor_temp'], errors='coerce')
    
    invalid_mask = (
        df3['Valor'].isna()
        & df3['Valor_Str'].notna()
        & (df3['Valor_Str'].astype(str).str.strip() != '')
        & ~df3['Valor_Str'].astype(str).str.lower().isin(['nan','<na>','null','none'])
    )
    if invalid_mask.sum() > 0:
        exemplos = df3.loc[invalid_mask, 'Valor_Str'].unique()[:3]
        # Em um backend, não usamos st.warning, mas podemos logar ou retornar isso no JSON
        print(f"AVISO: No arquivo '{filename}' foram encontradas dados em 'Valor' que não puderam ser convertidas: Exemplos: {', '.join(exemplos)}")

    df3.drop(columns=['Valor_temp', 'Valor_Str'], inplace=True)
    return df3[['Matricula', 'Verba', 'Valor']]

def compare_dataframes(df1, df2):
    """
    Compara dois DataFrames e retorna um DataFrame com as diferenças.
    """
    merged_df = pd.merge(df1, df2, on=["Matricula", "Verba"], how="outer", suffixes=["_1", "_2"], indicator=True)

    # Identificar linhas que estão apenas em um dos DFs
    only_in_df1 = merged_df["_merge"] == "left_only"
    only_in_df2 = merged_df["_merge"] == "right_only"

    # Identificar valores divergentes (presentes em ambos, mas com diferença > TOLERANCE)
    # Preencher NaNs com 0 para cálculo da diferença, mas manter NaNs para exibição
    merged_df["Valor_1_filled"] = merged_df["Valor_1"].fillna(0)
    merged_df["Valor_2_filled"] = merged_df["Valor_2"].fillna(0)

    value_divergence = (
        (merged_df["_merge"] == "both")
        & (abs(merged_df["Valor_1_filled"] - merged_df["Valor_2_filled"]) > TOLERANCE)
    )

    # Criar o DataFrame de diferenças
    diffs = merged_df[
        only_in_df1 | only_in_df2 | value_divergence
    ].copy()

    diffs["Status"] = ""
    diffs.loc[only_in_df1, "Status"] = "Apenas na Planilha 1"
    diffs.loc[only_in_df2, "Status"] = "Apenas na Planilha 2"
    diffs.loc[value_divergence, "Status"] = "Valor Divergente"

    # Calcular a diferença para valores divergentes
    diffs["Diferença"] = diffs["Valor_1"] - diffs["Valor_2"]

    # Remover colunas temporárias
    diffs.drop(columns=["Valor_1_filled", "Valor_2_filled"], inplace=True)

    return diffs

def create_diff_excel(df1, df2, df_diffs_raw):
    """Gera o arquivo Excel completo com abas de Estatísticas, Lado a Lado e Diferenças."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Define specific hex colors for easier reference
        COLOR_POSITIVE = '#C6EFCE' # Light Green
        COLOR_NEGATIVE = '#FFC7CE' # Light Red
        COLOR_MISSING = '#FFEB9C'  # Light Yellow

        # Add a number format for 2 decimal places
        decimal_format = workbook.add_format({'num_format': '#,##0.00'})        
        # Formats for data cells with background colors and decimal formatting
        format_positive_diff_cell = workbook.add_format({'bg_color': COLOR_POSITIVE, 'num_format': '#,##0.00'}) # Light Green
        format_negative_diff_cell = workbook.add_format({'bg_color': COLOR_NEGATIVE, 'num_format': '#,##0.00'}) # Light Red
        format_missing_data_cell_num = workbook.add_format({'bg_color': COLOR_MISSING, 'num_format': '#,##0.00'}) # Light Yellow
        format_default_text = workbook.add_format() # Default format for text without color

        # Aba "Estatísticas"
        summary = pd.DataFrame({
            'Métrica': [
                'Linhas (Planilha 1)', 'Linhas (Planilha 2)',
                'Matrículas únicas (Planilha 1)', 'Matrículas únicas (Planilha 2)',
                'Verbas únicas (Planilha 1)', 'Verbas únicas (Planilha 2)'
            ],
            'Valor': [
                len(df1), len(df2), df1['Matricula'].nunique(), df2['Matricula'].nunique(),
                df1['Verba'].nunique(), df2['Verba'].nunique()
            ]
        })
        summary.to_excel(writer, sheet_name='Estatísticas', index=False)
        
        # Aba "Lado a Lado"
        merged_all_excel = pd.merge(
            df1.copy(), df2.copy(), on=['Matricula', 'Verba'],
            how='outer', suffixes=('_1', '_2')
        )
        merged_all_excel['Diferença'] = (merged_all_excel['Valor_1'] - merged_all_excel['Valor_2']).round(2)
        lado_df_excel = pd.DataFrame({
            'Matricula (P1)': merged_all_excel['Matricula'], 'Verba (P1)': merged_all_excel['Verba'],
            'Valor Planilha 1': merged_all_excel['Valor_1'], 'Matricula (P2)': merged_all_excel['Matricula'],
            'Verba (P2)': merged_all_excel['Verba'], 'Valor Planilha 2': merged_all_excel['Valor_2'],
            'Diferença': merged_all_excel['Diferença']
        })
        
        # We write column headers manually to control format
        header_lado_a_lado = list(lado_df_excel.columns)
        worksheet_lado_a_lado = workbook.add_worksheet('Lado a Lado')
        for col_num, value in enumerate(header_lado_a_lado):
            worksheet_lado_a_lado.write(0, col_num, value)

        # Set column width for better visibility
        worksheet_lado_a_lado.set_column('A:A', 15) # 'Matricula (P1)'
        worksheet_lado_a_lado.set_column('B:B', 15) # 'Verba (P1)'
        worksheet_lado_a_lado.set_column('C:C', 15) # 'Valor Planilha 1'
        worksheet_lado_a_lado.set_column('D:D', 15) # 'Matricula (P2)'
        worksheet_lado_a_lado.set_column('E:E', 15) # 'Verba (P2)'
        worksheet_lado_a_lado.set_column('F:F', 15) # 'Valor Planilha 2'
        worksheet_lado_a_lado.set_column('G:G', 15) # 'Diferença'

        for row_num, row_data in enumerate(lado_df_excel.itertuples(index=False)):
            val1 = row_data[2] # Valor Planilha 1
            val2 = row_data[5] # Valor Planilha 2
            diff = row_data[6] # Diferença
            
            current_bg_color = None
            
            # Check for missing data (one is NaN and the other is not)
            if (pd.isna(val1) and pd.notna(val2)) or (pd.notna(val1) and pd.isna(val2)):
                current_bg_color = COLOR_MISSING
            # Check for value divergence
            elif pd.notna(val1) and pd.notna(val2) and abs(round(val1, 2) - round(val2, 2)) > TOLERANCE:
                if diff > 0:
                    current_bg_color = COLOR_POSITIVE
                elif diff < 0:
                    current_bg_color = COLOR_NEGATIVE
            
            # Write data to cells with appropriate format
            for col_idx, cell_value in enumerate(row_data):
                write_format = None
                if current_bg_color:
                    if col_idx in [2, 5, 6]: # Value columns
                        # Apply specific colored format with number format
                        if current_bg_color == COLOR_POSITIVE:
                            write_format = format_positive_diff_cell
                        elif current_bg_color == COLOR_NEGATIVE:
                            write_format = format_negative_diff_cell
                        else: # COLOR_MISSING
                            write_format = format_missing_data_cell_num
                        
                        # Handle NaN values for numeric cells
                        if pd.isna(cell_value):
                            worksheet_lado_a_lado.write(row_num + 1, col_idx, '', workbook.add_format({'bg_color': current_bg_color}))
                        else:
                            worksheet_lado_a_lado.write(row_num + 1, col_idx, cell_value, write_format)
                    else: # Text columns like Matricula, Verba
                        worksheet_lado_a_lado.write(row_num + 1, col_idx, cell_value, workbook.add_format({'bg_color': current_bg_color}))
                else: # No highlight
                    if col_idx in [2, 5, 6]: # Value columns
                        worksheet_lado_a_lado.write(row_num + 1, col_idx, cell_value, decimal_format)
                    else: # Text columns
                        worksheet_lado_a_lado.write(row_num + 1, col_idx, cell_value, format_default_text)


        # Aba "Diferenças" (Only non-zero differences and missing data)
        # Calculate 'Diferença' column in df_diffs_raw if not already present
        # and rename columns for consistency before filtering for Excel output.
        df_diffs_for_excel_tab = df_diffs_raw.copy()
        
        # Ensure 'Valor Planilha 1' and 'Valor Planilha 2' are present, or rename them if they're still 'Valor_1'/'Valor_2'
        if 'Valor_1' in df_diffs_for_excel_tab.columns and 'Valor_2' in df_diffs_for_excel_tab.columns:
             df_diffs_for_excel_tab.rename(columns={'Valor_1': 'Valor Planilha 1', 'Valor_2': 'Valor Planilha 2'}, inplace=True)
        
        # Ensure 'Diferença' column exists and is rounded
        if 'Diferença' not in df_diffs_for_excel_tab.columns:
            if 'Valor Planilha 1' in df_diffs_for_excel_tab.columns and 'Valor Planilha 2' in df_diffs_for_excel_tab.columns:
                df_diffs_for_excel_tab['Diferença'] = (df_diffs_for_excel_tab['Valor Planilha 1'] - df_diffs_for_excel_tab['Valor Planilha 2']).round(2)
            else:
                df_diffs_for_excel_tab['Diferença'] = np.nan # Fallback
        else: # If 'Diferença' already exists, ensure it's rounded
             df_diffs_for_excel_tab['Diferença'] = df_diffs_for_excel_tab['Diferença'].round(2)


        # Filter for actual differences (non-zero or missing data)
        df_final_diffs_for_excel = df_diffs_for_excel_tab[
            (df_diffs_for_excel_tab['Status'] == 'Apenas na Planilha 1') |
            (df_diffs_for_excel_tab['Status'] == 'Apenas na Planilha 2') |
            (
                (df_diffs_for_excel_tab['Status'] == 'Valor Divergente') &
                (df_diffs_for_excel_tab['Diferença'].abs() > TOLERANCE) # Use absolute difference for non-zero check
            )
        ].copy()
        
        # Drop the _merge column before writing to Excel
        if '_merge' in df_final_diffs_for_excel.columns:
            df_final_diffs_for_excel = df_final_diffs_for_excel.drop(columns=['_merge'])


        if not df_final_diffs_for_excel.empty:
            # We write column headers manually to control format
            header_diffs = list(df_final_diffs_for_excel.columns)
            worksheet_diffs = workbook.add_worksheet('Diferenças')
            for col_num, value in enumerate(header_diffs):
                worksheet_diffs.write(0, col_num, value)

            # Set column width for better visibility
            worksheet_diffs.set_column('A:A', 15) # 'Matricula'
            worksheet_diffs.set_column('B:B', 15) # 'Verba'
            worksheet_diffs.set_column('C:C', 15) # 'Valor Planilha 1'
            worksheet_diffs.set_column('D:D', 15) # 'Valor Planilha 2'
            worksheet_diffs.set_column('E:E', 20) # 'Status'
            worksheet_diffs.set_column('F:F', 15) # 'Diferença'


            for row_num, row_data in enumerate(df_final_diffs_for_excel.itertuples(index=False)):
                # Get values by column name for clarity and robustness
                val1 = row_data[df_final_diffs_for_excel.columns.get_loc('Valor Planilha 1')]
                val2 = row_data[df_final_diffs_for_excel.columns.get_loc('Valor Planilha 2')]
                diff = row_data[df_final_diffs_for_excel.columns.get_loc('Diferença')]
                status = row_data[df_final_diffs_for_excel.columns.get_loc('Status')]

                current_bg_color = None
                if status == 'Apenas na Planilha 1' or status == 'Apenas na Planilha 2':
                    current_bg_color = COLOR_MISSING
                elif status == 'Valor Divergente':
                    if diff > 0:
                        current_bg_color = COLOR_POSITIVE
                    elif diff < 0:
                        current_bg_color = COLOR_NEGATIVE
                
                # Write data to cells with appropriate format
                for col_idx, cell_value_raw in enumerate(row_data):
                    col_name = header_diffs[col_idx] # Get column name from header list

                    if current_bg_color:
                        if col_name in ['Valor Planilha 1', 'Valor Planilha 2', 'Diferença']:
                            # Apply specific colored format with number format
                            if current_bg_color == COLOR_POSITIVE:
                                write_format = format_positive_diff_cell
                            elif current_bg_color == COLOR_NEGATIVE:
                                write_format = format_negative_diff_cell
                            else: # COLOR_MISSING
                                write_format = format_missing_data_cell_num
                            
                            # Handle NaN values for numeric cells
                            if pd.isna(cell_value_raw):
                                worksheet_diffs.write(row_num + 1, col_idx, '', workbook.add_format({'bg_color': current_bg_color}))
                            else:
                                worksheet_diffs.write(row_num + 1, col_idx, cell_value_raw, write_format)
                        else: # Text columns like Matricula, Verba, Status
                            worksheet_diffs.write(row_num + 1, col_idx, cell_value_raw, workbook.add_format({'bg_color': current_bg_color}))
                    else: # No highlight
                        if col_name in ['Valor Planilha 1', 'Valor Planilha 2', 'Diferença']:
                            worksheet_diffs.write(row_num + 1, col_idx, cell_value_raw, decimal_format)
                        else: # Text columns
                            worksheet_diffs.write(row_num + 1, col_idx, cell_value_raw, format_default_text)


    output.seek(0)
    return output.getvalue()

