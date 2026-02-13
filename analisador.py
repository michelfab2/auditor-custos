import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- Configuração da Página ---
st.set_page_config(page_title="Auditor de Custos (Sniper)", layout="wide")
st.title("🛡️ Auditor de Propostas (Visual Sniper)")
st.markdown("""
**Melhoria Visual:**
* As cores agora aparecem **apenas** nas células onde o erro ocorreu.
* Se o erro é Quantidade, só a Quantidade brilha.
* Mantida a ordem: **Referência (Esq) -> Proposta (Dir)**.
""")

# --- Funções de Limpeza (Robustas) ---
def clean_code(code):
    if pd.isna(code): return None
    code_str = str(code).strip()
    nums = ''.join(filter(str.isdigit, code_str))
    if nums: return str(int(nums))
    return code_str

def clean_float(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('R$', '').replace(' ', '')
    if ',' in val_str and '.' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

# --- Processamento ---
def parse_base(df):
    data = []
    current_parent = None
    for idx, row in df.iterrows():
        try:
            row_type = str(row.iloc[0]).strip()
            code_raw = row.iloc[3]
            if pd.isna(code_raw): continue
            code = clean_code(code_raw)
            
            if row_type == 'Composição':
                current_parent = code
                continue 
            
            if current_parent and row_type in ['Insumo', 'Composição Auxiliar']:
                data.append({
                    'PARENT_CODE': current_parent,
                    'ITEM_CODE': code,
                    'DESC_REF': str(row.iloc[5]),
                    'UND_REF': str(row.iloc[6]),
                    'COEF_REF': clean_float(row.iloc[7]),
                    'UNIT_PRICE_REF': clean_float(row.iloc[8]),
                    'TOTAL_REF': clean_float(row.iloc[9])
                })
        except: continue
    return pd.DataFrame(data)

def parse_empresa(df):
    data = []
    current_parent = None
    col_map = {c.upper().strip(): c for c in df.columns}
    
    if 'ITEM' not in col_map: return pd.DataFrame()

    for idx, row in df.iterrows():
        item_val = str(row[col_map.get('ITEM', 'ITEM')]).strip()
        code_raw = row.get(col_map.get('CÓDIGOS', 'CÓDIGOS'))
        if pd.isna(code_raw): continue
        code = clean_code(code_raw)
        
        is_parent = item_val != '-' and item_val != '' and item_val.lower() != 'nan'
        if is_parent:
            current_parent = code
            continue 
        
        if current_parent:
            data.append({
                'PARENT_CODE': current_parent,
                'ITEM_CODE': code,
                'DESC_PROP': row.get(col_map.get('DESCRIÇÃO', 'DESCRIÇÃO')),
                'UND_PROP': row.get(col_map.get('UND', 'UND')),
                'COEF_PROP': clean_float(row.get(col_map.get('COEF', 'COEF'))),
                'UNIT_PRICE_PROP': clean_float(row.get(col_map.get('PREÇO UNIT', 'PREÇO UNIT'))),
                'TOTAL_PROP': clean_float(row.get(col_map.get('PREÇO TOTAL', 'PREÇO TOTAL')))
            })
    return pd.DataFrame(data)

# --- Interface ---
col1, col2 = st.columns(2)
file_base = col1.file_uploader("📂 Planilha BASE (Referência)", type=['xlsx', 'csv'])
file_empresa = col2.file_uploader("📂 Planilha EMPRESA (Proposta)", type=['xlsx', 'csv'])

if file_base and file_empresa:
    with st.spinner("Realizando auditoria de precisão..."):
        try:
            # Leitura
            try: df_base_raw = pd.read_excel(file_base)
            except: 
                file_base.seek(0)
                df_base_raw = pd.read_csv(file_base, header=None)
            
            try: df_empresa_raw = pd.read_excel(file_empresa)
            except: 
                file_empresa.seek(0)
                df_empresa_raw = pd.read_csv(file_empresa)

            df_base_clean = parse_base(df_base_raw)
            df_empresa_clean = parse_empresa(df_empresa_raw)
            
            if df_base_clean.empty or df_empresa_clean.empty:
                st.error("Erro: Dados não identificados.")
            else:
                # Merge
                merged = pd.merge(
                    df_empresa_clean,
                    df_base_clean,
                    on=['PARENT_CODE', 'ITEM_CODE'],
                    how='left', 
                    suffixes=('_PROP', '_REF')
                )
                
                # Cálculos
                merged['VAR_COEF_%'] = ((merged['COEF_PROP'] - merged['COEF_REF']) / merged['COEF_REF'].replace(0, 1)) * 100
                merged['VAR_PRECO_%'] = ((merged['UNIT_PRICE_PROP'] - merged['UNIT_PRICE_REF']) / merged['UNIT_PRICE_REF'].replace(0, 1)) * 100
                
                def get_status(row):
                    status = []
                    # Unidade
                    if str(row['UND_REF']) != 'nan' and str(row['UND_REF']).strip() != str(row['UND_PROP']).strip():
                        status.append("UND DIFERENTE")
                    # Coeficiente (1% tol)
                    if abs(row['VAR_COEF_%']) > 1.0:
                        status.append("QTD ALTERADA")
                    # Preço
                    if row['UNIT_PRICE_PROP'] > row['UNIT_PRICE_REF']:
                        status.append("SOBREPREÇO")
                    # Desconto Excessivo
                    if row['UNIT_PRICE_REF'] > 0 and row['UNIT_PRICE_PROP'] < (0.70 * row['UNIT_PRICE_REF']):
                        status.append("DESC. SUSPEITO")
                        
                    if pd.isna(row['COEF_REF']):
                        return "ITEM EXTRA"
                        
                    return "OK" if not status else " | ".join(status)

                merged['STATUS'] = merged.apply(get_status, axis=1)
                
                # Ordenação de Colunas
                cols_order = [
                    'PARENT_CODE', 'ITEM_CODE',
                    'DESC_REF', 'UND_REF', 'COEF_REF', 'UNIT_PRICE_REF', 'TOTAL_REF',
                    'DESC_PROP', 'UND_PROP', 'COEF_PROP', 'UNIT_PRICE_PROP', 'TOTAL_PROP',
                    'VAR_COEF_%', 'VAR_PRECO_%', 'STATUS'
                ]
                final_cols = [c for c in cols_order if c in merged.columns]
                df_final = merged[final_cols]
                
                st.write(f"Auditoria concluída: {len(df_final)} itens processados.")

                # --- Exportação Excel com Formatação Condicional "Sniper" ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, sheet_name='Auditoria', index=False)
                    
                    wb = writer.book
                    ws = writer.sheets['Auditoria']
                    
                    # Formatos de Destaque
                    fmt_num = wb.add_format({'num_format': '#,##0.00'})
                    
                    # Cores de Alerta
                    fmt_red = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})     # Sobrepreço
                    fmt_yellow = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})  # Qtd Alterada
                    fmt_orange = wb.add_format({'bg_color': '#FFD966', 'font_color': '#333333'})  # Und Diferente
                    fmt_blue = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})    # Desconto
                    
                    # Formata colunas numéricas (Loop em todas as colunas)
                    for i, col in enumerate(df_final.columns):
                        if any(x in col for x in ['COEF', 'PRICE', 'TOTAL', 'VAR']):
                            ws.set_column(i, i, 12, fmt_num)
                        elif 'DESC' in col:
                            ws.set_column(i, i, 40)
                    
                    # --- Lógica de Condicional ---
                    # Identificar Índices das Colunas Importantes
                    def get_col_idx(name):
                        return df_final.columns.get_loc(name) if name in df_final.columns else -1
                    
                    idx_status = get_col_idx('STATUS')
                    
                    # Letra da coluna de Status (ex: $O2) para usar na fórmula
                    col_status_letter = xl_col_to_name(idx_status) 
                    
                    last_row = len(df_final) + 1
                    
                    # Helper para aplicar formato condicional
                    def apply_cond(col_name_list, trigger_text, fmt):
                        for col_name in col_name_list:
                            idx = get_col_idx(col_name)
                            if idx != -1:
                                # Fórmula: Se a coluna STATUS contiver o texto X, pinte esta célula
                                # Ex: =ISNUMBER(SEARCH("SOBREPREÇO", $O2))
                                formula = f'=ISNUMBER(SEARCH("{trigger_text}", ${col_status_letter}2))'
                                ws.conditional_format(1, idx, last_row, idx, {
                                    'type': 'formula',
                                    'criteria': formula,
                                    'format': fmt
                                })

                    # 1. SOBREPREÇO -> Pinta Unitário Ref e Prop (Vermelho)
                    apply_cond(['UNIT_PRICE_REF', 'UNIT_PRICE_PROP'], 'SOBREPREÇO', fmt_red)
                    
                    # 2. QTD ALTERADA -> Pinta Coef Ref e Prop (Amarelo)
                    apply_cond(['COEF_REF', 'COEF_PROP'], 'QTD', fmt_yellow)
                    
                    # 3. UND DIFERENTE -> Pinta Und Ref e Prop (Laranja)
                    apply_cond(['UND_REF', 'UND_PROP'], 'UND', fmt_orange)
                    
                    # 4. DESC SUSPEITO -> Pinta Unitário Prop (Azul/Verde)
                    apply_cond(['UNIT_PRICE_PROP'], 'DESC', fmt_blue)
                    
                    # 5. ITEM EXTRA -> Pinta a descrição da Proposta
                    apply_cond(['DESC_PROP'], 'EXTRA', fmt_yellow)

                st.download_button("📥 Baixar Planilha (Destaque Inteligente)", buffer.getvalue(), "Auditoria_Sniper.xlsx")

        except Exception as e:
            st.error(f"Erro: {e}")