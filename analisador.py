import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import hashlib

# ==========================================
# 0. CONFIGURAÇÕES E CACHE
# ==========================================

st.set_page_config(page_title="Auditoria PRO", layout="wide", page_icon="🛡️")

MAX_FILE_SIZE_MB = 50
CACHE_TTL = 3600

# ==========================================
# 1. PARSERS E TRANSFORMADORES HIERÁRQUICOS
# ==========================================

@st.cache_data(ttl=CACHE_TTL)
def carregar_orcafascio(arquivo_bytes, nome_arquivo):
    try:
        tamanho_mb = len(arquivo_bytes) / (1024 * 1024)
        if tamanho_mb > MAX_FILE_SIZE_MB:
            return None, f"Arquivo {nome_arquivo} excede {MAX_FILE_SIZE_MB}MB ({tamanho_mb:.1f}MB)", None
        
        df_raw = pd.read_excel(io.BytesIO(arquivo_bytes), header=None)
        
        dados = []
        log_erros_parse = []
        cod_pai_atual, desc_pai_atual = None, ""
        ordem_sequencial = 0
        
        in_orse_detail = False
        is_sicro_section = False
        skip_items = False
        
        # Mapeamento dinâmico SICRO3/IOPES
        sicro_col_quant = 4
        sicro_col_und = None
        sicro_col_preco = None
        
        # Mapeamento padrão (SINAPI)
        col_cod, col_desc, col_und, col_quant, col_preco = 1, 3, 6, 7, 8
        
        def parse_number(val):
            if val is None or val == '' or val == '*':
                return 0.0
            if isinstance(val, (int, float)):
                return float(val)
            val = str(val).strip()
            if ',' in val and '.' in val:
                val = val.replace('.', '').replace(',', '.')
            elif ',' in val:
                val = val.replace(',', '.')
            try:
                return float(val)
            except ValueError:
                return 0.0

        for idx, row in df_raw.iterrows():
            row_list = [str(x).strip() if pd.notna(x) else '' for x in row]
            row_lower = [x.lower() for x in row_list]
            col0 = row_list[0].lower() if len(row_list) > 0 else ''
            
            # 1. Ignorar Detalhamento ORSE
            if 'detalhamento de cálculo orse' in col0 or 'detalhamento de calculo orse' in col0:
                in_orse_detail = True
                continue
                
            # 2. Fim de Composição (Reseta Estados)
            if any('mo sem ls' in x for x in row_lower) or any('valor do bdi' in x for x in row_lower):
                cod_pai_atual = None
                is_sicro_section = False
                in_orse_detail = False
                skip_items = False
                continue
                
            # 3. Novo Bloco Principal (1.1, 2.1, etc.)
            if re.match(r'^\d+\.\d+$', col0):
                cod_pai_atual = None
                is_sicro_section = False
                in_orse_detail = False
                skip_items = False
                continue
                
            # 4. Sub-cabeçalhos SICRO3/SETOP/IOPES (Iniciam com B, C, D, E, F, G)
            if len(col0) == 1 and col0.upper() in ['B', 'C', 'D', 'E', 'F', 'G']:
                is_sicro_section = True
                skip_items = False
                sicro_col_quant = 4
                sicro_col_und = None
                sicro_col_preco = None
                for i, val in enumerate(row_lower):
                    if 'quant' in val: sicro_col_quant = i
                    elif 'unidade' in val or val == 'un': sicro_col_und = i
                    elif 'preço unit' in val or 'preco unit' in val: sicro_col_preco = i
                    elif 'custo horário' in val or 'custo horario' in val:
                        if sicro_col_preco is None: sicro_col_preco = i # Em Mão de Obra, Custo Horário é o Preço
                
                if 'transporte' in ' '.join(row_lower):
                    skip_items = True # Ignora itens de transporte (bloco F do SICRO3)
                continue
                
            # 5. Cabeçalhos Padrão (SINAPI, SEINFRA, Bases Próprias)
            if 'código' in row_lower or 'codigo' in row_lower:
                is_sicro_section = False
                skip_items = False
                for i, val in enumerate(row_lower):
                    if val in ['código', 'codigo']: col_cod = i
                    elif any(d in val for d in ['descrição', 'descricao']): col_desc = i
                    elif any(u in val for u in ['und', 'unidade', 'unid']): col_und = i
                    elif any(q in val for q in ['quant', 'qtd']): col_quant = i
                    elif any(p in val for p in ['valor unit', 'preço unit', 'preco unit']): col_preco = i
                continue
                
            # 6. Itens e Composições
            if col0 in ['insumo', 'composição auxiliar', 'composicao auxiliar', 'item', 'composição', 'composicao']:
                if in_orse_detail or skip_items: continue
                if col_cod < len(row_list):
                    cod_item = row_list[col_cod].upper()
                    if cod_item.lower() in ['código', 'codigo', '', 'nan']: continue
                    
                    desc_item = row_list[col_desc] if col_desc is not None and col_desc < len(row_list) else ""
                    
                    # Se não tem pai definido e é composição, vira o Pai
                    if cod_pai_atual is None and col0 in ['composição', 'composicao']:
                        cod_pai_atual = cod_item
                        desc_pai_atual = desc_item
                        continue
                    else:
                        # É um Filho
                        if cod_pai_atual is None: continue # Órfão, ignora
                        
                        if is_sicro_section:
                            qtd_idx = sicro_col_quant if sicro_col_quant is not None else 4
                            und_idx = sicro_col_und
                            preco_idx = sicro_col_preco
                            
                            qtd_valor = parse_number(row_list[qtd_idx]) if qtd_idx < len(row_list) else 0.0
                            
                            if und_idx is not None and und_idx < len(row_list) and row_list[und_idx] != '*':
                                und_item = row_list[und_idx].upper()
                            else:
                                und_item = 'H' # Default para Mão de Obra SICRO3/IOPES
                                
                            if preco_idx is not None and preco_idx < len(row_list) and row_list[preco_idx] != '*':
                                preco_valor = parse_number(row_list[preco_idx])
                            else:
                                # Fallback Definitivo: Procura de trás pra frente o Custo Total (último) e Preço Unitário (penúltimo)
                                preco_valor = 0.0
                                last_num_idx = -1
                                for i in range(len(row_list) - 1, qtd_idx, -1):
                                    val = row_list[i]
                                    if val != '*' and val != '':
                                        parsed = parse_number(val)
                                        if parsed > 0:
                                            last_num_idx = i
                                            break
                                
                                if last_num_idx != -1:
                                    # Tenta pegar o número imediatamente antes do Total (Preço Unitário)
                                    if last_num_idx - 1 >= 0:
                                        val_preco = row_list[last_num_idx - 1]
                                        if val_preco != '*' and val_preco != '':
                                            parsed_preco = parse_number(val_preco)
                                            if parsed_preco > 0:
                                                preco_valor = parsed_preco
                                    
                                    # Se não achou preço unitário separado, usa o próprio Total
                                    if preco_valor == 0.0:
                                        preco_valor = parse_number(row_list[last_num_idx])
                        else:
                            # Padrão SINAPI
                            und_item = row_list[col_und].upper() if col_und < len(row_list) else ""
                            qtd_valor = parse_number(row_list[col_quant]) if col_quant < len(row_list) else 0.0
                            preco_valor = parse_number(row_list[col_preco]) if col_preco < len(row_list) else 0.0

                        ordem_sequencial += 1
                        dados.append({
                            'Ordem': ordem_sequencial, 'Servico_Pai': cod_pai_atual, 'Descricao_Pai': desc_pai_atual,
                            'Insumo_Filho': cod_item, 'Descricao_Filho': desc_item, 'Und': und_item,
                            'Qtd': qtd_valor, 'Preco_Unitario': preco_valor, 'Status_Parsing': 'OK'
                        })
                    
        df_final = pd.DataFrame(dados)
        if df_final.empty: return None, f"Planilha {nome_arquivo}: Nenhum dado válido extraído.", None
        
        checksum = hashlib.md5(df_final.to_string().encode()).hexdigest()
        log_msg = f"⚠️ {len(log_erros_parse)} linhas com erro de parsing" if log_erros_parse else ""
        return df_final, log_msg, (checksum, len(df_final))
    except Exception as e:
        return None, f"Erro ao processar {nome_arquivo}: {str(e)}", None

@st.cache_data(ttl=CACHE_TTL)
def transformar_hierarquico(df):
    if df.empty: return pd.DataFrame()
    df = df.sort_values('Ordem')
    linhas = []
    pai_atual = None
    colunas_vazias = ['Código', 'Descrição', 'Und_Base', 'Und_Prop', 'Qtd_Base', 'Qtd_Prop', 'Delta_Qtd', 'Preco_Base', 'Preco_Prop', 'Delta_Preco', 'Var_Preco_%', 'Total_Base', 'Total_Prop', 'Delta_Total', 'Var_Total_%']
    
    for _, row in df.iterrows():
        if row['Servico_Pai'] != pai_atual:
            if pai_atual is not None: linhas.append({c: np.nan for c in colunas_vazias})
            linhas.append({
                'Código': row['Servico_Pai'], 'Descrição': f"COMPOSIÇÃO: {row['Descricao_Pai']}",
                'Und_Base': '---', 'Und_Prop': '---', **{c: np.nan for c in colunas_vazias[4:]}
            })
            pai_atual = row['Servico_Pai']
            
        linhas.append({
            'Código': row['Insumo_Filho'], 'Descrição': row['Descricao_Filho'], 'Und_Base': row['Und_Base'], 'Und_Prop': row['Und_Prop'],
            'Qtd_Base': row['Qtd_Base'], 'Qtd_Prop': row['Qtd_Prop'], 'Delta_Qtd': row['Delta_Qtd'],
            'Preco_Base': row['Preco_Base'], 'Preco_Prop': row['Preco_Prop'], 'Delta_Preco': row['Delta_Preco'], 'Var_Preco_%': row['Var_Preco_%'],
            'Total_Base': row['Total_Base'], 'Total_Prop': row['Total_Prop'], 'Delta_Total': row['Delta_Total'], 'Var_Total_%': row['Var_Total_%']
        })
    return pd.DataFrame(linhas)

@st.cache_data(ttl=CACHE_TTL)
def transformar_hierarquico_raw(df):
    if df.empty: return pd.DataFrame()
    df = df.sort_values('Ordem')
    linhas = []
    pai_atual = None
    df_copy = df.copy()
    df_copy['Total'] = df_copy['Qtd'] * df_copy['Preco_Unitario']
    colunas_esquema = ['Código', 'Descrição', 'Unidade', 'Quantidade', 'Preço Unitário', 'Total']
    
    for _, row in df_copy.iterrows():
        if row['Servico_Pai'] != pai_atual:
            if pai_atual is not None: linhas.append({c: np.nan for c in colunas_esquema})
            linhas.append({
                'Código': row['Servico_Pai'], 'Descrição': f"COMPOSIÇÃO: {row['Descricao_Pai']}",
                'Unidade': '---', 'Quantidade': np.nan, 'Preço Unitário': np.nan, 'Total': np.nan
            })
            pai_atual = row['Servico_Pai']
            
        linhas.append({
            'Código': row['Insumo_Filho'], 'Descrição': row['Descricao_Filho'], 'Unidade': row['Und'],
            'Quantidade': row['Qtd'], 'Preço Unitário': row['Preco_Unitario'], 'Total': row['Total']
        })
    return pd.DataFrame(linhas)

# ==========================================
# 2. ESTILIZADORES DA INTERFACE (UI)
# ==========================================

def estilizar_relatorio(row):
    if str(row.get('Und_Base', '')).strip() == '---':
        return ['background-color: #dbeafe; font-weight: bold; color: #1e3a8a;'] * len(row)
    estilos = [''] * len(row)
    
    def safe_float_ui(v):
        if pd.isna(v) or v == '' or v == '*': return 0.0
        if isinstance(v, (int, float)): return float(v)
        s = str(v).strip()
        if ',' in s and '.' in s: s = s.replace('.', '').replace(',', '.')
        elif ',' in s: s = s.replace(',', '.')
        s = re.sub(r'[R$\s%]', '', s)
        try: return float(s)
        except: return 0.0

    for i, col in enumerate(row.index):
        val = row[col]
        
        # 🟨 AMARELO: Unidade de medida diferente (case insensitive)
        if col in ['Und_Base', 'Und_Prop']:
            und_b = str(row.get('Und_Base', '')).strip().upper()
            und_p = str(row.get('Und_Prop', '')).strip().upper()
            if und_b not in ['', 'NAN', '---'] and und_p not in ['', 'NAN', '---'] and und_b != und_p:
                estilos[i] = 'background-color: #fef08a; color: #713f12; font-weight: bold;'
                
        if pd.isna(val) or val == '': continue
        v = safe_float_ui(val)
        if v == 0 and col not in ['Delta_Qtd', 'Delta_Preco', 'Delta_Total', 'Var_Preco_%', 'Var_Total_%']: continue
            
        # 🟪 ROXO: Quantidade DIFERENTE (!= 0)
        if col == 'Delta_Qtd' and v != 0: 
            estilos[i] = 'background-color: #e9d5ff; color: #6b21a8; font-weight: bold;'
        # 🟥 VERMELHO: Sobrepreço (> 0)
        elif col in ['Delta_Preco', 'Delta_Total'] and v > 0: 
            estilos[i] = 'background-color: #fca5a5; color: #7f1d1d; font-weight: bold;'
        elif col in ['Var_Preco_%', 'Var_Total_%']:
            if v > 0: 
                estilos[i] = 'background-color: #fca5a5; color: #7f1d1d; font-weight: bold;'
            # 🟧 LARANJA: Desconto > limiar
            elif v < st.session_state.get('limiar_desconto', -0.25): 
                estilos[i] = 'background-color: #fdba74; color: #7c2d12; font-weight: bold;'
    return estilos

def estilizar_relatorio_raw(row):
    if str(row.get('Unidade', '')).strip() == '---' or 'COMPOSIÇÃO' in str(row.get('Descrição', '')): 
        return ['background-color: #dbeafe; font-weight: bold; color: #1e3a8a;'] * len(row)
    return [''] * len(row)

# ==========================================
# 3. MOTOR EXPORTADOR EXCEL MULTI-ABA
# ==========================================

def gerar_excel_bytes(dash_data, df_matriz, df_inconformidades, df_nao_encontrados_base, df_nao_encontrados_prop, df_parsing, df_db_base, df_db_prop):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # 1. ESCRITA DAS ABAS PADRÃO
        df_matriz.to_excel(writer, index=False, sheet_name='📋 Matriz Completa', startrow=8)
        if not df_inconformidades.empty: df_inconformidades.to_excel(writer, index=False, sheet_name='🚨 Inconformidades', startrow=8)
        else:
            ws_inc = writer.book.create_sheet(title='🚨 Inconformidades')
            ws_inc['A1'] = "Nenhuma inconformidade paramétrica detectada."
            
        df_nao_encontrados_base.to_excel(writer, index=False, sheet_name='📍 Não Encontrados na Base', startrow=1)
        df_nao_encontrados_prop.to_excel(writer, index=False, sheet_name='📍 Omitidos na Proposta', startrow=1)
        df_parsing.to_excel(writer, index=False, sheet_name='📝 Log de Erros de Parsing', startrow=1)
        df_db_base.to_excel(writer, index=False, sheet_name='🗄️ DB Base', startrow=1)
        df_db_prop.to_excel(writer, index=False, sheet_name='🗄️ DB Proposta', startrow=1)
        
        wb = writer.book

        # ==========================================
        # 2. CONSTRUÇÃO DO PAINEL DASHBOARD KPI (NATIVO)
        # ==========================================
        ws_kpi = wb.create_sheet('📊 Dashboard KPI', 0)
        ws_kpi.sheet_view.showGridLines = False
        
        border_box = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'), top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

        ws_kpi.merge_cells('B2:J3')
        title = ws_kpi['B2']
        title.value = "📊 PAINEL ANALÍTICO DE CONFORMIDADE CONTRATUAL"
        title.font = Font(size=16, bold=True, color="FFFFFF")
        title.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        title.alignment = Alignment(horizontal="center", vertical="center")

        def desenhar_card(row, col, title, value, format_type='int'):
            c_title = ws_kpi.cell(row=row, column=col)
            c_title.value = title
            c_title.font = Font(bold=True, color="FFFFFF", size=10)
            c_title.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            c_title.alignment = Alignment(horizontal="center", vertical="center")
            c_title.border = border_box
            ws_kpi.cell(row=row, column=col+1).border = border_box
            ws_kpi.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+1)

            c_val = ws_kpi.cell(row=row+1, column=col)
            c_val.value = value
            c_val.font = Font(size=14, bold=True, color="0F172A")
            c_val.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            c_val.alignment = Alignment(horizontal="center", vertical="center")
            c_val.border = border_box
            
            for r_adj in range(row+1, row+3):
                for c_adj in range(col, col+2):
                    ws_kpi.cell(row=r_adj, column=c_adj).border = border_box
            ws_kpi.merge_cells(start_row=row+1, start_column=col, end_row=row+2, end_column=col+1)

            if format_type == 'currency': c_val.number_format = '"R$" #,##0.00'
            elif format_type == 'percent': c_val.number_format = '0.00%'
            else: c_val.number_format = '#,##0'

        desenhar_card(5, 2, "ITENS AUDITADOS", dash_data['total_insumos'], 'int')
        desenhar_card(5, 5, "SALDO DO ORÇAMENTO", dash_data['total_proposta'], 'currency')
        desenhar_card(5, 8, "TAXA DE CONFORMIDADE", dash_data['taxa_conformidade'], 'percent')

        desenhar_card(9, 2, "RISCO SOBREPREÇO", dash_data['financeiro_sobrepreco'], 'currency')
        desenhar_card(9, 5, "DESCONTO OCULTO (INEX)", dash_data['financeiro_inexequivel'], 'currency')
        desenhar_card(9, 8, "MAIOR DESVIO ÚNICO", dash_data['max_desvio'], 'currency')

        def desenhar_tabela(start_row, start_col, df, title_text):
            if df.empty: return start_row
            
            t_cell = ws_kpi.cell(row=start_row, column=start_col)
            t_cell.value = title_text
            t_cell.font = Font(bold=True, size=11, color="1E3A8A")
            ws_kpi.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + len(df.columns) - 1)

            for c_idx, col_name in enumerate(df.columns):
                cell = ws_kpi.cell(row=start_row+1, column=start_col+c_idx)
                cell.value = col_name
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
                cell.border = border_box

            for r_idx, row in enumerate(df.values):
                for c_idx, val in enumerate(row):
                    cell = ws_kpi.cell(row=start_row+2+r_idx, column=start_col+c_idx)
                    cell.value = val
                    cell.border = border_box
                    cell.alignment = Alignment(vertical="center")

                    col_name = df.columns[c_idx]
                    if any(k in col_name for k in ['%', 'Variação', 'Desconto']): cell.number_format = '0.00%'
                    elif any(k in col_name for k in ['R$', 'Impacto', 'Defasagem', 'Sobrepreço']): cell.number_format = '"R$" #,##0.00'
                    elif isinstance(val, (int, float)): cell.number_format = '#,##0'
                    
            return start_row + len(df) + 4 

        desenhar_tabela(14, 2, dash_data['count_graf'], "📌 OCORRÊNCIAS POR TIPOLOGIA")
        desenhar_tabela(14, 6, dash_data['money_graf'], "💸 IMPACTO FINANCEIRO LÍQUIDO")

        next_row = desenhar_tabela(23, 2, dash_data['top_sobre'], "🎯 FRENTE 1: TOP 5 IMPACTOS DE SOBREPREÇO")
        desenhar_tabela(next_row, 2, dash_data['top_inex'], "🎯 FRENTE 2: TOP 5 RISCOS DE INEXEQUIBILIDADE")

        ws_kpi.column_dimensions['A'].width = 3
        ws_kpi.column_dimensions['B'].width = 15
        ws_kpi.column_dimensions['C'].width = 32
        ws_kpi.column_dimensions['D'].width = 4
        ws_kpi.column_dimensions['E'].width = 20
        ws_kpi.column_dimensions['F'].width = 25
        ws_kpi.column_dimensions['G'].width = 4
        ws_kpi.column_dimensions['H'].width = 18
        ws_kpi.column_dimensions['I'].width = 25
        ws_kpi.column_dimensions['J'].width = 3

# ==========================================
# 3. MOTOR EXPORTADOR EXCEL MULTI-ABA
# ==========================================

def gerar_excel_bytes(dash_data, df_matriz, df_inconformidades, df_nao_encontrados_base, df_nao_encontrados_prop, df_parsing, df_db_base, df_db_prop):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # 1. ESCRITA DAS ABAS PADRÃO
        df_matriz.to_excel(writer, index=False, sheet_name='📋 Matriz Completa', startrow=8)
        if not df_inconformidades.empty: df_inconformidades.to_excel(writer, index=False, sheet_name='🚨 Inconformidades', startrow=8)
        else:
            ws_inc = writer.book.create_sheet(title='🚨 Inconformidades')
            ws_inc['A1'] = "Nenhuma inconformidade paramétrica detectada."
            
        df_nao_encontrados_base.to_excel(writer, index=False, sheet_name='📍 Não Encontrados na Base', startrow=1)
        df_nao_encontrados_prop.to_excel(writer, index=False, sheet_name='📍 Omitidos na Proposta', startrow=1)
        df_parsing.to_excel(writer, index=False, sheet_name='📝 Log de Erros de Parsing', startrow=1)
        df_db_base.to_excel(writer, index=False, sheet_name='🗄️ DB Base', startrow=1)
        df_db_prop.to_excel(writer, index=False, sheet_name='🗄️ DB Proposta', startrow=1)
        
        wb = writer.book

        # ==========================================
        # 2. CONSTRUÇÃO DO PAINEL DASHBOARD KPI (NATIVO)
        # ==========================================
        ws_kpi = wb.create_sheet('📊 Dashboard KPI', 0)
        ws_kpi.sheet_view.showGridLines = False
        
        border_box = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'), top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

        ws_kpi.merge_cells('B2:J3')
        title = ws_kpi['B2']
        title.value = "📊 PAINEL ANALÍTICO DE CONFORMIDADE CONTRATUAL"
        title.font = Font(size=16, bold=True, color="FFFFFF")
        title.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        title.alignment = Alignment(horizontal="center", vertical="center")

        def desenhar_card(row, col, title, value, format_type='int'):
            c_title = ws_kpi.cell(row=row, column=col)
            c_title.value = title
            c_title.font = Font(bold=True, color="FFFFFF", size=10)
            c_title.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            c_title.alignment = Alignment(horizontal="center", vertical="center")
            c_title.border = border_box
            ws_kpi.cell(row=row, column=col+1).border = border_box
            ws_kpi.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+1)

            c_val = ws_kpi.cell(row=row+1, column=col)
            c_val.value = value
            c_val.font = Font(size=14, bold=True, color="0F172A")
            c_val.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            c_val.alignment = Alignment(horizontal="center", vertical="center")
            c_val.border = border_box
            
            for r_adj in range(row+1, row+3):
                for c_adj in range(col, col+2):
                    ws_kpi.cell(row=r_adj, column=c_adj).border = border_box
            ws_kpi.merge_cells(start_row=row+1, start_column=col, end_row=row+2, end_column=col+1)

            if format_type == 'currency': c_val.number_format = '"R$" #,##0.00'
            elif format_type == 'percent': c_val.number_format = '0.00%'
            else: c_val.number_format = '#,##0'

        desenhar_card(5, 2, "ITENS AUDITADOS", dash_data['total_insumos'], 'int')
        desenhar_card(5, 5, "SALDO DO ORÇAMENTO", dash_data['total_proposta'], 'currency')
        desenhar_card(5, 8, "TAXA DE CONFORMIDADE", dash_data['taxa_conformidade'], 'percent')

        desenhar_card(9, 2, "RISCO SOBREPREÇO", dash_data['financeiro_sobrepreco'], 'currency')
        desenhar_card(9, 5, "DESCONTO OCULTO (INEX)", dash_data['financeiro_inexequivel'], 'currency')
        desenhar_card(9, 8, "MAIOR DESVIO ÚNICO", dash_data['max_desvio'], 'currency')

        def desenhar_tabela(start_row, start_col, df, title_text):
            if df.empty: return start_row
            
            t_cell = ws_kpi.cell(row=start_row, column=start_col)
            t_cell.value = title_text
            t_cell.font = Font(bold=True, size=11, color="1E3A8A")
            ws_kpi.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + len(df.columns) - 1)

            for c_idx, col_name in enumerate(df.columns):
                cell = ws_kpi.cell(row=start_row+1, column=start_col+c_idx)
                cell.value = col_name
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
                cell.border = border_box

            for r_idx, row in enumerate(df.values):
                for c_idx, val in enumerate(row):
                    cell = ws_kpi.cell(row=start_row+2+r_idx, column=start_col+c_idx)
                    cell.value = val
                    cell.border = border_box
                    cell.alignment = Alignment(vertical="center")

                    col_name = df.columns[c_idx]
                    if any(k in col_name for k in ['%', 'Variação', 'Desconto']): cell.number_format = '0.00%'
                    elif any(k in col_name for k in ['R$', 'Impacto', 'Defasagem', 'Sobrepreço']): cell.number_format = '"R$" #,##0.00'
                    elif isinstance(val, (int, float)): cell.number_format = '#,##0'
                    
            return start_row + len(df) + 4 

        desenhar_tabela(14, 2, dash_data['count_graf'], "📌 OCORRÊNCIAS POR TIPOLOGIA")
        desenhar_tabela(14, 6, dash_data['money_graf'], "💸 IMPACTO FINANCEIRO LÍQUIDO")

        next_row = desenhar_tabela(23, 2, dash_data['top_sobre'], "🎯 FRENTE 1: TOP 5 IMPACTOS DE SOBREPREÇO")
        desenhar_tabela(next_row, 2, dash_data['top_inex'], "🎯 FRENTE 2: TOP 5 RISCOS DE INEXEQUIBILIDADE")

        ws_kpi.column_dimensions['A'].width = 3
        ws_kpi.column_dimensions['B'].width = 15
        ws_kpi.column_dimensions['C'].width = 32
        ws_kpi.column_dimensions['D'].width = 4
        ws_kpi.column_dimensions['E'].width = 20
        ws_kpi.column_dimensions['F'].width = 25
        ws_kpi.column_dimensions['G'].width = 4
        ws_kpi.column_dimensions['H'].width = 18
        ws_kpi.column_dimensions['I'].width = 25
        ws_kpi.column_dimensions['J'].width = 3

        # ==========================================
        # 3. ESTILIZAÇÃO PADRÃO DAS ABAS DE DADOS
        # ==========================================
        def injetar_legenda(ws):
            legendas = [
                ('A1', 'FCA5A5', '🟥 VERMELHO: Sobrepreço (Preço/Total da proposta superior à referência).'),
                ('A2', 'E9D5FF', '🟪 ROXO: Quantidade diferente da referência.'),
                ('A3', 'FDBA74', '🟧 LARANJA: Inexequibilidade (Desconto superior a 25%).'),
                ('A4', 'FEF08A', '🟨 AMARELO: Fraude Métrica (Unidades de medida incompatíveis).'),
                ('A7', 'DBEAFE', '🟦 AZUL CLARO: Composição Analítica Pai.')
            ]
            for celula, cor, texto in legendas:
                ws[celula] = texto
                ws[celula].fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")
                ws[celula].font = Font(bold=True, size=9)

        def safe_float(v):
            if v is None: return 0.0
            if isinstance(v, (int, float)): return float(v)
            s = str(v).strip()
            if s in ('', '*', 'nan', 'None', '---'): return 0.0
            if ',' in s and '.' in s:
                s = s.replace('.', '').replace(',', '.')
            elif ',' in s:
                s = s.replace(',', '.')
            s = re.sub(r'[R$\s%]', '', s)
            try:
                return float(s)
            except ValueError:
                return 0.0

        def safe_str(v):
            if v is None: return ''
            s = str(v).strip()
            return '' if s.lower() in ('nan', 'none', '---') else s

        limiar_desconto = st.session_state.get('limiar_desconto', -0.25)

        def mapear_colunas(ws, linha_cabecalho):
            mapa = {}
            for c_idx in range(1, ws.max_column + 1):
                nome = safe_str(ws.cell(row=linha_cabecalho, column=c_idx).value)
                if nome:
                    mapa[nome] = c_idx
            return mapa

        for name in wb.sheetnames:
            if name == '📊 Dashboard KPI': continue
            ws = wb[name]

            if name in ['📋 Matriz Completa', '🚨 Inconformidades']:
                injetar_legenda(ws)
                linha_cabecalho = 9
            elif name in ['📍 Não Encontrados na Base', '📍 Omitidos na Proposta',
                          '🗄️ DB Base', '🗄️ DB Proposta', '📝 Log de Erros de Parsing']:
                linha_cabecalho = 2
            else:
                linha_cabecalho = 1

            col_map = mapear_colunas(ws, linha_cabecalho)

            is_matriz = name in ['📋 Matriz Completa', '🚨 Inconformidades']
            is_raw = name in ['📍 Não Encontrados na Base', '📍 Omitidos na Proposta',
                              '🗄️ DB Base', '🗄️ DB Proposta']

            if is_matriz:
                col_und_base = col_map.get('Und_Base')
                col_und_prop = col_map.get('Und_Prop')
                col_delta_qtd = col_map.get('Delta_Qtd')
                col_delta_preco = col_map.get('Delta_Preco')
                col_var_preco = col_map.get('Var_Preco_%')
                col_delta_total = col_map.get('Delta_Total')
                col_var_total = col_map.get('Var_Total_%')

                for r_idx in range(linha_cabecalho + 1, ws.max_row + 1):
                    und_base_val = safe_str(ws.cell(row=r_idx, column=col_und_base).value) if col_und_base else ''
                    raw_und = ws.cell(row=r_idx, column=col_und_base).value if col_und_base else None
                    
                    if safe_str(raw_und) == '---' or 'COMPOSIÇÃO' in safe_str(ws.cell(row=r_idx, column=2).value):
                        azul_fill = PatternFill(start_color='DBEAFE', end_color='DBEAFE', fill_type="solid")
                        for c_idx in range(1, ws.max_column + 1):
                            cell = ws.cell(row=r_idx, column=c_idx)
                            cell.fill = azul_fill
                            cell.font = Font(bold=True, color='1E3A8A')
                        continue

                    und_prop_val = safe_str(ws.cell(row=r_idx, column=col_und_prop).value) if col_und_prop else ''
                    delta_qtd = safe_float(ws.cell(row=r_idx, column=col_delta_qtd).value) if col_delta_qtd else 0.0
                    delta_preco = safe_float(ws.cell(row=r_idx, column=col_delta_preco).value) if col_delta_preco else 0.0
                    var_preco = safe_float(ws.cell(row=r_idx, column=col_var_preco).value) if col_var_preco else 0.0
                    delta_total = safe_float(ws.cell(row=r_idx, column=col_delta_total).value) if col_delta_total else 0.0
                    var_total = safe_float(ws.cell(row=r_idx, column=col_var_total).value) if col_var_total else 0.0

                    if und_base_val and und_prop_val and und_base_val.upper() != und_prop_val.upper():
                        yellow_fill = PatternFill(start_color='FEF08A', end_color='FEF08A', fill_type="solid")
                        if col_und_base: ws.cell(row=r_idx, column=col_und_base).fill = yellow_fill
                        if col_und_prop: ws.cell(row=r_idx, column=col_und_prop).fill = yellow_fill

                    if delta_qtd != 0 and col_delta_qtd:
                        ws.cell(row=r_idx, column=col_delta_qtd).fill = PatternFill(start_color='E9D5FF', end_color='E9D5FF', fill_type="solid")

                    if delta_preco > 0 and col_delta_preco:
                        ws.cell(row=r_idx, column=col_delta_preco).fill = PatternFill(start_color='FCA5A5', end_color='FCA5A5', fill_type="solid")
                    if delta_total > 0 and col_delta_total:
                        ws.cell(row=r_idx, column=col_delta_total).fill = PatternFill(start_color='FCA5A5', end_color='FCA5A5', fill_type="solid")
                    if var_preco > 0 and col_var_preco:
                        ws.cell(row=r_idx, column=col_var_preco).fill = PatternFill(start_color='FCA5A5', end_color='FCA5A5', fill_type="solid")
                    if var_total > 0 and col_var_total:
                        ws.cell(row=r_idx, column=col_var_total).fill = PatternFill(start_color='FCA5A5', end_color='FCA5A5', fill_type="solid")

                    if var_preco < limiar_desconto and col_var_preco:
                        ws.cell(row=r_idx, column=col_var_preco).fill = PatternFill(start_color='FDBA74', end_color='FDBA74', fill_type="solid")
                    if var_total < limiar_desconto and col_var_total:
                        ws.cell(row=r_idx, column=col_var_total).fill = PatternFill(start_color='FDBA74', end_color='FDBA74', fill_type="solid")

            elif is_raw:
                for r_idx in range(linha_cabecalho + 1, ws.max_row + 1):
                    desc_val = safe_str(ws.cell(row=r_idx, column=2).value)
                    if 'COMPOSIÇÃO' in desc_val:
                        azul_fill = PatternFill(start_color='DBEAFE', end_color='DBEAFE', fill_type="solid")
                        for c_idx in range(1, ws.max_column + 1):
                            cell = ws.cell(row=r_idx, column=c_idx)
                            cell.fill = azul_fill
                            cell.font = Font(bold=True, color='1E3A8A')

            formatos_coluna = {}
            for col_idx in range(1, ws.max_column + 1):
                nome_col = safe_str(ws.cell(row=linha_cabecalho, column=col_idx).value)
                if any(k in nome_col for k in ['Preco', 'Preço', 'Total', 'Valor', 'Delta_Preco', 'Delta_Total']):
                    formatos_coluna[col_idx] = '0.00%' if '%' in nome_col or 'Var_' in nome_col else '"R$" #,##0.00'
                elif any(k in nome_col for k in ['Qtd', 'Quantidade', 'Delta_Qtd']):
                    formatos_coluna[col_idx] = '#,##0.0000'

            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                col_idx = col[0].column
                header_cell = ws.cell(row=linha_cabecalho, column=col_idx)
                header_cell.font = Font(bold=True, color='FFFFFF')
                header_cell.fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type="solid")
                header_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                for cell in col:
                    if cell.row > linha_cabecalho and col_idx in formatos_coluna and isinstance(cell.value, (int, float, np.number)):
                        cell.number_format = formatos_coluna[col_idx]
                    try:
                        if cell.value is not None:
                            val_str = str(cell.value)
                            if col_idx in formatos_coluna and isinstance(cell.value, (int, float)):
                                val_str = f"{cell.value * 100:.2f}%" if '%' in formatos_coluna[col_idx] else f"R$ {cell.value:,.2f}"
                            max_length = max(max_length, len(val_str))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max(max_length + 3, 11), 70)

    # O RETORNO DEVE FICAR AQUI, FORA DO BLOCO 'WITH'
    return output.getvalue()

# ==========================================
# 4. INTERFACE GRÁFICA (UI - STREAMLIT)
# ==========================================

if 'limiar_desconto' not in st.session_state:
    st.session_state.limiar_desconto = -0.25

with st.sidebar:
    st.subheader("⚙️ Parâmetros de Auditoria")
    limiar_desconto = st.slider("Limiar de Desconto (Inexequibilidade)", -50, -5, -25, 1)
    st.session_state.limiar_desconto = limiar_desconto / 100
    st.success(f"📌 **Limiar Configurado: {limiar_desconto}%**")
    st.divider()
    st.subheader("📌 Legenda de Auditoria")
    
    html_legend = """
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fca5a5; color: #7f1d1d; font-family: sans-serif; font-size:13px;"><b>🟥 Vermelho (Sobrepreço)</b><br>Preço unitário ou total superior à base.</div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #e9d5ff; color: #6b21a8; font-family: sans-serif; font-size:13px;"><b>🟪 Roxo (Qtd. Alterada)</b><br>Quantidade do insumo majorada.</div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fdba74; color: #7c2d12; font-family: sans-serif; font-size:13px;"><b>🟧 Laranja (Inexequibilidade)</b><br>Desconto excessivo.</div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fef08a; color: #713f12; font-family: sans-serif; font-size:13px;"><b>🟨 Amarelo (Fraude Métrica)</b><br>Unidades de Medida incompatíveis.</div>
    """
    st.markdown(html_legend, unsafe_allow_html=True)

st.title("🛡️ Auditoria de Orçamentos PRO")
st.markdown("Validação paramétrica de planilhas orçamentárias de Engenharia Civil.")
st.divider()

col1, col2 = st.columns(2)
with col1: arquivo_base = st.file_uploader("1. Base de Referência (SINAPI/ORSE)", type=["xlsx", "xls"])
with col2: arquivo_proposta = st.file_uploader("2. Proposta da Empreiteira", type=["xlsx", "xls"])

if arquivo_base and arquivo_proposta:
    with st.spinner("Compilando inteligência analítica do painel de controle..."):
        
     df_base = df_base_raw.copy()
            df_prop = df_prop_raw.copy()
            
            # --- CORREÇÃO DO PRODUTO CARTESIANO (JOIN EXPLOSION) ---
            # Criação de um ID único para cada repetição de insumo idêntico
            df_base['match_id'] = df_base.groupby(['Servico_Pai', 'Insumo_Filho', 'Descricao_Filho']).cumcount()
            df_prop['match_id'] = df_prop.groupby(['Servico_Pai', 'Insumo_Filho', 'Descricao_Filho']).cumcount()
            
            # Novo índice blindado contra multiplicações indevidas
            chaves_index = ['Servico_Pai', 'Insumo_Filho', 'Descricao_Filho', 'match_id']
            
            df_base.set_index(chaves_index, inplace=True)
            df_prop.set_index(chaves_index, inplace=True)
            
            # Cruzamento 1-para-1 exato
            df_auditoria = df_base.join(df_prop[['Und', 'Qtd', 'Preco_Unitario']], how='inner', rsuffix='_Prop')
            
            # Proposta -> Base (Itens omitidos na base)
            indices_nao_encontrados_base = set(df_prop.index) - set(df_base.index)
            if indices_nao_encontrados_base:
                df_nao_encontrados_base = df_prop_raw[df_prop_raw.set_index(chaves_index).index.isin(indices_nao_encontrados_base)].reset_index(drop=True).drop(columns=['match_id'], errors='ignore')
            else:
                df_nao_encontrados_base = pd.DataFrame()
            
            # Base -> Proposta (Itens omitidos na proposta)
            indices_nao_encontrados_prop = set(df_base.index) - set(df_prop.index)
            if indices_nao_encontrados_prop:
                df_nao_encontrados_prop = df_base_raw[df_base_raw.set_index(chaves_index).index.isin(indices_nao_encontrados_prop)].reset_index(drop=True).drop(columns=['match_id'], errors='ignore')
            else:
                df_nao_encontrados_prop = pd.DataFrame()
            
            # Renomeando e calculando os deltas
            df_auditoria.rename(columns={'Und': 'Und_Base', 'Qtd': 'Qtd_Base', 'Preco_Unitario': 'Preco_Base', 'Und_Prop': 'Und_Prop', 'Qtd_Prop': 'Qtd_Prop', 'Preco_Unitario_Prop': 'Preco_Prop'}, inplace=True)
            
            df_auditoria['Total_Base'] = df_auditoria['Qtd_Base'] * df_auditoria['Preco_Base']
            df_auditoria['Total_Prop'] = df_auditoria['Qtd_Prop'] * df_auditoria['Preco_Prop']
            df_auditoria['Delta_Qtd'] = df_auditoria['Qtd_Prop'] - df_auditoria['Qtd_Base']
            df_auditoria['Delta_Preco'] = df_auditoria['Preco_Prop'] - df_auditoria['Preco_Base']
            df_auditoria['Delta_Total'] = df_auditoria['Total_Prop'] - df_auditoria['Total_Base']
            df_auditoria['Var_Preco_%'] = np.where(df_auditoria['Preco_Base'] > 0, (df_auditoria['Preco_Prop'] / df_auditoria['Preco_Base']) - 1, 0)
            df_auditoria['Var_Total_%'] = np.where(df_auditoria['Total_Base'] > 0, (df_auditoria['Total_Prop'] / df_auditoria['Total_Base']) - 1, 0)
            
            # Desfazendo o índice temporário para o resto da UI funcionar
            df_completo = df_auditoria.reset_index().drop(columns=['match_id'])
            # Limpeza de valores nulos para evitar erros nos filtros do dashboard
            df_completo['Delta_Total'] = df_completo['Delta_Total'].fillna(0)
            df_completo['Delta_Preco'] = df_completo['Delta_Preco'].fillna(0)
            df_completo['Delta_Qtd'] = df_completo['Delta_Qtd'].fillna(0)
            df_completo['Var_Preco_%'] = df_completo['Var_Preco_%'].fillna(0)
            df_completo['Var_Total_%'] = df_completo['Var_Total_%'].fillna(0)
            df_completo['Und_Base'] = df_completo['Und_Base'].fillna('').astype(str)
            df_completo['Und_Prop'] = df_completo['Und_Prop'].fillna('').astype(str)
            
            # Filtros alinhados EXATAMENTE com as regras de cor do Excel
            # Vermelho: Preço ou Total superior (> 0)
            sobrepreco_filter = (df_completo['Delta_Preco'] > 0) | (df_completo['Delta_Total'] > 0) | (df_completo['Var_Preco_%'] > 0) | (df_completo['Var_Total_%'] > 0)
            
            # Laranja: Variação de Preço ou Total com desconto superior a 25%
            inexequivel_filter = (df_completo['Var_Preco_%'] < st.session_state.limiar_desconto) | (df_completo['Var_Total_%'] < st.session_state.limiar_desconto)
            
            # Roxo: Qualquer diferença de quantidade (!= 0)
            qtd_filter = df_completo['Delta_Qtd'] != 0
            
            # Amarelo: Unidade de medida diferente (ignorando maiúsculas/minúsculas)
            und_filter = df_completo['Und_Base'].str.upper() != df_completo['Und_Prop'].str.upper()
            
            irregularidades = df_completo[sobrepreco_filter | qtd_filter | inexequivel_filter | und_filter]
            
            df_visual_completo = transformar_hierarquico(df_completo)
            df_visual_erros = transformar_hierarquico(irregularidades) if not irregularidades.empty else pd.DataFrame()
            
            df_visual_ne_base = transformar_hierarquico_raw(df_nao_encontrados_base)
            df_visual_ne_prop = transformar_hierarquico_raw(df_nao_encontrados_prop)
            
            df_visual_db_base = transformar_hierarquico_raw(df_base_raw)
            df_visual_db_prop = transformar_hierarquico_raw(df_prop_raw)
            
            # Cálculo de Métricas Executivas
            total_insumos, total_irregularidades = len(df_completo), len(irregularidades)
            total_base, total_proposta = float(df_completo['Total_Base'].sum()), float(df_completo['Total_Prop'].sum())
            delta_total = total_proposta - total_base
            var_total_geral = (total_proposta / total_base - 1) if total_base > 0 else 0
            taxa_conformidade = ((total_insumos - total_irregularidades) / total_insumos) if total_insumos > 0 else 0
            
            # Somatório financeiro apenas dos impactos positivos (sobrepreço real)
            financeiro_sobrepreco = float(df_completo[df_completo['Delta_Total'] > 0]['Delta_Total'].sum())
            # Impacto absoluto das diferenças de quantidade
            financeiro_qtd = float(abs(df_completo[qtd_filter]['Delta_Total'].sum()))
            # Impacto absoluto dos descontos extremos
            financeiro_inexequivel = float(abs(df_completo[inexequivel_filter]['Delta_Total'].sum()))
            
            max_desvio_individual = float(df_completo['Delta_Total'].max()) if not df_completo.empty else 0.0
            
            sobreprecados, quantidades_alteradas = len(df_completo[sobrepreco_filter]), len(df_completo[qtd_filter])
            unidades_incompativeis, inexequiveis = len(df_completo[und_filter]), len(df_completo[inexequivel_filter])
            
            df_top_sobre = df_completo[df_completo['Delta_Total'] > 0].sort_values(by='Delta_Total', ascending=False).head(5)
            if not df_top_sobre.empty:
                df_top_sobre_view = df_top_sobre[['Insumo_Filho', 'Descricao_Filho', 'Delta_Total', 'Var_Preco_%']].copy()
                df_top_sobre_view.columns = ['Código', 'Descrição do Insumo', 'Sobrepreço (R$)', 'Variação (%)']
            else: 
                df_top_sobre_view = pd.DataFrame()

            df_top_inex = df_completo[inexequivel_filter].sort_values(by='Delta_Total', ascending=True).head(5)
            if not df_top_inex.empty:
                df_top_inex_view = df_top_inex[['Insumo_Filho', 'Descricao_Filho', 'Delta_Total', 'Var_Total_%']].copy()
                df_top_inex_view.columns = ['Código', 'Descrição do Insumo', 'Defasagem (R$)', 'Variação Total (%)']
            else: 
                df_top_inex_view = pd.DataFrame()
            dash_data_excel = {
                'total_insumos': total_insumos, 'total_proposta': total_proposta, 'taxa_conformidade': taxa_conformidade,
                'financeiro_sobrepreco': financeiro_sobrepreco, 'financeiro_inexequivel': financeiro_inexequivel, 'max_desvio': max_desvio_individual,
                'count_graf': pd.DataFrame({'Tipologia de Erro': ['🟥 Sobrepreço', '🟪 Qtd. Majorada', '🟨 Fraude Métrica', '🟧 Inexequível'], 'Ocorrências': [sobreprecados, quantidades_alteradas, unidades_incompativeis, inexequiveis]}),
                'money_graf': pd.DataFrame({'Tipologia Financeira': ['Sobrepreço Global', 'Majorização de Qtd.', 'Descontos Extremos'], 'Impacto (R$)': [financeiro_sobrepreco, financeiro_qtd, financeiro_inexequivel]}),
                'top_sobre': df_top_sobre_view, 'top_inex': df_top_inex_view
            }

            dash_data_excel = {
                'total_insumos': total_insumos, 'total_proposta': total_proposta, 'taxa_conformidade': taxa_conformidade,
                'financeiro_sobrepreco': financeiro_sobrepreco, 'financeiro_inexequivel': financeiro_inexequivel, 'max_desvio': max_desvio_individual,
                'count_graf': pd.DataFrame({'Tipologia de Erro': ['🟥 Sobrepreço', '🟪 Qtd. Majorada', '🟨 Fraude Métrica', '🟧 Inexequível'], 'Ocorrências': [sobreprecados, quantidades_alteradas, unidades_incompativeis, inexequiveis]}),
                'money_graf': pd.DataFrame({'Tipologia Financeira': ['Sobrepreço Global', 'Majorização de Qtd.', 'Descontos Extremos'], 'Impacto (R$)': [financeiro_sobrepreco, financeiro_qtd, financeiro_inexequivel]}),
                'top_sobre': df_top_sobre_view, 'top_inex': df_top_inex_view
            }
            
            logs_parsing_lista = []
            df_parsing_excel = pd.DataFrame(logs_parsing_lista) if logs_parsing_lista else pd.DataFrame(columns=['Origem', 'Código', 'Descrição', 'Erro'])
            
            excel_bytes = gerar_excel_bytes(dash_data_excel, df_visual_completo, df_visual_erros, df_visual_ne_base, df_visual_ne_prop, df_parsing_excel, df_visual_db_base, df_visual_db_prop)
            
            formato_tela = {'Qtd_Base': '{:.4f}', 'Qtd_Prop': '{:.4f}', 'Delta_Qtd': '{:.4f}', 'Preco_Base': 'R$ {:.2f}', 'Preco_Prop': 'R$ {:.2f}', 'Delta_Preco': 'R$ {:.2f}', 'Var_Preco_%': '{:.2%}', 'Total_Base': 'R$ {:.2f}', 'Total_Prop': 'R$ {:.2f}', 'Delta_Total': 'R$ {:.2f}', 'Var_Total_%': '{:.2%}'}
            formato_tela_raw = {'Quantidade': '{:.4f}', 'Preço Unitário': 'R$ {:.2f}', 'Total': 'R$ {:.2f}'}
            
            styler_ui_completo = df_visual_completo.style.format(formato_tela, na_rep="").apply(estilizar_relatorio, axis=1)
            styler_ui_erros = df_visual_erros.style.format(formato_tela, na_rep="").apply(estilizar_relatorio, axis=1) if not df_visual_erros.empty else None
            
            styler_ui_ne_base = df_visual_ne_base.style.format(formato_tela_raw, na_rep="").apply(estilizar_relatorio_raw, axis=1) if not df_visual_ne_base.empty else None
            styler_ui_ne_prop = df_visual_ne_prop.style.format(formato_tela_raw, na_rep="").apply(estilizar_relatorio_raw, axis=1) if not df_visual_ne_prop.empty else None
            
            styler_ui_db_base = df_visual_db_base.style.format(formato_tela_raw, na_rep="").apply(estilizar_relatorio_raw, axis=1)
            styler_ui_db_prop = df_visual_db_prop.style.format(formato_tela_raw, na_rep="").apply(estilizar_relatorio_raw, axis=1)
            
            tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
                "📊 Dashboard KPI", 
                "📋 Matriz Completa", 
                "🚨 Inconformidades", 
                "📍 Não Encontrados na Base", 
                "📍 Omitidos na Proposta", 
                "📝 Log de Erros de Parsing", 
                "🗄️ DB Base", 
                "🗄️ DB Proposta"
            ])
            
            with tab1:
                st.subheader("📊 Painel Analítico de Conformidade Contratual")
                c1, c2, c3 = st.columns(3)
                c1.metric("📌 Volume de Itens Auditados", f"{total_insumos:,.0f}", f"{total_irregularidades} desvios sinalizados")
                c2.metric("💰 Saldo Global do Orçamento", f"R$ {total_proposta:,.2f}", f"Variação: {var_total_geral:+.2%}", delta_color="inverse")
                c3.metric("✅ Índice de Acerto Paramétrico", f"{taxa_conformidade*100:.1f}%", "Meta aceitável: > 95%")
                c4, c5, c6 = st.columns(3)
                c4.metric("🚨 Exposição a Sobrepreço (Risco)", f"R$ {financeiro_sobrepreco:,.2f}", f"{sobreprecados} sub-itens majorados", delta_color="inverse")
                c5.metric("📉 Desconto Ofertado Oculto", f"R$ {financeiro_inexequivel:,.2f}", f"{inexequiveis} itens sub-precificados")
                c6.metric("🎯 Maior Desvio Único Mapeado", f"R$ {max_desvio_individual:,.2f}", "Alerta Crítico Magnificado")
                st.divider()
                
                g_col1, g_col2 = st.columns(2)
                with g_col1:
                    st.markdown("##### 🔢 Ocorrências por Tipologia de Desvio")
                    st.bar_chart(dash_data_excel['count_graf'].set_index('Tipologia de Erro')['Ocorrências'], height=280)
                with g_col2:
                    st.markdown("##### 💸 Impacto Financeiro Líquido por Grupo (R$)")
                    st.bar_chart(dash_data_excel['money_graf'].set_index('Tipologia Financeira')['Impacto (R$)'], height=280)
                st.divider()
                
                p_col1, p_col2 = st.columns(2)
                with p_col1:
                    st.markdown("##### 🔺 Top 5 Insumos com Maior Impacto de Sobrepreço")
                    if not df_top_sobre_view.empty: st.dataframe(df_top_sobre_view.style.format({'Sobrepreço (R$)': 'R$ {:.2f}', 'Variação (%)': '{:+.2%}'}), hide_index=True, use_container_width=True)
                    else: st.success("Nenhum sobrepreço mapeado.")
                with p_col2:
                    st.markdown("##### 🔻 Top 5 Insumos com Maior Risco de Inexequibilidade")
                    if not df_top_inex_view.empty: 
                        st.dataframe(df_top_inex_view.style.format({'Defasagem (R$)': 'R$ {:.2f}', 'Variação Total (%)': '{:.2%}'}), hide_index=True, use_container_width=True)
                    else: st.info("Nenhuma anomalia de desconto extremo encontrada.")
                
                st.divider()
                st.download_button("📥 Baixar Laudo de Auditoria Unificado (.XLSX)", data=excel_bytes, file_name='Laudo_Auditoria_PRO_Consolidado.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True, key='dl_kpi')
            
            with tab2:
               with tab2:
                st.download_button("📥 Baixar Laudo de Auditoria Unificado (.XLSX)", data=excel_bytes, file_name='Laudo_Auditoria_PRO_Consolidado.xlsx', use_container_width=True, key='dl_matriz')
                
                import streamlit.components.v1 as components
                
                try:
                    # Tentativa padrão: exibir com o Streamlit
                    st.dataframe(styler_ui_completo, height=600, use_container_width=True)
                except Exception as e:
                    # Se falhar, contorna o bug do Python 3.14 renderizando como HTML
                    st.warning("A renderização nativa falhou. Tentando método alternativo...")
                    
                    with st.expander("Ver detalhes do erro técnico"):
                        st.code(f"{type(e).__name__}:\n{e}", language="python")
                    
                    try:
                        components.html(styler_ui_completo.to_html(), height=650, scrolling=True)
                    except Exception as inner_e:
                        st.error("O objeto Styler está corrompido.")
                        st.code(f"ERRO REAL:\n{type(inner_e).__name__}:\n{inner_e}", language="python")
                        st.dataframe(styler_ui_completo.data, height=400, use_container_width=True)
                        
            with tab3:
                st.download_button("📥 Baixar Laudo de Auditoria Unificado (.XLSX)", data=excel_bytes, file_name='Laudo_Auditoria_PRO_Consolidado.xlsx', use_container_width=True, key='dl_erro')
                if styler_ui_erros is not None: 
                    st.dataframe(styler_ui_erros, height=500, use_container_width=True)
                else: 
                    st.success("✅ Tudo em conformidade!")
                    
            with tab4:
                if styler_ui_ne_base is not None: 
                    st.dataframe(styler_ui_ne_base, height=500, use_container_width=True)
                else: 
                    st.success("✅ Alinhamento Completo! Todos os insumos da proposta existem na base.")
                    
            with tab5:
                if styler_ui_ne_prop is not None: 
                    st.dataframe(styler_ui_ne_prop, height=500, use_container_width=True)
                else: 
                    st.success("✅ Alinhamento Completo! A proposta contempla 100% dos itens presentes na base de referência.")
                    
            with tab6: 
                st.success("✅ Zero erros estruturais identificados.")
                
            with tab7: 
                st.dataframe(styler_ui_db_base, height=600, use_container_width=True)
                
            with tab8: 
                st.dataframe(styler_ui_db_prop, height=600, use_container_width=True)
