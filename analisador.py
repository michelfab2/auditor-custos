import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import hashlib

# ==========================================
# 0. CONFIGURAÇÕES E CACHE
# ==========================================

st.set_page_config(page_title="Auditoria PRO", layout="wide", page_icon="🛡️")

# Constantes
MAX_FILE_SIZE_MB = 50
CACHE_TTL = 3600

# ==========================================
# 1. PARSERS E TRANSFORMADORES HIERÁRQUICOS
# ==========================================

@st.cache_data(ttl=CACHE_TTL)
def carregar_orcafascio(arquivo_bytes, nome_arquivo):
    """
    Lê o Orçafascio preservando a ordem e capturando as Unidades para comparação cruzada.
    Retorna: (df_processado, log_erros, checksum)
    """
    try:
        tamanho_mb = len(arquivo_bytes) / (1024 * 1024)
        if tamanho_mb > MAX_FILE_SIZE_MB:
            return None, f"Arquivo {nome_arquivo} excede {MAX_FILE_SIZE_MB}MB ({tamanho_mb:.1f}MB)", None
        
        df_raw = pd.read_excel(io.BytesIO(arquivo_bytes), header=None)
        
        dados = []
        log_erros_parse = []
        cod_pai_atual, desc_pai_atual = None, ""
        ordem_sequencial = 0
        
        col_tipo = 0
        col_cod, col_desc, col_und, col_quant, col_preco = 1, None, None, None, None
        
        for idx, row in df_raw.iterrows():
            row_lower = [str(x).strip().lower() for x in row]
            
            if any(h in row_lower for h in ['código', 'codigo']):
                for i, val in enumerate(row_lower):
                    if val in ['código', 'codigo']: 
                        col_cod = i
                    elif any(d in val for d in ['descrição', 'descricao', 'desc']): 
                        col_desc = i
                    elif any(u in val for u in ['und', 'unidade', 'unid', 'u.m.']): 
                        col_und = i
                    elif any(q in val for q in ['quant', 'quantidade', 'qtd']): 
                        col_quant = i
                    elif any(p in val for p in ['valor unit', 'preço unit', 'preco unit', 'unit', 'preço', 'preco']): 
                        col_preco = i
                continue
                
            if col_quant is None or col_preco is None:
                continue
                
            try:
                tipo_item = str(row[col_tipo]).strip().lower()
                cod_item = str(row[col_cod]).strip().upper()
                desc_item = str(row[col_desc]).strip() if col_desc is not None else ""
                und_item = str(row[col_und]).strip().upper() if col_und is not None else ""
                
                try:
                    qtd_valor = float(row[col_quant])
                    preco_valor = float(row[col_preco])
                except (ValueError, TypeError):
                    log_erros_parse.append({
                        'Linha': idx + 1,
                        'Código': cod_item,
                        'Descrição': desc_item,
                        'Erro': f"Conversão numérica falhou (Qtd ou Preço inválidos)"
                    })
                    continue
                
                if tipo_item in ['composição', 'composicao']:
                    cod_pai_atual = cod_item
                    desc_pai_atual = desc_item
                    
                elif tipo_item in ['insumo', 'composição auxiliar', 'composicao auxiliar']:
                    if cod_pai_atual and cod_item and cod_item != 'NAN':
                        ordem_sequencial += 1
                        dados.append({
                            'Ordem': ordem_sequencial,
                            'Servico_Pai': cod_pai_atual,
                            'Descricao_Pai': desc_pai_atual,
                            'Insumo_Filho': cod_item,
                            'Descricao_Filho': desc_item,
                            'Und': und_item,
                            'Qtd': qtd_valor,
                            'Preco_Unitario': preco_valor,
                            'Status_Parsing': 'OK'
                        })
                        
            except Exception as e:
                log_erros_parse.append({
                    'Linha': idx + 1,
                    'Código': str(row[col_cod]) if col_cod < len(row) else 'N/A',
                    'Descrição': str(row[col_desc]) if col_desc and col_desc < len(row) else 'N/A',
                    'Erro': str(e)
                })
                    
        df_final = pd.DataFrame(dados)
        if df_final.empty:
            return None, f"Planilha {nome_arquivo}: Nenhum dado válido extraído.", None
        
        checksum = hashlib.md5(df_final.to_string().encode()).hexdigest()
        log_msg = f"⚠️ {len(log_erros_parse)} linhas com erro de parsing" if log_erros_parse else ""
        
        return df_final, log_msg, (checksum, len(df_final))
        
    except Exception as e:
        return None, f"Erro ao processar {nome_arquivo}: {str(e)}", None


@st.cache_data(ttl=CACHE_TTL)
def transformar_hierarquico(df):
    """Layout hierárquico cruzado para a Matriz de Auditoria."""
    if df.empty: 
        return pd.DataFrame()
    
    df = df.sort_values('Ordem')
    linhas = []
    pai_atual = None
    colunas_vazias = ['Código', 'Descrição', 'Und_Base', 'Und_Prop', 'Qtd_Base', 'Qtd_Prop', 'Delta_Qtd', 'Preco_Base', 'Preco_Prop', 'Delta_Preco', 'Var_Preco_%', 'Total_Base', 'Total_Prop', 'Delta_Total', 'Var_Total_%']
    
    for _, row in df.iterrows():
        if row['Servico_Pai'] != pai_atual:
            if pai_atual is not None:
                linhas.append({c: np.nan for c in colunas_vazias})
            
            linhas.append({
                'Código': row['Servico_Pai'],
                'Descrição': f"COMPOSIÇÃO: {row['Descricao_Pai']}",
                'Und_Base': '---', 'Und_Prop': '---',
                **{c: np.nan for c in colunas_vazias[4:]}
            })
            pai_atual = row['Servico_Pai']
            
        linhas.append({
            'Código': row['Insumo_Filho'], 'Descrição': row['Descricao_Filho'], 
            'Und_Base': row['Und_Base'], 'Und_Prop': row['Und_Prop'],
            'Qtd_Base': row['Qtd_Base'], 'Qtd_Prop': row['Qtd_Prop'], 'Delta_Qtd': row['Delta_Qtd'],
            'Preco_Base': row['Preco_Base'], 'Preco_Prop': row['Preco_Prop'], 'Delta_Preco': row['Delta_Preco'],
            'Var_Preco_%': row['Var_Preco_%'],
            'Total_Base': row['Total_Base'], 'Total_Prop': row['Total_Prop'], 'Delta_Total': row['Delta_Total'],
            'Var_Total_%': row['Var_Total_%']
        })
    return pd.DataFrame(linhas)


@st.cache_data(ttl=CACHE_TTL)
def transformar_hierarquico_raw(df):
    """Aplica a árvore estrutural para as bases de dados brutas e itens não encontrados."""
    if df.empty: 
        return pd.DataFrame()
    
    df = df.sort_values('Ordem')
    linhas = []
    pai_atual = None
    
    df_copy = df.copy()
    df_copy['Total'] = df_copy['Qtd'] * df_copy['Preco_Unitario']
    colunas_esquema = ['Código', 'Descrição', 'Unidade', 'Quantidade', 'Preço Unitário', 'Total']
    
    for _, row in df_copy.iterrows():
        if row['Servico_Pai'] != pai_atual:
            if pai_atual is not None:
                # Linha de respiro em branco
                linhas.append({c: np.nan for c in colunas_esquema})
            
            # Cabeçalho da composição pai
            linhas.append({
                'Código': row['Servico_Pai'],
                'Descrição': f"COMPOSIÇÃO: {row['Descricao_Pai']}",
                'Unidade': '---',
                'Quantidade': np.nan,
                'Preço Unitário': np.nan,
                'Total': np.nan
            })
            pai_atual = row['Servico_Pai']
            
        linhas.append({
            'Código': row['Insumo_Filho'],
            'Descrição': row['Descricao_Filho'],
            'Unidade': row['Und'],
            'Quantidade': row['Qtd'],
            'Preço Unitário': row['Preco_Unitario'],
            'Total': row['Total']
        })
    return pd.DataFrame(linhas)

# ==========================================
# 2. ESTILIZADORES DA INTERFACE (UI)
# ==========================================

def estilizar_relatorio(row):
    """Estilização condicional para as matrizes de comparação."""
    if row['Und_Base'] == '---':
        return ['background-color: #dbeafe; font-weight: bold; color: #1e3a8a;'] * len(row)
        
    estilos = [''] * len(row)
    for i, col in enumerate(row.index):
        val = row[col]
        if col in ['Und_Base', 'Und_Prop'] and str(row['Und_Base']).strip() != str(row['Und_Prop']).strip() and row['Und_Base'] != '---':
            estilos[i] = 'background-color: #fef08a; color: #713f12; font-weight: bold;'
        
        if pd.isna(val) or val == '': continue
        try:
            v = float(val)
            if col == 'Delta_Qtd' and v > 0:
                estilos[i] = 'background-color: #e9d5ff; color: #6b21a8; font-weight: bold;'
            elif col in ['Delta_Preco', 'Delta_Total'] and v > 0:
                estilos[i] = 'background-color: #fca5a5; color: #7f1d1d; font-weight: bold;'
            elif col in ['Var_Preco_%', 'Var_Total_%']:
                if v > 0: estilos[i] = 'background-color: #fca5a5; color: #7f1d1d; font-weight: bold;'
                elif v < st.session_state.get('limiar_desconto', -0.25):
                    estilos[i] = 'background-color: #fdba74; color: #7c2d12; font-weight: bold;'
        except Exception: pass
    return estilos

def estilizar_relatorio_raw(row):
    """Estilização para os cabeçalhos de composição nas abas de tabelas brutas."""
    if row['Unidade'] == '---':
        return ['background-color: #dbeafe; font-weight: bold; color: #1e3a8a;'] * len(row)
    return [''] * len(row)

# ==========================================
# 3. MOTOR EXPORTADOR EXCEL MULTI-ABA
# ==========================================

def gerar_excel_bytes(df_kpi, df_matriz, df_inconformidades, df_nao_encontrados, df_parsing, df_db_base, df_db_prop):
    """Gera o laudo unificado em formato executivo respeitando as árvores estruturais."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        df_kpi.to_excel(writer, index=False, sheet_name='📊 Dashboard KPI', startrow=1)
        df_matriz.to_excel(writer, index=False, sheet_name='📋 Matriz Completa', startrow=8)
        
        if not df_inconformidades.empty:
            df_inconformidades.to_excel(writer, index=False, sheet_name='🚨 Inconformidades', startrow=8)
        else:
            ws_inc = writer.book.create_sheet(title='🚨 Inconformidades')
            ws_inc['A1'] = "Nenhuma inconformidade paramétrica detectada."
            ws_inc['A1'].font = Font(bold=True, size=11)
            
        df_nao_encontrados.to_excel(writer, index=False, sheet_name='📍 Itens Não Encontrados', startrow=1)
        df_parsing.to_excel(writer, index=False, sheet_name='📝 Log de Erros de Parsing', startrow=1)
        df_db_base.to_excel(writer, index=False, sheet_name='🗄️ DB Base', startrow=1)
        df_db_prop.to_excel(writer, index=False, sheet_name='🗄️ DB Proposta', startrow=1)
        
        def injetar_legenda(ws):
            legendas = [
                ('A1', 'FCA5A5', '🟥 VERMELHO: Sobrepreço (Valor unitário ou total superior à referência).'),
                ('A2', 'E9D5FF', '🟪 ROXO: Quantidade Adulterada (Quantitativo majorado na proposta).'),
                ('A3', 'FDBA74', '🟧 LARANJA: Inexequibilidade (Desconto excessivo fora das margens).'),
                ('A4', 'FEF08A', '🟨 AMARELO: Fraude Métrica (Unidades de medida incompatíveis).'),
                ('A7', 'DBEAFE', '🟦 AZUL CLARO: Estrutura (Linha de Cabeçalho da Composição Analítica Pai).')
            ]
            for celula, cor, texto in legendas:
                ws[celula] = texto
                ws[celula].fill = PatternFill(start_color=cor, end_color=cor, fill_type='solid')
                ws[celula].font = Font(bold=True, size=9)

        wb = writer.book
        for name in wb.sheetnames:
            ws = wb[name]
            
            # Configuração de contexto e linha de cabeçalho por tipo de aba
            if name in ['📋 Matriz Completa', '🚨 Inconformidades']:
                injetar_legenda(ws)
                linha_cabecalho = 9
                col_unidade_idx = 3 # Coluna 'Und_Base'
            elif name in ['📍 Itens Não Encontrados', '🗄️ DB Base', '🗄️ DB Proposta']:
                linha_cabecalho = 1
                col_unidade_idx = 3 # Coluna 'Unidade'
            else:
                linha_cabecalho = 1
                col_unidade_idx = None
            
            # Processamento de cores hierárquicas nativas nas planilhas de engenharia
            if col_unidade_idx is not None:
                for r_idx in range(linha_cabecalho + 1, ws.max_row + 1):
                    und_val = str(ws.cell(row=r_idx, column=col_unidade_idx).value).strip()
                    
                    if und_val == '---':
                        azul_fill = PatternFill(start_color='DBEAFE', end_color='DBEAFE', fill_type='solid')
                        for c_idx in range(1, ws.max_column + 1):
                            cell = ws.cell(row=r_idx, column=c_idx)
                            cell.fill = azul_fill
                            cell.font = Font(bold=True, color='1E3A8A')
                        continue
                    
                    # Estilos adicionais exclusivos para as abas cruzadas
                    if name in ['📋 Matriz Completa', '🚨 Inconformidades']:
                        try:
                            und_prop_val = str(ws.cell(row=r_idx, column=4).value).strip()
                            delta_qtd = float(ws.cell(row=r_idx, column=7).value or 0)
                            delta_preco = float(ws.cell(row=r_idx, column=10).value or 0)
                            var_preco_p = float(ws.cell(row=r_idx, column=11).value or 0)
                            
                            if und_val != und_prop_val and und_val != 'None':
                                ws.cell(row=r_idx, column=3).fill = PatternFill(start_color='FEF08A', end_color='FEF08A', fill_type='solid')
                                ws.cell(row=r_idx, column=4).fill = PatternFill(start_color='FEF08A', end_color='FEF08A', fill_type='solid')
                            if delta_qtd > 0:
                                ws.cell(row=r_idx, column=7).fill = PatternFill(start_color='E9D5FF', end_color='E9D5FF', fill_type='solid')
                            if delta_preco > 0:
                                ws.cell(row=r_idx, column=10).fill = PatternFill(start_color='FCA5A5', end_color='FCA5A5', fill_type='solid')
                            if var_preco_p > 0:
                                ws.cell(row=r_idx, column=11).fill = PatternFill(start_color='FCA5A5', end_color='FCA5A5', fill_type='solid')
                            if var_preco_p < st.session_state.get('limiar_desconto', -0.25):
                                ws.cell(row=r_idx, column=11).fill = PatternFill(start_color='FDBA74', end_color='FDBA74', fill_type='solid')
                        except Exception: pass

            # Formatação numérica viva e auto-ajuste inteligente de colunas
            formatos_coluna = {}
            for col_idx in range(1, ws.max_column + 1):
                nome_col = str(ws.cell(row=linha_cabecalho, column=col_idx).value)
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
                header_cell.fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
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
                    except Exception: pass
                ws.column_dimensions[col_letter].width = min(max(max_length + 3, 11), 70)
                
    return output.getvalue()

# ==========================================
# 4. INTERFACE GRÁFICA DO USUÁRIO (STREAMLIT)
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
    st.markdown("""
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fca5a5; color: #7f1d1d; font-family: sans-serif; font-size:13px;"><b>🟥 Vermelho (Sobrepreço)</b><br>Preço unitário ou total superior à base.</div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #e9d5ff; color: #6b21a8; font-family: sans-serif; font-size:13px;"><b>🟪 Roxo (Qtd. Alterada)</b><br>Quantidade do insumo majorada.</div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fdba74; color: #7c2d12; font-family: sans-serif; font-size:13px;"><b>🟧 Laranja (Inexequibilidade)</b><br>Desconto excessivo.</div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fef08a; color: #713f12; font-family: sans-serif; font-size:13px;"><b>🟨 Amarelo (Fraude Métrica)</b><br>Unidades de Medida incompatíveis.</div>
    """, unsafe_allow_html=True)

st.title("🛡️ Auditoria de Orçamentos PRO")
st.markdown("Validação paramétrica: Superfaturamento, Quantitativos, Inexequibilidade e Conformidade Métrica.")
st.divider()

col1, col2 = st.columns(2)
with col1: arquivo_base = st.file_uploader("1. Base de Referência (SINAPI/ORSE)", type=["xlsx", "xls"])
with col2: arquivo_proposta = st.file_uploader("2. Proposta da Empreiteira", type=["xlsx", "xls"])

if arquivo_base and arquivo_proposta:
    with st.spinner("Estruturando laudos, aplicando árvores analíticas e organizando respiros..."):
        
        df_base_raw, msg_base, check_base = carregar_orcafascio(arquivo_base.getvalue(), "Base")
        df_prop_raw, msg_prop, check_prop = carregar_orcafascio(arquivo_proposta.getvalue(), "Proposta")
        
        if msg_base: st.warning(f"Base: {msg_base}")
        if msg_prop: st.warning(f"Proposta: {msg_prop}")
        
        if df_base_raw is not None and df_prop_raw is not None:
            
            df_base = df_base_raw.copy().set_index(['Servico_Pai', 'Insumo_Filho'])
            df_prop = df_prop_raw.copy().set_index(['Servico_Pai', 'Insumo_Filho'])
            
            df_auditoria = df_base.join(df_prop[['Und', 'Qtd', 'Preco_Unitario']], how='inner', rsuffix='_Prop')
            indices_nao_encontrados = set(df_prop.index) - set(df_base.index)
            
            df_nao_encontrados = df_prop_raw[df_prop_raw.set_index(['Servico_Pai', 'Insumo_Filho']).index.isin(indices_nao_encontrados)].reset_index(drop=True) if indices_nao_encontrados else pd.DataFrame()
            
            taxa_match = len(df_auditoria) / len(df_prop_raw) if len(df_prop_raw) > 0 else 0
            if taxa_match < 0.95:
                st.warning(f"⚠️ ATENÇÃO: {(1-taxa_match):.1%} dos insumos da proposta não encontraram correspondência na base de dados de referência.")
            
            df_auditoria.rename(columns={'Und': 'Und_Base', 'Qtd': 'Qtd_Base', 'Preco_Unitario': 'Preco_Base', 'Und_Prop': 'Und_Prop', 'Qtd_Prop': 'Qtd_Prop', 'Preco_Unitario_Prop': 'Preco_Prop'}, inplace=True)
            df_auditoria['Total_Base'] = df_auditoria['Qtd_Base'] * df_auditoria['Preco_Base']
            df_auditoria['Total_Prop'] = df_auditoria['Qtd_Prop'] * df_auditoria['Preco_Prop']
            df_auditoria['Delta_Qtd'] = df_auditoria['Qtd_Prop'] - df_auditoria['Qtd_Base']
            df_auditoria['Delta_Preco'] = df_auditoria['Preco_Prop'] - df_auditoria['Preco_Base']
            df_auditoria['Delta_Total'] = df_auditoria['Total_Prop'] - df_auditoria['Total_Base']
            df_auditoria['Var_Preco_%'] = np.where(df_auditoria['Preco_Base'] > 0, (df_auditoria['Preco_Prop'] / df_auditoria['Preco_Base']) - 1, 0)
            df_auditoria['Var_Total_%'] = np.where(df_auditoria['Total_Base'] > 0, (df_auditoria['Total_Prop'] / df_auditoria['Total_Base']) - 1, 0)
            
            df_completo = df_auditoria.reset_index()
            irregularidades = df_completo[(df_completo['Delta_Qtd'] > 0) | (df_completo['Delta_Preco'] > 0) | (df_completo['Var_Preco_%'] < st.session_state.limiar_desconto) | (df_completo['Und_Base'] != df_completo['Und_Prop'])]
            
            # --------------------------------------------------
            # GERANDO AS ESTRUTURAS HIERÁRQUICAS PARA AS ABAS
            # --------------------------------------------------
            df_visual_completo = transformar_hierarquico(df_completo)
            df_visual_erros = transformar_hierarquico(irregularidades) if not irregularidades.empty else pd.DataFrame()
            
            # Aplicação da nova lógica de árvore com espaçamentos nas outras tabelas
            df_visual_ne = transformar_hierarquico_raw(df_nao_encontrados)
            df_visual_db_base = transformar_hierarquico_raw(df_base_raw)
            df_visual_db_prop = transformar_hierarquico_raw(df_prop_raw)
            
            # Máscaras de Formatação de Visualização na UI
            formato_tela = {
                'Qtd_Base': '{:.4f}', 'Qtd_Prop': '{:.4f}', 'Delta_Qtd': '{:.4f}',
                'Preco_Base': 'R$ {:.2f}', 'Preco_Prop': 'R$ {:.2f}', 'Delta_Preco': 'R$ {:.2f}',
                'Var_Preco_%': '{:.2%}', 'Total_Base': 'R$ {:.2f}', 'Total_Prop': 'R$ {:.2f}',
                'Delta_Total': 'R$ {:.2f}', 'Var_Total_%': '{:.2%}'
            }
            formato_tela_raw = {'Quantidade': '{:.4f}', 'Preço Unitário': 'R$ {:.2f}', 'Total': 'R$ {:.2f}'}
            
            # Construindo os Stylers das tabelas para renderização no Streamlit
            styler_ui_completo = df_visual_completo.style.format(formato_tela, na_rep="").apply(estilizar_relatorio, axis=1)
            styler_ui_erros = df_visual_erros.style.format(formato_tela, na_rep="").apply(estilizar_relatorio, axis=1) if not df_visual_erros.empty else None
            
            styler_ui_ne = df_visual_ne.style.format(formato_tela_raw, na_rep="").apply(estilizar_relatorio_raw, axis=1) if not df_visual_ne.empty else None
            styler_ui_db_base = df_visual_db_base.style.format(formato_tela_raw, na_rep="").apply(estilizar_relatorio_raw, axis=1)
            styler_ui_db_prop = df_visual_db_prop.style.format(formato_tela_raw, na_rep="").apply(estilizar_relatorio_raw, axis=1)
            
            # Cálculos do Dashboard Resumo
            total_insumos, total_irregularidades = len(df_completo), len(irregularidades)
            sobreprecados = len(df_completo[(df_completo['Delta_Preco'] > 0) | (df_completo['Delta_Total'] > 0)])
            quantidades_alteradas = len(df_completo[df_completo['Delta_Qtd'] > 0])
            unidades_incompativeis = len(df_completo[df_completo['Und_Base'] != df_completo['Und_Prop']])
            inexequiveis = len(df_completo[df_completo['Var_Preco_%'] < st.session_state.limiar_desconto])
            
            total_base, total_proposta = df_completo['Total_Base'].sum(), df_completo['Total_Prop'].sum()
            delta_total = total_proposta - total_base
            var_total_geral = (total_proposta / total_base - 1) if total_base > 0 else 0
            valor_sobreprecado = irregularidades['Delta_Total'].sum() if not irregularidades.empty else 0
            taxa_conformidade = ((total_insumos - total_irregularidades) / total_insumos) if total_insumos > 0 else 0
            
            df_kpi_excel = pd.DataFrame({
                'Métrica de Controle': [
                    'Total de Insumos Auditados', 'Insumos com Irregularidades', 'Taxa de Conformidade Geral',
                    'Valor Total Orçamento Base', 'Valor Total Orçamento Proposto', 'Delta Financeiro Absoluto', 
                    'Variação Percentual Global', 'Risco Financeiro Total (Sobrepreço)',
                    'Qtd Itens Sobreprecados (🟥)', 'Qtd Itens com Quantidade Alterada (🟪)', 
                    'Qtd Itens com Unidade Incompatível (🟨)', 'Qtd Itens Inexequíveis (🟧)'
                ],
                'Valor Encontrado': [
                    total_insumos, total_irregularidades, taxa_conformidade,
                    total_base, total_proposta, delta_total, var_total_geral, valor_sobreprecado,
                    sobreprecados, quantidades_alteradas, unidades_incompativeis, inexequiveis
                ],
                'Status': ['OK', 'Alerta' if total_irregularidades > 0 else 'Conforme', 'Crítico' if taxa_conformidade < 0.8 else 'OK', 'Referencial', 'Análise', 'Desvio', 'Atenção', 'Risco', 'Diligenciar', 'Diligenciar', 'Corrigir', 'Avaliar']
            })
            
            logs_parsing_lista = []
            if 'Status_Parsing' in df_base_raw.columns:
                for _, r in df_base_raw[df_base_raw['Status_Parsing'] != 'OK'].iterrows():
                    logs_parsing_lista.append({'Origem': 'Base Referência', 'Código': r['Insumo_Filho'], 'Descrição': r['Descricao_Filho'], 'Erro': 'Falha estrutural'})
            if 'Status_Parsing' in df_prop_raw.columns:
                for _, r in df_prop_raw[df_prop_raw['Status_Parsing'] != 'OK'].iterrows():
                    logs_parsing_lista.append({'Origem': 'Proposta Empreiteira', 'Código': r['Insumo_Filho'], 'Descrição': r['Descricao_Filho'], 'Erro': 'Falha estrutural'})
            df_parsing_excel = pd.DataFrame(logs_parsing_lista) if logs_parsing_lista else pd.DataFrame(columns=['Origem', 'Código', 'Descrição', 'Erro'])
            
            # Compilação unificada para o download
            excel_bytes = gerar_excel_bytes(
                df_kpi=df_kpi_excel, df_matriz=df_visual_completo, df_inconformidades=df_visual_erros,
                df_nao_encontrados=df_visual_ne, df_parsing=df_parsing_excel,
                df_db_base=df_visual_db_base, df_db_prop=df_visual_db_prop
            )
            
            # Renderização de Tabs na Interface Web
            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "📊 Dashboard KPI", "📋 Matriz Completa", "🚨 Inconformidades", 
                "📍 Itens Não Encontrados", "📝 Log de Erros de Parsing", "🗄️ DB Base", "🗄️ DB Proposta"
            ])
            
            with tab1:
                st.subheader("📊 Resumo Executivo da Auditoria")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("📌 Total de Insumos", f"{total_insumos:,.0f}", f"{total_irregularidades} desvios")
                c2.metric("💰 Delta Financeiro", f"R$ {delta_total:,.2f}", f"{var_total_geral:+.2%}", delta_color="inverse")
                c3.metric("✅ Taxa de Conformidade", f"{taxa_conformidade*100:.1f}%")
                c4.metric("🚨 Risco Financeiro", f"R$ {abs(valor_sobreprecado):,.2f}")
                st.divider()
                st.download_button("📥 Baixar Laudo de Auditoria Unificado (.XLSX)", data=excel_bytes, file_name='Laudo_Auditoria_PRO_Estruturado.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True, key='dl_kpi')
            
            with tab2:
                st.download_button("📥 Baixar Laudo de Auditoria Unificado (.XLSX)", data=excel_bytes, file_name='Laudo_Auditoria_PRO_Estruturado.xlsx', use_container_width=True, key='dl_matriz')
                st.dataframe(styler_ui_completo, height=600, use_container_width=True)
            
            with tab3:
                st.download_button("📥 Baixar Laudo de Auditoria Unificado (.XLSX)", data=excel_bytes, file_name='Laudo_Auditoria_PRO_Estruturado.xlsx', use_container_width=True, key='dl_erro')
                if styler_ui_erros is not None: st.dataframe(styler_ui_erros, height=500, use_container_width=True)
                else: st.success("✅ Tudo em conformidade!")
            
            with tab4:
                st.subheader("📍 Insumos da Proposta Ausentes na Base")
                if styler_ui_ne is not None:
                    st.dataframe(styler_ui_ne, height=500, use_container_width=True)
                else:
                    st.success("✅ Alinhamento Completo! Todos os insumos da proposta existem na base.")
            
            with tab5:
                if not df_parsing_excel.empty: st.dataframe(df_parsing_excel, use_container_width=True)
                else: st.success("✅ Zero erros estruturais identificados.")
            
            with tab6:
                st.subheader("🗄️ Base de Referência Organizada por Composição")
                st.dataframe(styler_ui_db_base, height=600, use_container_width=True)
            
            with tab7:
                st.subheader("🗄️ Proposta Comercial Organizada por Composição")
                st.dataframe(styler_ui_db_prop, height=600, use_container_width=True)
