import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl.styles import PatternFill, Font
import hashlib

# ==========================================
# 0. CONFIGURAÇÕES E CACHE
# ==========================================

st.set_page_config(page_title="Auditoria PRO", layout="wide", page_icon="🛡️")

# Constantes
MAX_FILE_SIZE_MB = 50
CACHE_TTL = 3600

# ==========================================
# 1. PARSER AVANÇADO (PADRÃO ORÇAFASCIO)
# ==========================================

@st.cache_data(ttl=CACHE_TTL)
def carregar_orcafascio(arquivo_bytes, nome_arquivo):
    """
    Lê o Orçafascio preservando a ordem e capturando as Unidades para comparação cruzada.
    Retorna: (df_processado, log_erros, checksum)
    """
    try:
        # Validação de tamanho
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
            
            # Detecção de cabeçalho com validação rigorosa
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
                
            # Validação de colunas obrigatórias
            if col_quant is None or col_preco is None:
                continue
                
            try:
                tipo_item = str(row[col_tipo]).strip().lower()
                cod_item = str(row[col_cod]).strip().upper()
                desc_item = str(row[col_desc]).strip() if col_desc is not None else ""
                und_item = str(row[col_und]).strip().upper() if col_und is not None else ""
                
                # Tentativa de conversão numérica com rastreamento de erro
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
        
        # Gerar checksum para validação de integridade
        checksum = hashlib.md5(df_final.to_string().encode()).hexdigest()
        
        # Log de erros
        log_msg = ""
        if log_erros_parse:
            log_msg = f"⚠️ {len(log_erros_parse)} linhas com erro de parsing (veja aba 'Log de Erros')"
        
        return df_final, log_msg, (checksum, len(df_final))
        
    except Exception as e:
        return None, f"Erro ao processar {nome_arquivo}: {str(e)}", None


@st.cache_data(ttl=CACHE_TTL)
def transformar_hierarquico(df):
    """Gera o layout de leitura incorporando a comparação de Unidades."""
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
                'Und_Base': '---',
                'Und_Prop': '---',
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

def estilizar_relatorio(row):
    """
    Taxonomia de Cores:
    - Azul Claro: Estrutura.
    - Amarelo: Fraude Métrica.
    - Vermelho: Sobrepreço Financeiro.
    - Laranja: Inexequibilidade.
    - Roxo: Aumento de Quantitativo.
    """
    if row['Und_Base'] == '---':
        return ['background-color: #dbeafe; font-weight: bold; color: #1e3a8a;'] * len(row)
        
    estilos = [''] * len(row)
    for i, col in enumerate(row.index):
        val = row[col]
        
        if col in ['Und_Base', 'Und_Prop']:
            if str(row['Und_Base']).strip() != str(row['Und_Prop']).strip() and row['Und_Base'] != '---':
                estilos[i] = 'background-color: #fef08a; color: #713f12; font-weight: bold;'
        
        if pd.isna(val) or val == '': 
            continue
        
        try:
            v = float(val)
            # 1. Roxo: Variação de Quantidade
            if col == 'Delta_Qtd' and v > 0:
                estilos[i] = 'background-color: #e9d5ff; color: #6b21a8; font-weight: bold;'
            
            # 2. Vermelho: Sobrepreço Financeiro
            elif col in ['Delta_Preco', 'Delta_Total'] and v > 0:
                estilos[i] = 'background-color: #fca5a5; color: #7f1d1d; font-weight: bold;'
            
            elif col in ['Var_Preco_%', 'Var_Total_%']:
                if v > 0:
                    estilos[i] = 'background-color: #fca5a5; color: #7f1d1d; font-weight: bold;'
                # 3. Laranja: Inexequibilidade
                elif v < st.session_state.get('limiar_desconto', -0.25):
                    estilos[i] = 'background-color: #fdba74; color: #7c2d12; font-weight: bold;'
        except Exception:
            pass
            
    return estilos

def gerar_excel_bytes(styler_excel, styler_erros, df_base, df_prop):
    """Gera Excel injetando Legend, Auto-Fit e formatação matemática NATIVA."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        styler_excel.to_excel(writer, index=False, sheet_name='1. Matriz Completa', startrow=8)
        if styler_erros is not None:
            styler_erros.to_excel(writer, index=False, sheet_name='2. Inconformidades', startrow=8)
            
        def injetar_legenda(ws):
            legendas = [
                ('A1', 'FCA5A5', '🟥 VERMELHO: Sobrepreço (Preço Unit. ou Total superior à base).'),
                ('A2', 'E9D5FF', '🟪 ROXO: Quantidade Adulterada (Quantitativo majorado na proposta).'),
                ('A3', 'FDBA74', '🟧 LARANJA: Inexequibilidade (Desconto excessivo, conforme diligências).'),
                ('A4', 'FEF08A', '🟨 AMARELO: Fraude Métrica (Unidades de Medida incompatíveis).'),
                ('A7', 'DBEAFE', '🟦 AZUL CLARO: Estrutura (Cabeçalho da Composição Pai).')
            ]
            
            for celula, cor, texto in legendas:
                ws[celula] = texto
                ws[celula].fill = PatternFill(start_color=cor, end_color=cor, fill_type='solid')
                ws[celula].font = Font(bold=True)
                
        # Varredura Mestra para injetar formatação nativa (Moeda, % e Decimais) e ajustar largura
        for sheetname in writer.sheets:
            ws = writer.sheets[sheetname]
            
            if 'Matriz' in sheetname or 'Inconformidades' in sheetname:
                injetar_legenda(ws)
                linha_cabecalho = 9 
            else:
                linha_cabecalho = 1
                
            # Mapeia as colunas dinamicamente
            formatos_coluna = {}
            for col_idx in range(1, ws.max_column + 1):
                nome_col = str(ws.cell(row=linha_cabecalho, column=col_idx).value)
                if 'Preco' in nome_col or 'Total' in nome_col:
                    formatos_coluna[col_idx] = '0.00%' if '%' in nome_col else '"R$" #,##0.00'
                elif 'Qtd' in nome_col:
                    formatos_coluna[col_idx] = '#,##0.0000'

            # Aplica os formatos nas células e calcula o Auto-Fit
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                col_idx = col[0].column
                
                for cell in col:
                    # Aplica a formatação nativa do Excel para os números
                    if col_idx in formatos_coluna and cell.row > linha_cabecalho and isinstance(cell.value, (int, float)):
                        cell.number_format = formatos_coluna[col_idx]
                    
                    # Calcula o Auto-Fit
                    try:
                        if cell.value:
                            texto_visivel = str(cell.value)
                            if col_idx in formatos_coluna and isinstance(cell.value, (int, float)):
                                texto_visivel = f"R$ {cell.value:.2f}"
                            max_length = max(max_length, len(texto_visivel))
                    except Exception:
                        pass
                
                largura_ajustada = min(max_length + 2, 95)
                ws.column_dimensions[col_letter].width = largura_ajustada

    return output.getvalue()

# ==========================================
# 2. INTERFACE DO SISTEMA
# ==========================================

# Inicializar session_state para parâmetros
if 'limiar_desconto' not in st.session_state:
    st.session_state.limiar_desconto = -0.25

with st.sidebar:
    st.subheader("⚙️ Parâmetros de Auditoria")
    
    limiar_desconto = st.slider(
        "Limiar de Desconto (Inexequibilidade)",
        min_value=-50,
        max_value=-5,
        value=-25,
        step=1,
        help="Desconto acima deste limiar requer diligências. Ex: -25% = alerta se desconto > 25%"
    )
    
    # Converter para decimal para cálculos internos
    limiar_desconto_decimal = limiar_desconto / 100
    st.session_state.limiar_desconto = limiar_desconto_decimal
    
    # Display do valor atual em formato legível
    st.success(f"📌 **Limiar Configurado: {limiar_desconto}%**\n\n✓ Itens com desconto > {abs(limiar_desconto)}% serão destacados.")
    
    st.divider()
    st.subheader("📌 Legenda de Auditoria")
    st.markdown("""
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fca5a5; color: #7f1d1d;">
        <b>🟥 Vermelho (Sobrepreço)</b><br>Preço unitário ou total superior à base.
    </div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #e9d5ff; color: #6b21a8;">
        <b>🟪 Roxo (Quantidade Adulterada)</b><br>Quantidade do insumo majorada.
    </div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fdba74; color: #7c2d12;">
        <b>🟧 Laranja (Inexequibilidade)</b><br>Desconto excessivo (ajustável acima).
    </div>
    <div style="padding: 10px; border-radius: 5px; margin-bottom: 5px; background-color: #fef08a; color: #713f12;">
        <b>🟨 Amarelo (Fraude Métrica)</b><br>Unidades de Medida incompatíveis.
    </div>
    """, unsafe_allow_html=True)


st.title("🛡️ Auditoria de Orçamentos PRO")
st.markdown("Validação paramétrica: Superfaturamento, Quantitativos, Inexequibilidade e Conformidade Métrica.")
st.divider()

col1, col2 = st.columns(2)
with col1: 
    arquivo_base = st.file_uploader("1. Base de Referência (SINAPI/ORSE)", type=["xlsx", "xls"])
with col2: 
    arquivo_proposta = st.file_uploader("2. Proposta da Empreiteira", type=["xlsx", "xls"])

if arquivo_base and arquivo_proposta:
    with st.spinner("Processando dados, validando integridade e injetando formatação..."):
        
        # Carregar arquivos com validação
        df_base_raw, msg_base, check_base = carregar_orcafascio(arquivo_base.getvalue(), "Base")
        df_prop_raw, msg_prop, check_prop = carregar_orcafascio(arquivo_proposta.getvalue(), "Proposta")
        
        # Mostrar warnings de parsing
        if msg_base:
            st.warning(f"Base: {msg_base}")
        if msg_prop:
            st.warning(f"Proposta: {msg_prop}")
        
        if df_base_raw is not None and df_prop_raw is not None:
            
            # ==========================================
            # VALIDAÇÃO DE INTEGRIDADE DO JOIN
            # ==========================================
            df_base = df_base_raw.copy().set_index(['Servico_Pai', 'Insumo_Filho'])
            df_prop = df_prop_raw.copy().set_index(['Servico_Pai', 'Insumo_Filho'])
            
            # Cruzamento com validação
            df_auditoria = df_base.join(df_prop[['Und', 'Qtd', 'Preco_Unitario']], how='inner', rsuffix='_Prop')
            
            # Identificar itens NÃO encontrados (exclusivo da proposta)
            indices_proposta = set(df_prop.index)
            indices_base = set(df_base.index)
            indices_nao_encontrados = indices_proposta - indices_base
            
            # DataFrame dos itens não encontrados
            if indices_nao_encontrados:
                df_nao_encontrados = df_prop_raw[
                    df_prop_raw.set_index(['Servico_Pai', 'Insumo_Filho']).index.isin(indices_nao_encontrados)
                ].reset_index(drop=True)
            else:
                df_nao_encontrados = pd.DataFrame()
            
            # Alertar se muitas linhas desapareceram
            taxa_match = len(df_auditoria) / len(df_prop_raw) if len(df_prop_raw) > 0 else 0
            if taxa_match < 0.80:
                st.error(
                    f"⚠️ ALERTA CRÍTICO: Apenas {taxa_match:.1%} dos insumos da proposta "
                    f"encontraram correspondência na base. Possível incompatibilidade de estrutura."
                )
                st.info(
                    f"**{len(df_nao_encontrados)} itens da proposta não existem na base.** "
                    f"Veja a aba **'Itens Não Encontrados'** para detalhes."
                )
            elif taxa_match < 0.95:
                st.warning(
                    f"⚠️ ATENÇÃO: {(1-taxa_match):.1%} dos insumos da proposta não encontraram "
                    f"correspondência na base ({len(df_prop_raw) - len(df_auditoria)} linhas perdidas). "
                    f"Veja aba **'Itens Não Encontrados'** para análise."
                )
            
            df_auditoria.rename(columns={
                'Und': 'Und_Base', 'Qtd': 'Qtd_Base', 'Preco_Unitario': 'Preco_Base', 
                'Und_Prop': 'Und_Prop', 'Qtd_Prop': 'Qtd_Prop', 'Preco_Unitario_Prop': 'Preco_Prop'
            }, inplace=True)
            
            df_auditoria['Total_Base'] = df_auditoria['Qtd_Base'] * df_auditoria['Preco_Base']
            df_auditoria['Total_Prop'] = df_auditoria['Qtd_Prop'] * df_auditoria['Preco_Prop']
            
            df_auditoria['Delta_Qtd'] = (df_auditoria['Qtd_Prop'] - df_auditoria['Qtd_Base'])
            df_auditoria['Delta_Preco'] = (df_auditoria['Preco_Prop'] - df_auditoria['Preco_Base'])
            df_auditoria['Delta_Total'] = (df_auditoria['Total_Prop'] - df_auditoria['Total_Base'])
            
            df_auditoria['Var_Preco_%'] = np.where(df_auditoria['Preco_Base'] > 0, (df_auditoria['Preco_Prop'] / df_auditoria['Preco_Base']) - 1, 0)
            df_auditoria['Var_Total_%'] = np.where(df_auditoria['Total_Base'] > 0, (df_auditoria['Total_Prop'] / df_auditoria['Total_Base']) - 1, 0)
            
            df_completo = df_auditoria.reset_index()
            
            irregularidades = df_completo[
                (df_completo['Delta_Qtd'] > 0) | 
                (df_completo['Delta_Preco'] > 0) | 
                (df_completo['Var_Preco_%'] < st.session_state.limiar_desconto) |
                (df_completo['Und_Base'] != df_completo['Und_Prop']) 
            ]
            
            df_visual_completo = transformar_hierarquico(df_completo)
            df_visual_erros = transformar_hierarquico(irregularidades) if not irregularidades.empty else pd.DataFrame()
            
            # Formatação que a TELA vai mostrar (Strings bonitas)
            formato_tela = {
                'Qtd_Base': '{:.4f}', 'Qtd_Prop': '{:.4f}', 'Delta_Qtd': '{:.4f}',
                'Preco_Base': 'R$ {:.2f}', 'Preco_Prop': 'R$ {:.2f}', 'Delta_Preco': 'R$ {:.2f}',
                'Var_Preco_%': '{:.2%}',
                'Total_Base': 'R$ {:.2f}', 'Total_Prop': 'R$ {:.2f}', 'Delta_Total': 'R$ {:.2f}',
                'Var_Total_%': '{:.2%}'
            }
            
            # Separação crucial de responsabilidades:
            # 1. Styler para a UI (números formatados como strings)
            styler_ui_completo = df_visual_completo.style.format(formato_tela, na_rep="").apply(estilizar_relatorio, axis=1)
            styler_ui_erros = df_visual_erros.style.format(formato_tela, na_rep="").apply(estilizar_relatorio, axis=1) if not df_visual_erros.empty else None
            
            # 2. Styler para o Excel (números vivos, apenas colores)
            styler_excel_completo = df_visual_completo.style.apply(estilizar_relatorio, axis=1)
            styler_excel_erros = df_visual_erros.style.apply(estilizar_relatorio, axis=1) if not df_visual_erros.empty else None
            
            # Gera o arquivo Excel
            excel_bytes = gerar_excel_bytes(styler_excel_completo, styler_excel_erros, df_base_raw, df_prop_raw)
            
            # ==========================================
            # PREPARAR DADOS PARA DASHBOARD
            # ==========================================
            
            total_insumos = len(df_completo)
            total_irregularidades = len(irregularidades)
            
            # Desagregação de irregularidades
            sobreprecados = len(df_completo[(df_completo['Delta_Preco'] > 0) | (df_completo['Delta_Total'] > 0)])
            quantidades_alteradas = len(df_completo[df_completo['Delta_Qtd'] > 0])
            unidades_incompativeis = len(df_completo[df_completo['Und_Base'] != df_completo['Und_Prop']])
            inexequiveis = len(df_completo[df_completo['Var_Preco_%'] < st.session_state.limiar_desconto])
            
            # Valores totais
            total_base = df_completo['Total_Base'].sum()
            total_proposta = df_completo['Total_Prop'].sum()
            delta_total = total_proposta - total_base
            var_total_geral = (total_proposta / total_base - 1) if total_base > 0 else 0
            
            # Valores de irregularidades
            valor_sobreprecado = irregularidades['Delta_Total'].sum() if not irregularidades.empty else 0
            
            st.divider()
            
            # Criar tabs
            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "📊 Dashboard KPI",
                "📋 Matriz Completa", 
                "🚨 Inconformidades", 
                "📍 Itens Não Encontrados",
                "📝 Log de Erros de Parsing",
                "🗄️ DB Base", 
                "🗄️ DB Proposta"
            ])
            
            # ==========================================
            # TAB 1: DASHBOARD KPI
            # ==========================================
            with tab1:
                st.subheader("📊 Resumo Executivo da Auditoria")
                
                col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
                
                with col_kpi1:
                    st.metric(
                        "📌 Total de Insumos",
                        f"{total_insumos:,.0f}",
                        delta=f"{total_irregularidades} com irregularidades"
                    )
                
                with col_kpi2:
                    st.metric(
                        "💰 Delta Financeiro",
                        f"R$ {delta_total:,.2f}",
                        delta=f"{var_total_geral:+.2%}",
                        delta_color="inverse"
                    )
                
                with col_kpi3:
                    taxa_conformidade = ((total_insumos - total_irregularidades) / total_insumos * 100) if total_insumos > 0 else 0
                    st.metric(
                        "✅ Taxa de Conformidade",
                        f"{taxa_conformidade:.1f}%",
                        delta=f"{total_irregularidades} desvios"
                    )
                
                with col_kpi4:
                    st.metric(
                        "🚨 Risco Financeiro",
                        f"R$ {abs(valor_sobreprecado):,.2f}",
                        help="Valor total de sobrepreços identificados"
                    )
                
                st.divider()
                
                col_d1, col_d2 = st.columns(2)
                
                with col_d1:
                    st.subheader("Desagregação de Irregularidades")
                    irregularidades_data = {
                        'Tipo': [
                            '🟥 Sobrepreço',
                            '🟪 Qtd. Alterada',
                            '🟨 Unidade Incompat.',
                            '🟧 Inexequível'
                        ],
                        'Quantidade': [
                            sobreprecados,
                            quantidades_alteradas,
                            unidades_incompativeis,
                            inexequiveis
                        ]
                    }
                    df_irregular = pd.DataFrame(irregularidades_data)
                    
                    # Gráfico de barras horizontal
                    st.bar_chart(
                        df_irregular.set_index('Tipo')['Quantidade'],
                        height=300,
                        use_container_width=True
                    )
                
                with col_d2:
                    st.subheader("Resumo Financeiro")
                    
                    summary_data = {
                        'Métrica': ['Orçamento Base', 'Orçamento Proposto', 'Variação Absoluta'],
                        'Valor': [f'R$ {total_base:,.2f}', f'R$ {total_proposta:,.2f}', f'R$ {delta_total:,.2f}']
                    }
                    df_summary = pd.DataFrame(summary_data)
                    st.dataframe(df_summary, hide_index=True, use_container_width=True)
                    
                    # Colorir segunda linha se tiver irregularidades
                    if total_irregularidades > 0:
                        st.info(
                            f"**{total_irregularidades} insumos com desvios**\n\n"
                            f"- {sobreprecados} com sobrepreço\n"
                            f"- {quantidades_alteradas} com quantidade alterada\n"
                            f"- {unidades_incompativeis} com unidade incompatível\n"
                            f"- {inexequiveis} com desconto excessivo"
                        )
                    else:
                        st.success("✅ Orçamento em conformidade total!")
                
                st.divider()
                st.download_button(
                    "📥 Baixar Processo de Auditoria Completo (.XLSX)", 
                    data=excel_bytes, 
                    file_name='Auditoria_PRO_Laudo.xlsx', 
                    width='stretch',
                    key='download_dashboard'
                )
            
            # ==========================================
            # TAB 2: MATRIZ COMPLETA
            # ==========================================
            with tab2:
                st.download_button(
                    "📥 Baixar Processo de Auditoria Completo (.XLSX)", 
                    data=excel_bytes, 
                    file_name='Auditoria_PRO_Laudo.xlsx', 
                    width='stretch',
                    key='download_matriz'
                )
                st.dataframe(styler_ui_completo, width='stretch', height=600, use_container_width=True)
            
            # ==========================================
            # TAB 3: INCONFORMIDADES
            # ==========================================
            with tab3:
                if styler_ui_erros is not None:
                    st.dataframe(styler_ui_erros, width='stretch', height=400, use_container_width=True)
                else:
                    st.success("✅ Tudo em conformidade!")
            
            # ==========================================
            # TAB 4: ITENS NÃO ENCONTRADOS NA BASE
            # ==========================================
            with tab4:
                st.subheader("📍 Itens da Proposta que Não Existem na Base")
                
                if not df_nao_encontrados.empty:
                    st.error(
                        f"🚨 **{len(df_nao_encontrados)} insumos** da proposta não foram encontrados na base de referência. "
                        f"Isso representa **{(len(df_nao_encontrados)/len(df_prop_raw)*100):.1f}%** do total."
                    )
                    
                    st.info(
                        "**Possíveis causas:**\n"
                        "1. Diferença entre tabelas de referência (SINAPI vs ORSE ou versões diferentes)\n"
                        "2. Erros de digitação nos códigos de serviço/insumo\n"
                        "3. Insumos especiais ou adaptados não presentes na base\n"
                        "4. Estrutura hierárquica incompatível"
                    )
                    
                    st.divider()
                    
                    # Preparar visualização com agrupamento por Serviço Pai
                    df_agrp = df_nao_encontrados.groupby('Servico_Pai').agg({
                        'Descricao_Pai': 'first',
                        'Insumo_Filho': 'count',
                        'Qtd': 'sum',
                        'Preco_Unitario': 'mean'
                    }).rename(columns={
                        'Descricao_Pai': 'Descrição do Serviço',
                        'Insumo_Filho': 'Qtd de Insumos',
                        'Qtd': 'Total Quantitativo',
                        'Preco_Unitario': 'Preço Médio'
                    })
                    
                    st.subheader("Resumo por Serviço (Composição Pai)")
                    st.dataframe(
                        df_agrp.style.format({
                            'Qtd de Insumos': '{:.0f}',
                            'Total Quantitativo': '{:.4f}',
                            'Preço Médio': 'R$ {:.2f}'
                        }),
                        use_container_width=True
                    )
                    
                    st.divider()
                    
                    # Tabela detalhada
                    st.subheader("Detalhamento Completo")
                    
                    df_detalhe = df_nao_encontrados[[
                        'Servico_Pai', 'Descricao_Pai', 'Insumo_Filho', 
                        'Descricao_Filho', 'Und', 'Qtd', 'Preco_Unitario'
                    ]].copy()
                    
                    df_detalhe['Total'] = df_detalhe['Qtd'] * df_detalhe['Preco_Unitario']
                    df_detalhe = df_detalhe.rename(columns={
                        'Servico_Pai': 'Código Serviço',
                        'Descricao_Pai': 'Descrição Serviço',
                        'Insumo_Filho': 'Código Insumo',
                        'Descricao_Filho': 'Descrição Insumo',
                        'Und': 'Unidade',
                        'Qtd': 'Quantidade',
                        'Preco_Unitario': 'Preço Unit.'
                    })
                    
                    # Formatação para visualização
                    df_display = df_detalhe.copy()
                    df_display['Quantidade'] = df_display['Quantidade'].apply(lambda x: f"{x:.4f}")
                    df_display['Preço Unit.'] = df_display['Preço Unit.'].apply(lambda x: f"R$ {x:.2f}")
                    df_display['Total'] = df_display['Total'].apply(lambda x: f"R$ {x:.2f}")
                    
                    st.dataframe(df_display, use_container_width=True, height=500)
                    
                    # Botão para download
                    st.divider()
                    
                    # Criar arquivo de exportação
                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                        df_detalhe.to_excel(writer, sheet_name='Itens Não Encontrados', index=False)
                        df_agrp.to_excel(writer, sheet_name='Resumo por Serviço')
                    
                    st.download_button(
                        "📥 Baixar Lista de Itens Não Encontrados (.XLSX)",
                        data=output_excel.getvalue(),
                        file_name='Itens_Nao_Encontrados.xlsx',
                        key='download_nao_encontrados',
                        width='stretch'
                    )
                    
                else:
                    st.success(
                        "✅ **Perfeito!** Todos os insumos da proposta foram encontrados na base. "
                        "Taxa de correspondência: 100%"
                    )
            
            # ==========================================
            # TAB 5: LOG DE ERROS DE PARSING
            # ==========================================
            with tab5:
                # Extrair logs de ambos os arquivos
                logs_combinados = []
                
                # Logs do arquivo base
                if 'Status_Parsing' in df_base_raw.columns:
                    erros_base = df_base_raw[df_base_raw['Status_Parsing'] != 'OK']
                    if not erros_base.empty:
                        logs_combinados.append(("Base", erros_base[['Insumo_Filho', 'Descricao_Filho']]))
                
                # Logs do arquivo proposta
                if 'Status_Parsing' in df_prop_raw.columns:
                    erros_prop = df_prop_raw[df_prop_raw['Status_Parsing'] != 'OK']
                    if not erros_prop.empty:
                        logs_combinados.append(("Proposta", erros_prop[['Insumo_Filho', 'Descricao_Filho']]))
                
                if logs_combinados:
                    for origem, df_erros in logs_combinados:
                        st.subheader(f"Erros em {origem}")
                        st.dataframe(df_erros, use_container_width=True)
                else:
                    st.success("✅ Nenhum erro de parsing encontrado!")
            
            # ==========================================
            # TAB 6 & 7: DATABASES
            # ==========================================
            with tab6: 
                st.dataframe(df_base_raw, width='stretch', height=400, use_container_width=True)
            
            with tab7: 
                st.dataframe(df_prop_raw, width='stretch', height=400, use_container_width=True)
