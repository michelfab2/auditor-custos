import io
import re
import unicodedata

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MAX_FILE_SIZE_MB = 50
CACHE_TTL = 3600
SHEET_REALOCADOS = "Realocados - Estrutura"


def texto(valor):
    return "" if pd.isna(valor) else str(valor).strip()


def rotulo(valor):
    """Unifica ComposiçãoAuxiliar e Composição Auxiliar."""
    valor = unicodedata.normalize("NFKD", texto(valor))
    valor = valor.encode("ASCII", "ignore").decode("ASCII").lower()
    return re.sub(r"[^a-z]", "", valor)


def codigo(valor):
    return texto(valor).upper().replace(" ", "")


def numero(valor):
    if pd.isna(valor) or valor in (None, "", "*", "-"):
        return 0.0
    if isinstance(valor, (int, float, np.number)):
        return float(valor)
    valor = re.sub(r"[R$%\s]", "", texto(valor))
    if "," in valor and "." in valor:
        valor = valor.replace(".", "").replace(",", ".")
    elif "," in valor:
        valor = valor.replace(",", ".")
    try:
        return float(valor)
    except ValueError:
        return 0.0


def localizar_colunas(linha, atuais):
    colunas = atuais.copy()
    for indice, valor in enumerate(linha):
        nome = rotulo(valor)
        if nome in {"codigo", "cod"}:
            colunas["cod"] = indice
        elif "descricao" in nome:
            colunas["desc"] = indice
        elif nome in {"und", "unidade", "unid"}:
            colunas["und"] = indice
        elif "quant" in nome or nome in {"qtd", "qtde"}:
            colunas["qtd"] = indice
        elif any(chave in nome for chave in ("valorunit", "precounit", "custounit")):
            colunas["preco"] = indice
    return colunas


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def ler_orcafascio(arquivo_bytes, origem, higienizar=True):
    if len(arquivo_bytes) > MAX_FILE_SIZE_MB * 1024 * 1024:
        return None, pd.DataFrame(), f"{origem}: arquivo excede {MAX_FILE_SIZE_MB} MB."
    try:
        bruto = pd.read_excel(io.BytesIO(arquivo_bytes), header=None)
    except Exception as erro:
        return None, pd.DataFrame(), f"{origem}: não foi possível abrir o arquivo ({erro})."

    mapa = {"cod": 1, "desc": 3, "und": 6, "qtd": 7, "preco": 8}
    itens, erros = [], []
    cpu, descricao_cpu, ordem = None, "", 0
    tipos = {"composicao", "composicaoauxiliar", "insumo", "item"}

    for indice, linha in bruto.iterrows():
        valores = [texto(valor) for valor in linha.tolist()]
        primeiro = valores[0] if valores else ""
        tipo = rotulo(primeiro)

        if re.fullmatch(r"\d+\.\d+", primeiro):
            cpu, descricao_cpu = None, ""
            mapa = localizar_colunas(valores, mapa)
            continue
        if "codigo" in tipo or ("descricao" in tipo and "quant" in "".join(rotulo(v) for v in valores)):
            mapa = localizar_colunas(valores, mapa)
            continue
        if tipo not in tipos:
            continue

        def campo(chave):
            posicao = mapa[chave]
            return valores[posicao] if posicao < len(valores) else ""

        cod = codigo(campo("cod"))
        desc = campo("desc")
        if not cod:
            erros.append({"Origem": origem, "Linha": indice + 1, "Tipo": tipo, "Erro": "Código vazio"})
            continue
        if tipo == "composicao":
            cpu, descricao_cpu = cod, desc
        if cpu is None:
            erros.append({"Origem": origem, "Linha": indice + 1, "Tipo": tipo, "Erro": "Subitem sem composição principal"})
            continue

        ordem += 1
        itens.append({
            "Ordem": ordem, "CPU": cpu, "Descricao_CPU": descricao_cpu,
            "Codigo": cod, "Descricao": desc,
            "Tipo": {"composicao": "Composição", "composicaoauxiliar": "Composição auxiliar", "insumo": "Insumo", "item": "Item"}[tipo],
            "Und": codigo(campo("und")), "Qtd": numero(campo("qtd")), "Preco_Unitario": numero(campo("preco")),
        })

    dados = pd.DataFrame(itens)
    log = pd.DataFrame(erros, columns=["Origem", "Linha", "Tipo", "Erro"])
    if dados.empty:
        return None, log, f"{origem}: nenhuma CPU foi identificada."
    if higienizar:
        for coluna in ["CPU", "Codigo", "Und"]:
            dados[coluna] = dados[coluna].astype(str).str.strip().str.upper()
        for coluna in ["Descricao_CPU", "Descricao"]:
            dados[coluna] = dados[coluna].astype(str).str.strip()
    return dados, log, ""


def preparar_itens(df):
    """Preserva cada linha da CPU, inclusive códigos repetidos.

    A ocorrência sequencial impede que duas linhas iguais sejam somadas ou
    ocultadas antes da comparação. Quando há repetição, a primeira ocorrência
    é comparada com a primeira do outro arquivo, e assim sucessivamente.
    """
    saida = df.sort_values("Ordem").copy()
    saida["Ocorrencia"] = saida.groupby(["CPU", "Codigo"]).cumcount() + 1
    return saida


def dados_lado(df, lado):
    if df.empty:
        return pd.DataFrame(columns=["Ordem", "CPU", "Descricao_CPU", "Codigo", "Descricao", "Tipo", "Und", "Qtd", "Preco_Unitario"])
    sufixo = f"_{lado}"
    colunas = {"Ordem": f"Ordem{sufixo}", "CPU": "CPU", "Descricao_CPU": f"Descricao_CPU{sufixo}", "Codigo": "Codigo", "Descricao": f"Descricao{sufixo}", "Tipo": f"Tipo{sufixo}", "Und": f"Und{sufixo}", "Qtd": f"Qtd{sufixo}", "Preco_Unitario": f"Preco_Unitario{sufixo}"}
    return pd.DataFrame({novo: df[velho] for novo, velho in colunas.items()}).dropna(subset=["Codigo"])


def conciliar(base_raw, prop_raw):
    base, prop = preparar_itens(base_raw), preparar_itens(prop_raw)
    unido = base.merge(prop, on=["CPU", "Codigo", "Ocorrencia"], how="outer", suffixes=("_Base", "_Prop"), indicator=True)
    auditado = unido[unido["_merge"] == "both"].copy().rename(columns={
        "Descricao_CPU_Base": "Descricao_CPU", "Descricao_Base": "Descricao", "Ordem_Base": "Ordem",
        "Preco_Unitario_Base": "Preco_Base", "Preco_Unitario_Prop": "Preco_Prop",
    })
    for coluna in ["Qtd_Base", "Qtd_Prop", "Preco_Base", "Preco_Prop"]:
        auditado[coluna] = auditado[coluna].fillna(0.0)
    for coluna in ["Und_Base", "Und_Prop"]:
        auditado[coluna] = auditado[coluna].fillna("")
    auditado["Total_Base"] = auditado["Qtd_Base"] * auditado["Preco_Base"]
    auditado["Total_Prop"] = auditado["Qtd_Prop"] * auditado["Preco_Prop"]
    auditado["Delta_Qtd"] = auditado["Qtd_Prop"] - auditado["Qtd_Base"]
    auditado["Delta_Preco"] = auditado["Preco_Prop"] - auditado["Preco_Base"]
    auditado["Delta_Total"] = auditado["Total_Prop"] - auditado["Total_Base"]
    auditado["Var_Preco_%"] = np.where(auditado["Preco_Base"] != 0, auditado["Preco_Prop"] / auditado["Preco_Base"] - 1, 0.0)
    auditado["Var_Total_%"] = np.where(auditado["Total_Base"] != 0, auditado["Total_Prop"] / auditado["Total_Base"] - 1, 0.0)

    apenas_base = unido[unido["_merge"] == "left_only"].copy()
    apenas_prop = unido[unido["_merge"] == "right_only"].copy()
    codigos_prop, codigos_base = set(prop["Codigo"]), set(base["Codigo"])
    realocados = dados_lado(apenas_base[apenas_base["Codigo"].isin(codigos_prop)], "Base")
    omitidos = dados_lado(apenas_base[~apenas_base["Codigo"].isin(codigos_prop)], "Base")
    adicionados = dados_lado(apenas_prop[~apenas_prop["Codigo"].isin(codigos_base)], "Prop")
    return auditado, omitidos, adicionados, realocados


def hierarquia(df, auditoria=False):
    if df.empty:
        return pd.DataFrame()
    linhas, cpu_anterior = [], None
    for _, item in df.sort_values(["Ordem", "CPU", "Codigo"]).iterrows():
        if item["CPU"] != cpu_anterior:
            if cpu_anterior is not None:
                linhas.append({})
            if auditoria:
                linhas.append({"Código": item["CPU"], "Descrição": f"COMPOSIÇÃO: {item['Descricao_CPU']}", "Und_Base": "---", "Und_Prop": "---"})
            else:
                linhas.append({"Código": item["CPU"], "Descrição": f"COMPOSIÇÃO: {item['Descricao_CPU']}", "Unidade": "---"})
            cpu_anterior = item["CPU"]
        if auditoria:
            linhas.append({"Código": item["Codigo"], "Descrição": item["Descricao"], "Und_Base": item["Und_Base"], "Und_Prop": item["Und_Prop"], "Qtd_Base": item["Qtd_Base"], "Qtd_Prop": item["Qtd_Prop"], "Delta_Qtd": item["Delta_Qtd"], "Preco_Base": item["Preco_Base"], "Preco_Prop": item["Preco_Prop"], "Delta_Preco": item["Delta_Preco"], "Var_Preco_%": item["Var_Preco_%"], "Total_Base": item["Total_Base"], "Total_Prop": item["Total_Prop"], "Delta_Total": item["Delta_Total"], "Var_Total_%": item["Var_Total_%"]})
        else:
            linhas.append({"Código": item["Codigo"], "Descrição": item["Descricao"], "Unidade": item["Und"], "Quantidade": item["Qtd"], "Preço Unitário": item["Preco_Unitario"], "Total": item["Qtd"] * item["Preco_Unitario"], "Tipo": item["Tipo"]})
    return pd.DataFrame(linhas)


def estilizar(linha):
    if str(linha.get("Und_Base", "")) == "---":
        return ["background-color:#DBEAFE;font-weight:bold;color:#1E3A8A"] * len(linha)
    estilos = [""] * len(linha)
    for i, coluna in enumerate(linha.index):
        valor = linha[coluna]
        if coluna in {"Und_Base", "Und_Prop"} and linha.get("Und_Base") and linha.get("Und_Prop") and linha["Und_Base"] != linha["Und_Prop"]:
            estilos[i] = "background-color:#FEF08A;color:#713F12;font-weight:bold"
        elif coluna == "Delta_Qtd" and pd.notna(valor) and valor != 0:
            estilos[i] = "background-color:#E9D5FF;color:#6B21A8;font-weight:bold"
        elif coluna in {"Delta_Preco", "Delta_Total", "Var_Preco_%", "Var_Total_%"} and pd.notna(valor) and valor > 0:
            estilos[i] = "background-color:#FCA5A5;color:#7F1D1D;font-weight:bold"
        elif coluna in {"Var_Preco_%", "Var_Total_%"} and pd.notna(valor) and valor < st.session_state.get("limiar", -0.25):
            estilos[i] = "background-color:#FDBA74;color:#7C2D12;font-weight:bold"
    return estilos


def calcular_metricas(auditado, inconformidades, omitidos, realocados, limiar):
    total_base = float(auditado["Total_Base"].sum())
    total_prop = float(auditado["Total_Prop"].sum())
    total_itens = len(auditado)
    sobrepreco = float(auditado.loc[auditado["Delta_Total"] > 0, "Delta_Total"].sum())
    inexequivel = float(abs(auditado.loc[auditado["Var_Total_%"] < limiar, "Delta_Total"].sum()))
    return {
        "total_itens": total_itens,
        "total_base": total_base,
        "total_prop": total_prop,
        "variacao_geral": total_prop / total_base - 1 if total_base else 0.0,
        "conformidade": (total_itens - len(inconformidades)) / total_itens if total_itens else 0.0,
        "sobrepreco": sobrepreco,
        "inexequivel": inexequivel,
        "maior_desvio": float(auditado["Delta_Total"].max()) if total_itens else 0.0,
        "omitidos": len(omitidos),
        "realocados": len(realocados),
    }


def formatar_aba(ws, cabecalho):
    borda = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"), top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))
    for celula in ws[cabecalho]:
        celula.font, celula.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1E293B")
        celula.alignment, celula.border = Alignment(horizontal="center", vertical="center", wrap_text=True), borda
    mapa = {str(ws.cell(cabecalho, coluna).value or ""): coluna for coluna in range(1, ws.max_column + 1)}
    if cabecalho == 9:
        legendas = [
            ("A1", "FCA5A5", "🟥 Vermelho: sobrepreço."),
            ("A2", "E9D5FF", "🟪 Roxo: quantidade alterada."),
            ("A3", "FDBA74", "🟧 Laranja: desconto excessivo."),
            ("A4", "FEF08A", "🟨 Amarelo: unidade incompatível."),
            ("A5", "DBEAFE", "🟦 Azul: composição principal."),
        ]
        for endereco, cor, mensagem in legendas:
            ws[endereco] = mensagem
            ws[endereco].fill = PatternFill("solid", fgColor=cor)
            ws[endereco].font = Font(bold=True, size=9)
    for linha in ws.iter_rows(min_row=cabecalho + 1):
        if len(linha) > 1 and str(linha[1].value or "").startswith("COMPOSIÇÃO:"):
            for celula in linha:
                celula.fill, celula.font = PatternFill("solid", fgColor="DBEAFE"), Font(bold=True, color="1E3A8A")
        for celula in linha:
            celula.border = borda
        if cabecalho == 9 and not str(linha[1].value or "").startswith("COMPOSIÇÃO:"):
            def pintar(nome, cor):
                coluna = mapa.get(nome)
                if coluna:
                    ws.cell(linha[0].row, coluna).fill = PatternFill("solid", fgColor=cor)
            if mapa.get("Und_Base") and mapa.get("Und_Prop"):
                und_base, und_prop = ws.cell(linha[0].row, mapa["Und_Base"]).value, ws.cell(linha[0].row, mapa["Und_Prop"]).value
                if und_base and und_prop and und_base != und_prop:
                    pintar("Und_Base", "FEF08A"); pintar("Und_Prop", "FEF08A")
            if mapa.get("Delta_Qtd") and (ws.cell(linha[0].row, mapa["Delta_Qtd"]).value or 0) != 0: pintar("Delta_Qtd", "E9D5FF")
            for nome in ["Delta_Preco", "Delta_Total", "Var_Preco_%", "Var_Total_%"]:
                if mapa.get(nome) and (ws.cell(linha[0].row, mapa[nome]).value or 0) > 0: pintar(nome, "FCA5A5")
    for coluna in range(1, ws.max_column + 1):
        largura = max(len(str(ws.cell(linha, coluna).value or "")) for linha in range(1, ws.max_row + 1))
        ws.column_dimensions[get_column_letter(coluna)].width = min(max(largura + 2, 11), 65)
    ws.freeze_panes = f"A{cabecalho + 1}"


def gerar_excel(auditado, matriz, erros, omitidos, adicionados, realocados, base, prop, log, limiar):
def gerar_excel(auditado, matriz, erros, omitidos, adicionados, realocados, base, prop, log, limiar, metricas):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        matriz.to_excel(writer, index=False, sheet_name="Matriz Completa", startrow=8)
        erros.to_excel(writer, index=False, sheet_name="Inconformidades", startrow=8)
        hierarquia(omitidos).to_excel(writer, index=False, sheet_name="Omitidos na Proposta", startrow=1)
        hierarquia(adicionados).to_excel(writer, index=False, sheet_name="Nao Encontrados na Base", startrow=1)
        hierarquia(realocados).to_excel(writer, index=False, sheet_name=SHEET_REALOCADOS, startrow=1)
        log.to_excel(writer, index=False, sheet_name="Log de Parsing", startrow=1)
        hierarquia(base).to_excel(writer, index=False, sheet_name="DB Base", startrow=1)
        hierarquia(prop).to_excel(writer, index=False, sheet_name="DB Proposta", startrow=1)
        wb = writer.book
        painel = wb.create_sheet("Dashboard KPI", 0)
        painel.sheet_view.showGridLines = False
        painel.merge_cells("B2:J3")
        painel["B2"] = "PAINEL ANALITICO DE CONFORMIDADE CONTRATUAL"
        painel["B2"].font, painel["B2"].fill = Font(size=16, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="0F172A")
        painel["B2"].alignment = Alignment(horizontal="center", vertical="center")
        cards = [
            ("ITENS AUDITADOS", metricas["total_itens"], "int"),
            ("SALDO DO ORÇAMENTO", metricas["total_prop"], "money"),
            ("TAXA DE CONFORMIDADE", metricas["conformidade"], "percent"),
            ("RISCO SOBREPREÇO", metricas["sobrepreco"], "money"),
            ("DESCONTO OCULTO", metricas["inexequivel"], "money"),
            ("MAIOR DESVIO ÚNICO", metricas["maior_desvio"], "money"),
        metricas = [("ITENS AUDITADOS", len(auditado)), ("OMITIDOS REAIS", len(omitidos)), ("REALOCADOS", len(realocados)), ("LIMIAR DE DESCONTO", limiar)]
        for indice, (titulo, valor) in enumerate(metricas):
            coluna, linha = 2 + (indice % 2) * 4, 5 + (indice // 2) * 4
            painel.merge_cells(start_row=linha, start_column=coluna, end_row=linha, end_column=coluna + 2)
            painel.merge_cells(start_row=linha + 1, start_column=coluna, end_row=linha + 2, end_column=coluna + 2)
        ]
        for indice, (titulo, valor, formato) in enumerate(cards):
            coluna, linha = 2 + (indice % 3) * 3, 5 + (indice // 3) * 4
            painel.merge_cells(start_row=linha, start_column=coluna, end_row=linha, end_column=coluna + 1)
            painel.merge_cells(start_row=linha + 1, start_column=coluna, end_row=linha + 2, end_column=coluna + 1)
            painel.cell(linha, coluna, titulo).fill = PatternFill("solid", fgColor="334155")
            painel.cell(linha, coluna).font = Font(bold=True, color="FFFFFF")
            painel.cell(linha, coluna).alignment = Alignment(horizontal="center")
            painel.cell(linha + 1, coluna, valor).font = Font(size=15, bold=True)
            painel.cell(linha + 1, coluna).alignment = Alignment(horizontal="center", vertical="center")
            if indice == 3:
                painel.cell(linha + 1, coluna).number_format = "0%"
            if formato == "percent": painel.cell(linha + 1, coluna).number_format = "0.0%"
            elif formato == "money": painel.cell(linha + 1, coluna).number_format = '"R$" #,##0.00'
        for letra in "BCDEFGHIJ": painel.column_dimensions[letra].width = 18
        for nome in wb.sheetnames:
            if nome != "Dashboard KPI": formatar_aba(wb[nome], 9 if nome in {"Matriz Completa", "Inconformidades"} else 2)
    return output.getvalue()


def main():
    st.set_page_config(page_title="Auditoria PRO", layout="wide", page_icon="🛡️")
    st.session_state.setdefault("limiar", -0.25)
    with st.sidebar:
        st.subheader("⚙️ Parâmetros de Auditoria")
        st.session_state.limiar = st.slider("Limiar de Desconto (Inexequibilidade)", -50, -5, -25, 1) / 100
        higienizar = st.checkbox("🧹 Higienizar dados", value=True)
        st.info("CPUs são lidas com insumos e composições auxiliares. Itens em outra CPU são realocados, não omitidos.")
        st.success(f"📌 Limiar configurado: {st.session_state.limiar:.0%}")
        st.divider()
        st.subheader("📌 Legenda de Auditoria")
        st.markdown("""
        <div style="padding:8px;margin-bottom:5px;border-radius:5px;background:#fca5a5;color:#7f1d1d"><b>🟥 Vermelho</b><br>Sobrepreço na proposta.</div>
        <div style="padding:8px;margin-bottom:5px;border-radius:5px;background:#e9d5ff;color:#6b21a8"><b>🟪 Roxo</b><br>Quantidade alterada.</div>
        <div style="padding:8px;margin-bottom:5px;border-radius:5px;background:#fdba74;color:#7c2d12"><b>🟧 Laranja</b><br>Desconto excessivo.</div>
        <div style="padding:8px;margin-bottom:5px;border-radius:5px;background:#fef08a;color:#713f12"><b>🟨 Amarelo</b><br>Unidade incompatível.</div>
        <div style="padding:8px;margin-bottom:5px;border-radius:5px;background:#dbeafe;color:#1e3a8a"><b>🟦 Azul</b><br>Composição principal.</div>
        """, unsafe_allow_html=True)
        st.caption("Itens em outra CPU são realocados, não omitidos.")
    st.title("🛡️ Auditoria de Orçamentos PRO")
    st.markdown("Validação paramétrica de CPUs do OrçaFascio.")
    col1, col2 = st.columns(2)
    with col1: arquivo_base = st.file_uploader("1. Base de Referência", type=["xlsx", "xls"])
    with col2: arquivo_prop = st.file_uploader("2. Proposta da Empreiteira", type=["xlsx", "xls"])
    if not (arquivo_base and arquivo_prop): return
    with st.spinner("Lendo e conciliando as CPUs..."):
        base, log_base, erro_base = ler_orcafascio(arquivo_base.getvalue(), "Base", higienizar)
        prop, log_prop, erro_prop = ler_orcafascio(arquivo_prop.getvalue(), "Proposta", higienizar)
    if erro_base or erro_prop:
        st.error(erro_base or erro_prop); return
    auditado, omitidos, adicionados, realocados = conciliar(base, prop)
    filtro_preco = (auditado["Delta_Preco"] > 0) | (auditado["Delta_Total"] > 0)
    filtro_qtd, filtro_und = auditado["Delta_Qtd"] != 0, auditado["Und_Base"] != auditado["Und_Prop"]
    filtro_inex = (auditado["Var_Preco_%"] < st.session_state.limiar) | (auditado["Var_Total_%"] < st.session_state.limiar)
    inconformidades = auditado[filtro_preco | filtro_qtd | filtro_und | filtro_inex].copy()
    matriz, tabela_erros = hierarquia(auditado, True), hierarquia(inconformidades, True)
    log = pd.concat([log_base, log_prop], ignore_index=True)
    metricas = calcular_metricas(auditado, inconformidades, omitidos, realocados, st.session_state.limiar)
    excel = gerar_excel(auditado, matriz, tabela_erros, omitidos, adicionados, realocados, base, prop, log, st.session_state.limiar)
    excel = gerar_excel(auditado, matriz, tabela_erros, omitidos, adicionados, realocados, base, prop, log, st.session_state.limiar, metricas)
    formato = {"Qtd_Base": "{:.4f}", "Qtd_Prop": "{:.4f}", "Delta_Qtd": "{:.4f}", "Preco_Base": "R$ {:.2f}", "Preco_Prop": "R$ {:.2f}", "Delta_Preco": "R$ {:.2f}", "Var_Preco_%": "{:.2%}", "Total_Base": "R$ {:.2f}", "Total_Prop": "R$ {:.2f}", "Delta_Total": "R$ {:.2f}", "Var_Total_%": "{:.2%}"}
    tabs = st.tabs(["📊 Dashboard KPI", "📋 Matriz Completa", "🚨 Inconformidades", "📍 Omitidos Reais", "🔀 Realocados", "📍 Adicionados", "📝 Log", "🗄️ DB Base", "🗄️ DB Proposta"])
    with tabs[0]:
        a, b, c = st.columns(3)
        a.metric("Itens auditados", len(auditado), f"{len(inconformidades)} divergências")
        b.metric("Omitidos reais", len(omitidos), "código ausente na proposta")
        c.metric("Realocados", len(realocados), "código em outra CPU")
        a.metric("📌 Volume de Itens Auditados", f"{metricas['total_itens']:,}", f"{len(inconformidades)} desvios sinalizados")
        b.metric("💰 Saldo Global do Orçamento", f"R$ {metricas['total_prop']:,.2f}", f"Variação: {metricas['variacao_geral']:+.2%}", delta_color="inverse")
        c.metric("✅ Índice de Acerto Paramétrico", f"{metricas['conformidade']:.1%}", "Meta aceitável: > 95%")
        d, e, f = st.columns(3)
        d.metric("🚨 Exposição a Sobrepreço", f"R$ {metricas['sobrepreco']:,.2f}", delta_color="inverse")
        e.metric("📉 Desconto Ofertado Oculto", f"R$ {metricas['inexequivel']:,.2f}")
        f.metric("🎯 Maior Desvio Único", f"R$ {metricas['maior_desvio']:,.2f}")
        st.divider()
        grafico1, grafico2 = st.columns(2)
        with grafico1:
            st.markdown("##### 🔢 Ocorrências por Tipologia")
            contagens = pd.Series({"Sobrepreço": int(filtro_preco.sum()), "Qtd. alterada": int(filtro_qtd.sum()), "Unidade incompatível": int(filtro_und.sum()), "Inexequível": int(filtro_inex.sum()), "Omitidos reais": len(omitidos), "Realocados": len(realocados)})
            st.bar_chart(contagens, height=280)
        with grafico2:
            st.markdown("##### 💸 Impacto Financeiro Líquido")
            impactos = pd.Series({"Sobrepreço": metricas["sobrepreco"], "Descontos extremos": metricas["inexequivel"], "Variação global": metricas["total_prop"] - metricas["total_base"]})
            st.bar_chart(impactos, height=280)
        st.divider()
        esquerda, direita = st.columns(2)
        with esquerda:
            st.markdown("##### 🔺 Top 5 impactos de sobrepreço")
            top_sobre = auditado[auditado["Delta_Total"] > 0].nlargest(5, "Delta_Total")[["Codigo", "Descricao", "Delta_Total", "Var_Total_%"]]
            if top_sobre.empty:
                st.success("Nenhum sobrepreço mapeado.")
            else:
                st.dataframe(top_sobre.style.format({"Delta_Total": "R$ {:.2f}", "Var_Total_%": "{:+.2%}"}), hide_index=True, use_container_width=True)
        with direita:
            st.markdown("##### 🔻 Top 5 riscos de inexequibilidade")
            top_inex = auditado[filtro_inex].nsmallest(5, "Delta_Total")[["Codigo", "Descricao", "Delta_Total", "Var_Total_%"]]
            if top_inex.empty:
                st.info("Nenhuma anomalia de desconto extremo encontrada.")
            else:
                st.dataframe(top_inex.style.format({"Delta_Total": "R$ {:.2f}", "Var_Total_%": "{:.2%}"}), hide_index=True, use_container_width=True)
        st.divider()
        st.download_button("📥 Baixar Laudo de Auditoria (.XLSX)", excel, "Laudo_Auditoria_PRO_Consolidado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tabs[1]: st.dataframe(matriz.style.format(formato, na_rep="").apply(estilizar, axis=1), height=600, use_container_width=True)
    with tabs[2]: st.dataframe(tabela_erros.style.format(formato, na_rep="").apply(estilizar, axis=1), height=600, use_container_width=True)
    for aba, dados, mensagem in [(tabs[3], omitidos, "Nenhum item realmente omitido."), (tabs[4], realocados, "Nenhum item realocado."), (tabs[5], adicionados, "Nenhum item adicionado.")]:
        with aba:
            visual = hierarquia(dados)
            if visual.empty:
                st.success(f"✅ {mensagem}")
            else:
                st.dataframe(visual, height=600, use_container_width=True)
    with tabs[6]:
        if log.empty:
            st.success("✅ Sem erros de leitura.")
        else:
            st.dataframe(log, use_container_width=True)
    with tabs[7]: st.dataframe(hierarquia(base), height=600, use_container_width=True)
    with tabs[8]: st.dataframe(hierarquia(prop), height=600, use_container_width=True)


if __name__ == "__main__":
    main()
