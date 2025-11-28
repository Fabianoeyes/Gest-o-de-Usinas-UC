import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# =========================
# CONFIGURAÇÕES GERAIS
# =========================
st.set_page_config(
    page_title="Gestão de Usinas e UCs",
    page_icon="⚡",
    layout="wide",
)

EXCEL_DEFAULT_PATH = Path("Gestao_de_Usinas_e_UCs.xlsx")


def localizar_arquivo_excel(default_path: Path) -> Path | None:
    """Tenta localizar o arquivo Excel no diretório do app.

    1. Usa o caminho padrão, se existir.
    2. Se não existir, procura qualquer arquivo ``*.xlsx`` e pega o mais recente.
    """

    if default_path.exists():
        return default_path

    candidatos = sorted(Path(".").glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if candidatos:
        return candidatos[0]

    return None

@st.cache_data
def carregar_planilhas(path: Path):
    """
    Lê todas as abas do Excel em um dict: {nome_aba: DataFrame}.
    """
    xls = pd.read_excel(path, sheet_name=None, engine="openpyxl")
    # Normaliza nomes das abas (remove espaços nas pontas)
    sheets = {nome.strip(): df for nome, df in xls.items()}
    return sheets

def salvar_planilha(path: Path, sheets_dict: dict):
    """
    Salva os DataFrames de volta para um arquivo Excel.
    (Útil se você quiser persistir edições feitas no portal.)
    """
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for nome, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=nome, index=False)

# =========================
# FUNÇÕES DE CÁLCULO (ADAPTAÇÃO DAS FÓRMULAS DO EXCEL)
# =========================

def calcular_resumo_geral(sheets: dict):
    """
    Aqui você replica, em Python, os principais indicadores
    que hoje são calculados na planilha (soma, média, payback, etc).

    Você deve adaptar para a estrutura real dos seus dados.
    Abaixo é só um exemplo genérico.
    """
    resumo = {}

    # EXEMPLO: se existir uma aba "Usinas"
    if "Usinas" in sheets:
        df_usinas = sheets["Usinas"]

        # Exemplo de colunas que você pode ter (ajuste os nomes):
        # - "Usina"
        # - "Potencia_kW"
        # - "Geracao_mensal_kWh"
        # - "Receita_mensal_R$"

        col_potencia = [c for c in df_usinas.columns if "Potencia" in c or "potência" in c or "kW" in c]
        col_receita = [c for c in df_usinas.columns if "Receita" in c or "R$" in c]

        if col_potencia:
            resumo["potencia_total_kw"] = float(df_usinas[col_potencia[0]].fillna(0).sum())
        if col_receita:
            resumo["receita_mensal_total"] = float(df_usinas[col_receita[0]].fillna(0).sum())

        resumo["qtd_usinas"] = len(df_usinas)

    # EXEMPLO: se existir uma aba "UCs"
    if "UCs" in sheets:
        df_ucs = sheets["UCs"]
        resumo["qtd_ucs"] = len(df_ucs)

    return resumo

def calcular_metricas_usinas(df_usinas: pd.DataFrame):
    """
    Aqui você traz as mesmas fórmulas que o Excel usa
    para métricas específicas de usinas.
    """
    metricas = {}

    if df_usinas.empty:
        return metricas

    # Exemplos genéricos – ajuste com base nas colunas reais
    if "Potencia_kW" in df_usinas.columns:
        metricas["Potência total (kW)"] = df_usinas["Potencia_kW"].sum()

    if "Geracao_mensal_kWh" in df_usinas.columns:
        metricas["Geração mensal total (kWh)"] = df_usinas["Geracao_mensal_kWh"].sum()

    if "Receita_mensal_R$" in df_usinas.columns:
        metricas["Receita mensal total (R$)"] = df_usinas["Receita_mensal_R$"].sum()

    return metricas

def calcular_metricas_ucs(df_ucs: pd.DataFrame):
    """
    Métricas de UCs (unidades consumidoras).
    Replicar fórmulas da aba correspondente.
    """
    metricas = {}

    if df_ucs.empty:
        return metricas

    # Exemplo genérico:
    if "Consumo_mensal_kWh" in df_ucs.columns:
        metricas["Consumo mensal total (kWh)"] = df_ucs["Consumo_mensal_kWh"].sum()

    if "Economia_mensal_R$" in df_ucs.columns:
        metricas["Economia mensal total (R$)"] = df_ucs["Economia_mensal_R$"].sum()

    return metricas

# =========================
# LAYOUT DO PORTAL
# =========================

st.sidebar.title("Gestão de Usinas & UCs")

excel_path = localizar_arquivo_excel(EXCEL_DEFAULT_PATH)
if not excel_path:
    st.error(
        "Arquivo Excel não encontrado. Coloque um arquivo .xlsx na mesma pasta "
        "do aplicativo ou ajuste o nome em `EXCEL_DEFAULT_PATH`."
    )
    st.stop()

sheets = carregar_planilhas(excel_path)
st.sidebar.success(f"Arquivo carregado: {excel_path.name}")

if excel_path.name != EXCEL_DEFAULT_PATH.name:
    st.sidebar.info(
        "O arquivo padrão não foi encontrado; usando o arquivo Excel mais recente "
        "localizado na pasta."
    )
abas_disponiveis = list(sheets.keys())

pagina = st.sidebar.radio(
    "Navegação",
    [
        "📊 Dashboard geral",
        "📈 Dashboard operacional",
        "⚡ Usinas",
        "🏠 UCs",
        "📑 Editor de planilhas (avançado)",
    ],
)

# =========================
# DASHBOARD GERAL
# =========================

if pagina == "📊 Dashboard geral":
    st.title("📊 Dashboard Geral – Gestão de Usinas & UCs")

    resumo = calcular_resumo_geral(sheets)

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "Quantidade de Usinas",
            resumo.get("qtd_usinas", 0)
        )

    with col2:
        st.metric(
            "Quantidade de UCs",
            resumo.get("qtd_ucs", 0)
        )

    with col3:
        pot_kw = resumo.get("potencia_total_kw")
        st.metric(
            "Potência total instalada (kW)",
            f"{pot_kw:,.0f}" if pot_kw is not None else "-"
        )

    with col4:
        rec = resumo.get("receita_mensal_total")
        st.metric(
            "Receita mensal total (R$)",
            f"R$ {rec:,.0f}".replace(",", ".") if rec is not None else "-"
        )

    st.markdown("---")

    # Exemplo de gráfico de usinas
    if "Usinas" in sheets:
        st.subheader("Distribuição de potência por usina (exemplo)")
        df_usinas = sheets["Usinas"].copy()

        # Tenta achar colunas de nome da usina e potência
        col_nome = None
        for c in df_usinas.columns:
            if "usina" in c.lower() or "nome" in c.lower():
                col_nome = c
                break

        col_pot = None
        for c in df_usinas.columns:
            if "potencia" in c.lower() or "kw" in c.lower():
                col_pot = c
                break

        if col_nome and col_pot:
            df_chart = df_usinas[[col_nome, col_pot]].dropna()
            df_chart = df_chart.groupby(col_nome)[col_pot].sum().reset_index()
            df_chart = df_chart.set_index(col_nome)
            st.bar_chart(df_chart)

    st.info(
        "OBS: As métricas e gráficos podem (e devem) ser ajustados para "
        "reproduzir exatamente os cálculos da sua planilha."
    )

# =========================
# DASHBOARD OPERACIONAL
# =========================

elif pagina == "📈 Dashboard operacional":
    st.title("📈 Dashboard Operacional")

    nome_aba = "Dashboard Operacional"
    if nome_aba not in sheets:
        st.error("Aba 'Dashboard Operacional' não encontrada no Excel. Ajuste o nome no código.")
        st.stop()

    df_raw = sheets[nome_aba].copy()

    # A aba é composta por tabelas posicionadas (título em 144 e dados a partir de 145).
    # Para não depender das posições exatas, detectamos a linha onde aparecem as distribuidoras
    # (ex.: ENEL CE, EQUATORIAL PI, etc.) e usamos essa linha como cabeçalho dinâmico.
    header_idx = None
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains("ENEL CE", na=False).any():
            header_idx = idx
            break

    if header_idx is None:
        st.warning(
            "Não foi possível localizar os rótulos de distribuidoras na aba. "
            "Confira se a planilha mantém o texto 'ENEL CE' ou atualize o parser."
        )
        st.dataframe(df_raw)
        st.stop()

    header_row = df_raw.loc[header_idx]
    primeiro_valido = header_row.first_valid_index()

    # A planilha usa muitas colunas "Unnamed"; precisamos trabalhar por posição
    # para evitar erros do tipo "slice indices must be integers".
    pos_inicio = df_raw.columns.get_loc(primeiro_valido)
    colunas_validas = [
        (df_raw.columns[i], header_row.iloc[i])
        for i in range(pos_inicio, len(header_row))
        if pd.notna(header_row.iloc[i])
    ]

    dados = df_raw.loc[
        header_idx + 1 :,
        [header_row.index[1], *[col for col, _ in colunas_validas]],
    ].copy()
    dados.columns = ["Indicador", *[nome for _, nome in colunas_validas]]

    dados = dados.dropna(how="all")
    dados_indicadores = dados.dropna(axis=0, how="all").copy()
    dados_indicadores["Indicador"] = dados_indicadores["Indicador"].astype(str).str.strip()

    # Tabela com indicadores consolidados (MWh / R$)
    st.subheader("Indicadores por distribuidora")
    st.dataframe(dados_indicadores, use_container_width=True)

    # Destaques numéricos: usamos as linhas com valores numéricos nas primeiras colunas.
    destaques = dados_indicadores.set_index("Indicador").select_dtypes(include=["number", "float", "int"]).iloc[:4]

    if not destaques.empty:
        st.markdown("### Geração e venda (MWh e R$)")
        st.bar_chart(destaques)
    else:
        st.info("Nenhum dado numérico encontrado para montar os gráficos.")

    # Gráficos de pizza para replicar a planilha
    def plot_pizza(indicador: str, titulo: str):
        base = dados_indicadores.set_index("Indicador")
        if indicador not in base.index:
            st.info(f"Indicador '{indicador}' não encontrado para montar o gráfico.")
            return

        serie = pd.to_numeric(base.loc[indicador], errors="coerce").dropna()
        if serie.empty:
            st.info(f"Sem valores numéricos para '{indicador}'.")
            return

        fig, ax = plt.subplots()
        ax.pie(serie.values, labels=serie.index, autopct="%1.1f%%", startangle=90)
        ax.set_title(titulo)
        ax.axis("equal")
        st.pyplot(fig)

    st.markdown("### Distribuição por distribuidora (gráficos de pizza)")
    col_a, col_b = st.columns(2)

    with col_a:
        plot_pizza("ENERGIA MENSAL VENDIDA (MWh)", "Energia vendida (MWh)")

    with col_b:
        plot_pizza("ENERGIA MENSAL FALTANTE (MWh)", "Energia faltante (MWh)")

# =========================
# PÁGINA DE USINAS
# =========================

elif pagina == "⚡ Usinas":
    st.title("⚡ Gestão de Usinas")

    if "Usinas" not in sheets:
        st.error("Aba 'Usinas' não encontrada no Excel. Ajuste o nome no código.")
        st.stop()

    df_usinas = sheets["Usinas"].copy()

    st.subheader("Tabela de usinas (editável)")
    edited_df = st.data_editor(
        df_usinas,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_usinas"
    )

    metricas = calcular_metricas_usinas(edited_df)

    st.markdown("### Indicadores principais")
    cols = st.columns(len(metricas) or 1)
    for (nome, valor), col in zip(metricas.items(), cols):
        with col:
            if isinstance(valor, (int, float)):
                col.metric(nome, f"{valor:,.2f}".replace(",", "."))
            else:
                col.metric(nome, str(valor))

    # Aqui você pode adicionar gráficos específicos de usinas
    st.markdown("---")
    st.markdown("### Gráficos (personalize conforme suas colunas)")

    if len(edited_df.columns) >= 2:
        st.bar_chart(edited_df.select_dtypes("number"))

    st.warning(
        "Importante: neste exemplo as alterações são apenas em memória. "
        "Se quiser salvar de volta no Excel, é só chamar `salvar_planilha` "
        "com o dicionário de abas atualizado."
    )

# =========================
# PÁGINA DE UCs
# =========================

elif pagina == "🏠 UCs":
    st.title("🏠 Gestão de UCs (Unidades Consumidoras)")

    if "UCs" not in sheets:
        st.error("Aba 'UCs' não encontrada no Excel. Ajuste o nome no código.")
        st.stop()

    df_ucs = sheets["UCs"].copy()

    st.subheader("Tabela de UCs (editável)")
    edited_df = st.data_editor(
        df_ucs,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_ucs"
    )

    metricas = calcular_metricas_ucs(edited_df)

    st.markdown("### Indicadores principais")
    cols = st.columns(len(metricas) or 1)
    for (nome, valor), col in zip(metricas.items(), cols):
        with col:
            if isinstance(valor, (int, float)):
                col.metric(nome, f"{valor:,.2f}".replace(",", "."))
            else:
                col.metric(nome, str(valor))

    st.markdown("---")
    if len(edited_df.columns) >= 2:
        st.line_chart(edited_df.select_dtypes("number"))

# =========================
# EDITOR GENÉRICO
# =========================

elif pagina == "📑 Editor de planilhas (avançado)":
    st.title("📑 Editor de Planilhas – modo avançado")

    aba_escolhida = st.selectbox("Selecione a aba para editar", options=abas_disponiveis)
    df_sel = sheets[aba_escolhida].copy()

    st.subheader(f"Aba: {aba_escolhida}")
    edited_df = st.data_editor(
        df_sel,
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_{aba_escolhida}"
    )

    st.info(
        "Se quiser persistir as alterações no Excel, é possível adaptar o código "
        "para chamar `salvar_planilha` com o dicionário de abas atualizado."
    )
