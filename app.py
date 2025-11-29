import streamlit as st
import pandas as pd
from pathlib import Path

# =========================
# CONFIGURAÇÕES GERAIS
# =========================

st.set_page_config(
    page_title="Gestão de Usinas e UCs",
    page_icon="⚡",
    layout="wide",
)

EXCEL_PATH = Path("Gestao_de_Usinas_e_UCs.xlsx")

# =========================
# FUNÇÕES AUXILIARES
# =========================

@st.cache_data
def carregar_planilhas(path: Path) -> dict:
    """
    Lê todas as abas do Excel em um dict: {nome_aba: DataFrame}.
    """
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name, header=None)
        sheets[name.strip()] = df
    return sheets


def df_com_cabecalho(df: pd.DataFrame, linha_cabecalho: int = 0) -> pd.DataFrame:
    """
    Converte uma planilha no formato "linha de cabeçalho + dados"
    para um DataFrame bonitinho:
      - Usa a linha `linha_cabecalho` como header
      - Remove colunas vazias
      - Remove linhas completamente vazias
    """
    if df.empty:
        return df

    # garante que a linha de cabeçalho exista
    linha_cabecalho = min(max(linha_cabecalho, 0), len(df) - 1)

    header = df.iloc[linha_cabecalho].fillna("")
    dados = df.iloc[linha_cabecalho + 1 :].copy()
    dados.columns = header

    # Remove colunas "vazias" (sem nome)
    dados = dados.loc[:, dados.columns.notnull()]
    dados = dados.loc[:, dados.columns.astype(str).str.strip() != ""]
    # Remove linhas totalmente vazias
    dados = dados.dropna(how="all")

    return dados


def calcular_resumo_geral(sheets: dict) -> dict:
    """
    Calcula alguns indicadores gerais em cima das abas principais.
    Aqui você pode replicar, aos poucos, as fórmulas do Excel
    (Quadro Resumo, Dashboards, etc.).
    """
    resumo = {}

    # Exemplo: pegar informação da aba "Informações Usinas"
    nome_aba_usinas = "Informações Usinas"
    if nome_aba_usinas in sheets:
        df_raw = sheets[nome_aba_usinas]
        df_usinas = df_com_cabecalho(df_raw, linha_cabecalho=1)

        # tentar detectar colunas importantes pela descrição
        col_usina = None
        for c in df_usinas.columns:
            if "Usina" in str(c):
                col_usina = c
                break

        resumo["qtd_usinas"] = df_usinas[col_usina].nunique() if col_usina else len(df_usinas)

        # potência instalada (exemplo: procura algo com "Capacidade (MW CA)" ou "kW")
        col_potencia = None
        for c in df_usinas.columns:
            if "Capacidade" in str(c) and ("MW" in str(c) or "kW" in str(c)):
                col_potencia = c
                break

        if col_potencia:
            resumo["potencia_total"] = float(df_usinas[col_potencia].fillna(0).sum())

    # Exemplo: dados de clientes
    nome_aba_clientes = "Base SIGH - Clientes"
    if nome_aba_clientes in sheets:
        df_raw_cli = sheets[nome_aba_clientes]
        # Aqui não sei exatamente qual linha é cabeçalho;
        # ajuste se necessário (0, 1, 2...).
        df_cli = df_com_cabecalho(df_raw_cli, linha_cabecalho=0)
        resumo["qtd_clientes"] = len(df_cli)

    return resumo


def calcular_metricas_usinas(df_usinas: pd.DataFrame) -> dict:
    """
    Métricas específicas de usinas.
    Aqui você pode copiar fórmulas da planilha (potência total, geração, receita, etc).
    """
    metricas = {}
    if df_usinas.empty:
        return metricas

    # Exemplos genéricos: ajuste nomes das colunas conforme a sua planilha.
    for c in df_usinas.columns:
        if "Capacidade (MW CA" in str(c) or "Capacidade (MW CA)" in str(c):
            metricas["Potência total (MW CA)"] = df_usinas[c].fillna(0).sum()
        if "Capacidade (MWp" in str(c) or "Capacidade (MWp)" in str(c):
            metricas["Potência total (MWp)"] = df_usinas[c].fillna(0).sum()
        if "Tarifa Gerador" in str(c):
            metricas["Tarifa média gerador (R$/MWh)"] = df_usinas[c].fillna(0).mean()

    return metricas


def calcular_metricas_quadro_resumo(df_qr: pd.DataFrame) -> pd.DataFrame:
    """
    Exemplo de transformação do 'Quadro Resumo' em formato tabular
    (Distribuidora x Energia x Receita etc.).
    Isso depende muito da estrutura real da sua aba.
    Aqui eu apenas limpo e mantenho como tabela para visualização.
    """
    df = df_qr.copy()
    df = df.dropna(how="all")
    df.columns = [f"Col_{i}" for i in range(len(df.columns))]
    return df


# =========================
# CARREGAMENTO DA PLANILHA
# =========================

if not EXCEL_PATH.exists():
    st.error(f"Arquivo Excel não encontrado: {EXCEL_PATH}")
    st.stop()

sheets_raw = carregar_planilhas(EXCEL_PATH)
abas_disponiveis = list(sheets_raw.keys())

# =========================
# MENU LATERAL
# =========================

st.sidebar.title("Gestão de Usinas & UCs")

pagina = st.sidebar.radio(
    "Navegação",
    [
        "📊 Dashboard geral",
        "⚡ Informações das Usinas",
        "📋 Quadro Resumo",
        "📋 Quadro Resumo - Usinas Ativas",
        "📑 Visualizar qualquer aba (avançado)",
    ],
)

# =========================
# PÁGINA: DASHBOARD GERAL
# =========================

if pagina == "📊 Dashboard geral":
    st.title("📊 Dashboard Geral – Gestão de Usinas & UCs")

    resumo = calcular_resumo_geral(sheets_raw)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Quantidade de usinas", resumo.get("qtd_usinas", 0))

    with col2:
        pot = resumo.get("potencia_total")
        st.metric(
            "Potência total (soma das usinas)",
            f"{pot:,.2f}" if pot is not None else "-",
        )

    with col3:
        st.metric("Quantidade de clientes (Base SIGH)", resumo.get("qtd_clientes", 0))

    st.markdown("---")
    st.write(
        "Este painel é um ponto de partida. Você pode ir refinando as fórmulas "
        "em `calcular_resumo_geral` para espelhar exatamente o que o Excel faz "
        "nos dashboards `Dashboard Operacional` e `Dashboard Financeiro`."
    )

# =========================
# PÁGINA: INFORMAÇÕES DAS USINAS
# =========================

elif pagina == "⚡ Informações das Usinas":
    st.title("⚡ Informações das Usinas")

    nome_aba_usinas = "Informações Usinas"
    if nome_aba_usinas not in sheets_raw:
        st.error(f"Aba '{nome_aba_usinas}' não encontrada no Excel.")
        st.stop()

    df_usinas_raw = sheets_raw[nome_aba_usinas]
    # Pela análise da planilha, o cabeçalho começa na linha 1 (index = 1)
    df_usinas = df_com_cabecalho(df_usinas_raw, linha_cabecalho=1)

    st.subheader("Tabela de usinas (editável)")
    df_edit = st.data_editor(
        df_usinas,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_usinas",
    )

    st.markdown("### Indicadores das usinas")
    metricas = calcular_metricas_usinas(df_edit)
    if metricas:
        cols = st.columns(len(metricas))
        for (nome, valor), col in zip(metricas.items(), cols):
            with col:
                if isinstance(valor, (int, float)):
                    col.metric(nome, f"{valor:,.2f}")
                else:
                    col.metric(nome, str(valor))
    else:
        st.info("Ajuste a função `calcular_metricas_usinas` para usar as colunas da sua planilha.")

    st.markdown("---")
    st.markdown("### Gráficos numéricos (todas as colunas numéricas)")
    df_numerico = df_edit.select_dtypes("number")
    if not df_numerico.empty:
        st.bar_chart(df_numerico)
    else:
        st.info("Nenhuma coluna numérica foi identificada automaticamente.")

# =========================
# PÁGINA: QUADRO RESUMO
# =========================

elif pagina == "📋 Quadro Resumo":
    st.title("📋 Quadro Resumo")

    nome_aba_qr = "Quadro Resumo "
    if nome_aba_qr not in sheets_raw:
        st.error(f"Aba '{nome_aba_qr}' não encontrada no Excel.")
        st.stop()

    df_qr_raw = sheets_raw[nome_aba_qr]
    df_qr = calcular_metricas_quadro_resumo(df_qr_raw)

    st.subheader("Visão tabular do Quadro Resumo (limpo)")
    st.dataframe(df_qr, use_container_width=True)

    st.info(
        "Se o Quadro Resumo tiver colunas específicas de energia, receita, "
        "inadimplência etc., você pode criar funções de cálculo em Python "
        "para gerar gráficos e cards aqui."
    )

# =========================
# PÁGINA: QUADRO RESUMO - USINAS ATIVAS
# =========================

elif pagina == "📋 Quadro Resumo - Usinas Ativas":
    st.title("📋 Quadro Resumo - Usinas Ativas")

    nome_aba_qr_ativas = "Quadro Resumo - Usina Ativas"
    if nome_aba_qr_ativas not in sheets_raw:
        st.error(f"Aba '{nome_aba_qr_ativas}' não encontrada no Excel.")
        st.stop()

    df_qr_ativas_raw = sheets_raw[nome_aba_qr_ativas]
    df_qr_ativas = calcular_metricas_quadro_resumo(df_qr_ativas_raw)

    st.subheader("Tabela – Usinas ativas")
    st.dataframe(df_qr_ativas, use_container_width=True)

    st.markdown("---")
    st.markdown("### Gráfico rápido (se houver números)")

    df_num = df_qr_ativas.select_dtypes("number")
    if not df_num.empty:
        st.line_chart(df_num)
    else:
        st.info("Nenhuma coluna numérica identificada automaticamente.")

# =========================
# PÁGINA: VISUALIZAR QUALQUER ABA
# =========================

elif pagina == "📑 Visualizar qualquer aba (avançado)":
    st.title("📑 Visualizador de abas (modo avançado)")

    aba_escolhida = st.selectbox("Selecione a aba", options=abas_disponiveis)
    df_raw = sheets_raw[aba_escolhida]

    st.markdown(f"### Aba: `{aba_escolhida}` (dados brutos)")
    st.dataframe(df_raw, use_container_width=True)

    st.markdown("---")
    st.markdown("### Tentar aplicar linha de cabeçalho automaticamente")

    linha_header = st.number_input(
        "Linha de cabeçalho (0 = primeira linha)",
        min_value=0,
        max_value=max(len(df_raw) - 1, 0),
        value=0,
    )

    df_header = df_com_cabecalho(df_raw, linha_cabecalho=int(linha_header))
    st.dataframe(df_header, use_container_width=True)

    st.info(
        "Use esta aba para explorar a estrutura de cada planilha e depois "
        "levar as fórmulas para funções Python específicas, se quiser "
        "reproduzir 100% da lógica do Excel."
    )
