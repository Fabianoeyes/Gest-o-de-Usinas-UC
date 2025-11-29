import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Configuração da Página
st.set_page_config(page_title="Gestão de Usinas e UCs", layout="wide")

# Título Principal
st.title("📊 Painel de Gestão de Usinas e UCs")
st.markdown("---")

# --- FUNÇÃO DE CARREGAMENTO DE DADOS ---
@st.cache_data
def load_data():
    data = {}
    
    # Dicionário mapeando nome amigável -> (nome do arquivo, linhas para pular)
    files_map = {
        "Resumo": ("1Gestão_de_Usinas_e_UC´s_(28.11.25) (1).xls - Quadro Resumo .csv", 1),
        "Operacional": ("1Gestão_de_Usinas_e_UC´s_(28.11.25) (1).xls - Dashboard Operacional.csv", 39),
        "Financeiro": ("1Gestão_de_Usinas_e_UC´s_(28.11.25) (1).xls - Dashboard Financeiro.csv", 289),
        "Inadimplencia": ("1Gestão_de_Usinas_e_UC´s_(28.11.25) (1).xls - TD_Inadimplencia.csv", 34),
        "Usinas": ("1Gestão_de_Usinas_e_UC´s_(28.11.25) (1).xls - Usinas >>.csv", 8),
        "Clientes": ("1Gestão_de_Usinas_e_UC´s_(28.11.25) (1).xls - Base SIGH - Clientes.csv", 8)
    }

    for key, (filename, skip) in files_map.items():
        try:
            # Tenta carregar o CSV. Se falhar, cria um DataFrame vazio para não quebrar o app
            df = pd.read_csv(filename, skiprows=skip, encoding='utf-8', sep=',') # Ajuste o sep se necessário (ex: ';')
            # Limpeza básica: remove colunas totalmente vazias
            df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
            data[key] = df
        except FileNotFoundError:
            st.error(f"Arquivo não encontrado: {filename}")
            data[key] = pd.DataFrame()
        except Exception as e:
            st.warning(f"Erro ao ler {filename}: {e}. Tentando ler sem pular linhas...")
            try:
                data[key] = pd.read_csv(filename, encoding='latin1', sep=';') # Tentativa secundária
            except:
                data[key] = pd.DataFrame()

    return data

# Carregar dados
dfs = load_data()

# --- SIDEBAR (FILTROS GERAIS) ---
st.sidebar.header("Filtros Globais")
st.sidebar.info("Estes filtros afetam as visualizações abaixo.")

# Exemplo de filtro baseado nas Usinas (se a coluna existir)
df_usinas = dfs["Usinas"]
selected_usina = "Todas"
if not df_usinas.empty and len(df_usinas.columns) > 1:
    col_usina_nome = df_usinas.columns[0] # Assumindo que a 1ª coluna é o nome
    usinas_list = ["Todas"] + list(df_usinas[col_usina_nome].unique())
    selected_usina = st.sidebar.selectbox("Selecione a Usina:", usinas_list)

# --- LAYOUT DE ABAS ---
tab1, tab2, tab3, tab4 = st.tabs(["🏠 Visão Geral", "⚡ Operacional", "💰 Financeiro", "📋 Dados Brutos"])

# --- ABA 1: VISÃO GERAL ---
with tab1:
    st.header("Resumo Executivo")
    
    # Tenta pegar dados do Quadro Resumo
    df_resumo = dfs["Resumo"]
    
    if not df_resumo.empty:
        # Exibindo os primeiros indicadores como métricas (simulando os cards do Excel)
        # Como não sei o nome exato das colunas, pego por índice para demonstrar
        col1, col2, col3, col4 = st.columns(4)
        
        try:
            # Exemplo: Pegando valores da primeira linha do resumo
            val1 = df_resumo.iloc[0, 0] if len(df_resumo.columns) > 0 else 0
            val2 = df_resumo.iloc[0, 1] if len(df_resumo.columns) > 1 else 0
            
            with col1:
                st.metric("Total Usinas Ativas", f"{val1}")
            with col2:
                st.metric("Potência Total (kWp)", f"{val2}")
        except:
            st.write("Dados do quadro resumo não formatados conforme esperado.")
            st.dataframe(df_resumo.head())
    else:
        st.info("Arquivo de Resumo vazio ou não encontrado.")

# --- ABA 2: OPERACIONAL ---
with tab2:
    st.header("Dashboard Operacional")
    df_op = dfs["Operacional"]
    
    if not df_op.empty:
        # Análise Exploratória Automática
        col_names = df_op.columns.tolist()
        
        # Seletores para criar gráficos dinâmicos (replicando Tabelas Dinâmicas)
        c1, c2 = st.columns(2)
        with c1:
            x_axis = st.selectbox("Eixo X (Categoria):", col_names, index=0)
        with c2:
            y_axis = st.selectbox("Eixo Y (Valor):", col_names, index=1 if len(col_names)>1 else 0)
        
        # Gráfico de Barras
        fig_op = px.bar(df_op, x=x_axis, y=y_axis, title=f"Análise Operacional: {x_axis} vs {y_axis}")
        st.plotly_chart(fig_op, use_container_width=True)
        
        st.subheader("Detalhamento Operacional")
        st.dataframe(df_op)
    else:
        st.warning("Dados do Dashboard Operacional não carregados. Verifique o arquivo CSV.")

# --- ABA 3: FINANCEIRO ---
with tab3:
    st.header("Indicadores Financeiros & Inadimplência")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Inadimplência")
        df_inad = dfs["Inadimplencia"]
        if not df_inad.empty:
            # Calculando totais (Fórmulas do Excel replicadas em Python)
            # Assumindo que a última coluna pode ser valores
            try:
                numeric_cols = df_inad.select_dtypes(include=['float64', 'int64']).columns
                if len(numeric_cols) > 0:
                    total_divida = df_inad[numeric_cols[0]].sum()
                    st.metric("Total em Aberto", f"R$ {total_divida:,.2f}")
                    
                    # Gráfico de Pizza da Inadimplência
                    if len(df_inad.columns) > 1:
                        fig_pizza = px.pie(df_inad, values=numeric_cols[0], names=df_inad.columns[0], title="Distribuição da Inadimplência")
                        st.plotly_chart(fig_pizza, use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao calcular métricas: {e}")
            
            st.dataframe(df_inad)
        else:
            st.info("Sem dados de inadimplência.")

    with col2:
        st.subheader("Fluxo Financeiro")
        df_fin = dfs["Financeiro"]
        if not df_fin.empty:
            st.dataframe(df_fin.head(50)) # Mostra as primeiras 50 linhas
        else:
            st.info("Dados financeiros vazios (possível erro no skiprows).")

# --- ABA 4: DADOS BRUTOS ---
with tab4:
    st.header("Explorador de Arquivos")
    file_option = st.selectbox("Visualizar arquivo:", list(dfs.keys()))
    st.write(f"Visualizando dados de: **{file_option}**")
    st.dataframe(dfs[file_option])
