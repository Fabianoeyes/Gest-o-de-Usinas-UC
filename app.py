import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from pathlib import Path
from datetime import datetime

# =========================
# CONFIGURAÇÕES GERAIS
# =========================
st.set_page_config(
    page_title="Gestão de Usinas e UCs",
    page_icon="⚡",
    layout="wide",
)

EXCEL_DEFAULT_PATH = Path("Gestao_de_Usinas_e_UCs.xlsx")

# =========================
# FUNÇÕES AUXILIARES
# =========================

def localizar_arquivo_excel(default_path: Path) -> Path | None:
    """Localiza o arquivo Excel no diretório."""
    if default_path.exists():
        return default_path
    
    candidatos = sorted(Path(".").glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if candidatos:
        return candidatos[0]
    
    return None

@st.cache_data
def carregar_planilhas(path: Path):
    """Carrega todas as abas do Excel."""
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        sheets = {}
        
        for nome in xls.sheet_names:
            # Lê com header=None para manter estrutura original
            df = pd.read_excel(xls, sheet_name=nome, header=None)
            sheets[nome.strip()] = df
        
        return sheets
    except Exception as e:
        st.error(f"Erro ao carregar planilha: {e}")
        return {}

def identificar_estrutura_aba(df: pd.DataFrame):
    """Identifica a estrutura da aba (cabeçalhos, dados, etc)."""
    estrutura = {
        'header_row': None,
        'data_start': None,
        'title_row': None,
        'colunas': []
    }
    
    # Procura por linhas de título (células mescladas ou textos longos)
    for idx, row in df.iterrows():
        row_str = row.astype(str)
        # Identifica título se primeira célula tem texto e demais são vazias
        if pd.notna(row.iloc[0]) and row.iloc[1:5].isna().all():
            if estrutura['title_row'] is None:
                estrutura['title_row'] = idx
    
    # Procura linha de cabeçalho (linha com múltiplos valores não-nulos)
    for idx, row in df.iterrows():
        non_null = row.notna().sum()
        if non_null >= 3:  # Pelo menos 3 colunas preenchidas
            if estrutura['header_row'] is None:
                estrutura['header_row'] = idx
                estrutura['data_start'] = idx + 1
                estrutura['colunas'] = row.tolist()
                break
    
    return estrutura

def processar_aba_dashboard_operacional(df: pd.DataFrame):
    """Processa especificamente a aba Dashboard Operacional."""
    dados = {
        'tabelas': [],
        'graficos': []
    }
    
    # Procura por títulos de tabelas
    titulos_encontrados = []
    
    for idx, row in df.iterrows():
        primeira_celula = str(row.iloc[0]).strip()
        
        # Identifica títulos (células que parecem títulos)
        if primeira_celula and len(primeira_celula) > 10 and row.iloc[1:5].isna().all():
            titulos_encontrados.append({'titulo': primeira_celula, 'linha': idx})
    
    # Para cada título, extrai a tabela logo abaixo
    for i, titulo_info in enumerate(titulos_encontrados):
        inicio = titulo_info['linha'] + 1
        
        # Define fim da tabela (próximo título ou final)
        if i < len(titulos_encontrados) - 1:
            fim = titulos_encontrados[i + 1]['linha']
        else:
            fim = len(df)
        
        # Extrai a tabela
        tabela_df = df.iloc[inicio:fim].copy()
        
        # Remove linhas completamente vazias
        tabela_df = tabela_df.dropna(how='all')
        
        if not tabela_df.empty:
            # Primeira linha como cabeçalho
            header = tabela_df.iloc[0]
            tabela_df = tabela_df.iloc[1:]
            tabela_df.columns = header
            tabela_df = tabela_df.reset_index(drop=True)
            
            dados['tabelas'].append({
                'titulo': titulo_info['titulo'],
                'dados': tabela_df
            })
    
    return dados

def processar_aba_generica(df: pd.DataFrame, nome_aba: str):
    """Processa abas genéricas identificando estrutura automaticamente."""
    estrutura = identificar_estrutura_aba(df)
    
    resultado = {
        'titulo': None,
        'dados': None,
        'metadados': {}
    }
    
    # Extrai título se existir
    if estrutura['title_row'] is not None:
        resultado['titulo'] = str(df.iloc[estrutura['title_row'], 0])
    
    # Extrai dados
    if estrutura['header_row'] is not None and estrutura['data_start'] is not None:
        # Usa a linha de cabeçalho
        header_row = estrutura['header_row']
        data_start = estrutura['data_start']
        
        # Cria DataFrame com cabeçalho correto
        dados_df = df.iloc[data_start:].copy()
        dados_df.columns = df.iloc[header_row]
        dados_df = dados_df.reset_index(drop=True)
        
        # Remove linhas completamente vazias
        dados_df = dados_df.dropna(how='all')
        
        resultado['dados'] = dados_df
    else:
        # Se não encontrou estrutura, retorna df original
        resultado['dados'] = df.dropna(how='all')
    
    return resultado

def calcular_metricas_numericas(df: pd.DataFrame):
    """Calcula métricas para colunas numéricas."""
    metricas = {}
    
    if df is None or df.empty:
        return metricas
    
    # Identifica colunas numéricas
    colunas_numericas = df.select_dtypes(include=[np.number]).columns
    
    for col in colunas_numericas:
        col_data = df[col].dropna()
        if len(col_data) > 0:
            metricas[col] = {
                'soma': col_data.sum(),
                'media': col_data.mean(),
                'minimo': col_data.min(),
                'maximo': col_data.max(),
                'count': len(col_data)
            }
    
    return metricas

def criar_grafico_pizza(serie: pd.Series, titulo: str):
    """Cria gráfico de pizza."""
    if serie.empty:
        return None
    
    fig, ax = plt.subplots(figsize=(8, 6))
    colors = plt.cm.Set3(range(len(serie)))
    
    wedges, texts, autotexts = ax.pie(
        serie.values, 
        labels=serie.index,
        autopct='%1.1f%%',
        startangle=90,
        colors=colors
    )
    
    ax.set_title(titulo, fontsize=14, fontweight='bold')
    
    # Melhora legibilidade
    for text in texts:
        text.set_fontsize(10)
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(9)
    
    ax.axis('equal')
    return fig

# =========================
# INTERFACE PRINCIPAL
# =========================

st.sidebar.title("⚡ Gestão de Usinas & UCs")

# Carrega arquivo
excel_path = localizar_arquivo_excel(EXCEL_DEFAULT_PATH)
if not excel_path:
    st.error("❌ Arquivo Excel não encontrado. Coloque o arquivo na mesma pasta do app.")
    st.stop()

sheets = carregar_planilhas(excel_path)
if not sheets:
    st.error("❌ Não foi possível carregar as planilhas.")
    st.stop()

st.sidebar.success(f"✅ Arquivo: {excel_path.name}")

if excel_path.name != EXCEL_DEFAULT_PATH.name:
    st.sidebar.info(f"ℹ️ Usando arquivo mais recente: {excel_path.name}")

# Lista de abas disponíveis
abas_disponiveis = list(sheets.keys())
st.sidebar.markdown(f"**Abas encontradas:** {len(abas_disponiveis)}")

# Menu de navegação
pagina = st.sidebar.radio(
    "📋 Navegação",
    [
        "🏠 Visão Geral",
        "📊 Dashboard Operacional",
        "⚡ Usinas",
        "🏢 UCs",
        "📑 Visualizador Completo",
        "🔧 Análise de Estrutura"
    ],
)

# =========================
# PÁGINA: VISÃO GERAL
# =========================

if pagina == "🏠 Visão Geral":
    st.title("🏠 Visão Geral do Sistema")
    
    st.markdown("### 📊 Informações do Arquivo")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Arquivo carregado", excel_path.name)
    
    with col2:
        st.metric("Total de abas", len(abas_disponiveis))
    
    with col3:
        tamanho_mb = excel_path.stat().st_size / (1024 * 1024)
        st.metric("Tamanho do arquivo", f"{tamanho_mb:.2f} MB")
    
    st.markdown("---")
    st.markdown("### 📑 Abas Disponíveis")
    
    # Cria resumo de cada aba
    for nome_aba in abas_disponiveis:
        with st.expander(f"📄 {nome_aba}"):
            df = sheets[nome_aba]
            
            col_a, col_b, col_c = st.columns(3)
            
            with col_a:
                st.metric("Linhas", len(df))
            
            with col_b:
                st.metric("Colunas", len(df.columns))
            
            with col_c:
                # Conta células preenchidas
                total_cells = df.size
                filled_cells = df.notna().sum().sum()
                pct = (filled_cells / total_cells * 100) if total_cells > 0 else 0
                st.metric("Preenchimento", f"{pct:.1f}%")
            
            # Preview dos dados
            st.markdown("**Preview:**")
            st.dataframe(df.head(10), use_container_width=True, height=300)

# =========================
# PÁGINA: DASHBOARD OPERACIONAL
# =========================

elif pagina == "📊 Dashboard Operacional":
    st.title("📊 Dashboard Operacional")
    
    nome_aba = "Dashboard Operacional"
    
    if nome_aba not in sheets:
        st.error(f"❌ Aba '{nome_aba}' não encontrada.")
        st.info(f"Abas disponíveis: {', '.join(abas_disponiveis)}")
        st.stop()
    
    df_raw = sheets[nome_aba]
    
    # Processa a aba
    dados_processados = processar_aba_dashboard_operacional(df_raw)
    
    if not dados_processados['tabelas']:
        st.warning("⚠️ Nenhuma tabela foi identificada automaticamente.")
        st.markdown("**Dados brutos:**")
        st.dataframe(df_raw, use_container_width=True)
    else:
        # Exibe cada tabela encontrada
        for i, tabela in enumerate(dados_processados['tabelas']):
            st.markdown(f"### {tabela['titulo']}")
            
            df_tabela = tabela['dados']
            
            # Remove colunas completamente vazias
            df_tabela = df_tabela.dropna(axis=1, how='all')
            
            # Exibe tabela
            st.dataframe(df_tabela, use_container_width=True)
            
            # Tenta criar visualizações se houver dados numéricos
            colunas_num = df_tabela.select_dtypes(include=[np.number]).columns
            
            if len(colunas_num) > 0:
                with st.expander(f"📈 Visualizações - {tabela['titulo']}"):
                    # Gráfico de barras
                    st.markdown("**Comparação por coluna:**")
                    df_plot = df_tabela[colunas_num].sum()
                    st.bar_chart(df_plot)
                    
                    # Gráficos de pizza para cada coluna numérica
                    cols_pizza = st.columns(min(len(colunas_num), 3))
                    for idx, col in enumerate(colunas_num[:3]):
                        with cols_pizza[idx]:
                            serie = df_tabela[col].dropna()
                            if len(serie) > 0 and serie.sum() > 0:
                                # Tenta usar índice como rótulo
                                if df_tabela.columns[0] and df_tabela[df_tabela.columns[0]].dtype == 'object':
                                    labels = df_tabela[df_tabela.columns[0]]
                                    serie.index = labels
                                
                                fig = criar_grafico_pizza(serie, str(col))
                                if fig:
                                    st.pyplot(fig)
                                    plt.close()
            
            st.markdown("---")

# =========================
# PÁGINA: USINAS
# =========================

elif pagina == "⚡ Usinas":
    st.title("⚡ Gestão de Usinas")
    
    nome_aba = "Usinas"
    
    if nome_aba not in sheets:
        st.error(f"❌ Aba '{nome_aba}' não encontrada.")
        st.info(f"Abas disponíveis: {', '.join(abas_disponiveis)}")
        st.stop()
    
    df_raw = sheets[nome_aba]
    
    # Processa a aba
    dados = processar_aba_generica(df_raw, nome_aba)
    
    if dados['titulo']:
        st.markdown(f"## {dados['titulo']}")
    
    df_usinas = dados['dados']
    
    if df_usinas is None or df_usinas.empty:
        st.warning("⚠️ Nenhum dado encontrado.")
        st.dataframe(df_raw, use_container_width=True)
    else:
        # Remove colunas vazias
        df_usinas = df_usinas.dropna(axis=1, how='all')
        
        # Métricas
        metricas = calcular_metricas_numericas(df_usinas)
        
        if metricas:
            st.markdown("### 📊 Indicadores Principais")
            
            # Exibe métricas principais
            cols_metricas = st.columns(min(len(metricas), 4))
            
            for idx, (col_nome, valores) in enumerate(list(metricas.items())[:4]):
                with cols_metricas[idx]:
                    st.metric(
                        label=str(col_nome),
                        value=f"{valores['soma']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                        delta=f"Média: {valores['media']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    )
        
        st.markdown("---")
        st.markdown("### 📋 Tabela de Dados")
        
        # Tabela editável
        edited_df = st.data_editor(
            df_usinas,
            use_container_width=True,
            num_rows="dynamic",
            height=400
        )
        
        # Visualizações
        st.markdown("### 📈 Visualizações")
        
        colunas_num = edited_df.select_dtypes(include=[np.number]).columns
        
        if len(colunas_num) > 0:
            tab1, tab2 = st.tabs(["📊 Gráfico de Barras", "📈 Gráfico de Linhas"])
            
            with tab1:
                col_selecionada = st.selectbox("Selecione a coluna", colunas_num, key="bar_usinas")
                if col_selecionada:
                    st.bar_chart(edited_df[col_selecionada])
            
            with tab2:
                col_selecionada = st.selectbox("Selecione a coluna", colunas_num, key="line_usinas")
                if col_selecionada:
                    st.line_chart(edited_df[col_selecionada])

# =========================
# PÁGINA: UCs
# =========================

elif pagina == "🏢 UCs":
    st.title("🏢 Gestão de UCs")
    
    nome_aba = "UCs"
    
    if nome_aba not in sheets:
        st.error(f"❌ Aba '{nome_aba}' não encontrada.")
        st.info(f"Abas disponíveis: {', '.join(abas_disponiveis)}")
        st.stop()
    
    df_raw = sheets[nome_aba]
    
    # Processa a aba
    dados = processar_aba_generica(df_raw, nome_aba)
    
    if dados['titulo']:
        st.markdown(f"## {dados['titulo']}")
    
    df_ucs = dados['dados']
    
    if df_ucs is None or df_ucs.empty:
        st.warning("⚠️ Nenhum dado encontrado.")
        st.dataframe(df_raw, use_container_width=True)
    else:
        # Remove colunas vazias
        df_ucs = df_ucs.dropna(axis=1, how='all')
        
        # Métricas
        metricas = calcular_metricas_numericas(df_ucs)
        
        if metricas:
            st.markdown("### 📊 Indicadores Principais")
            
            cols_metricas = st.columns(min(len(metricas), 4))
            
            for idx, (col_nome, valores) in enumerate(list(metricas.items())[:4]):
                with cols_metricas[idx]:
                    st.metric(
                        label=str(col_nome),
                        value=f"{valores['soma']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                        delta=f"Média: {valores['media']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    )
        
        st.markdown("---")
        st.markdown("### 📋 Tabela de Dados")
        
        edited_df = st.data_editor(
            df_ucs,
            use_container_width=True,
            num_rows="dynamic",
            height=400
        )
        
        # Visualizações
        st.markdown("### 📈 Visualizações")
        
        colunas_num = edited_df.select_dtypes(include=[np.number]).columns
        
        if len(colunas_num) > 0:
            tab1, tab2 = st.tabs(["📊 Gráfico de Barras", "📈 Gráfico de Linhas"])
            
            with tab1:
                col_selecionada = st.selectbox("Selecione a coluna", colunas_num, key="bar_ucs")
                if col_selecionada:
                    st.bar_chart(edited_df[col_selecionada])
            
            with tab2:
                col_selecionada = st.selectbox("Selecione a coluna", colunas_num, key="line_ucs")
                if col_selecionada:
                    st.line_chart(edited_df[col_selecionada])

# =========================
# PÁGINA: VISUALIZADOR COMPLETO
# =========================

elif pagina == "📑 Visualizador Completo":
    st.title("📑 Visualizador Completo de Planilhas")
    
    aba_selecionada = st.selectbox("Selecione a aba", abas_disponiveis)
    
    df = sheets[aba_selecionada]
    
    st.markdown(f"### Aba: **{aba_selecionada}**")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total de linhas", len(df))
    with col2:
        st.metric("Total de colunas", len(df.columns))
    with col3:
        filled = df.notna().sum().sum()
        total = df.size
        st.metric("Células preenchidas", f"{filled} ({filled/total*100:.1f}%)")
    
    st.markdown("---")
    
    # Opções de visualização
    modo = st.radio("Modo de visualização", ["Dados brutos", "Dados processados"], horizontal=True)
    
    if modo == "Dados brutos":
        st.dataframe(df, use_container_width=True, height=600)
    else:
        dados = processar_aba_generica(df, aba_selecionada)
        
        if dados['titulo']:
            st.info(f"📋 Título identificado: {dados['titulo']}")
        
        if dados['dados'] is not None:
            df_processado = dados['dados'].dropna(axis=1, how='all')
            st.dataframe(df_processado, use_container_width=True, height=600)
        else:
            st.warning("Não foi possível processar automaticamente. Exibindo dados brutos.")
            st.dataframe(df, use_container_width=True, height=600)
    
    # Download
    st.markdown("---")
    if st.button("📥 Baixar aba como CSV"):
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download CSV",
            data=csv,
            file_name=f"{aba_selecionada}.csv",
            mime="text/csv"
        )

# =========================
# PÁGINA: ANÁLISE DE ESTRUTURA
# =========================

elif pagina == "🔧 Análise de Estrutura":
    st.title("🔧 Análise de Estrutura das Abas")
    
    st.markdown("""
    Esta página analisa automaticamente a estrutura de cada aba para ajudar
    a identificar cabeçalhos, títulos e organização dos dados.
    """)
    
    aba_selecionada = st.selectbox("Selecione a aba para analisar", abas_disponiveis)
    
    df = sheets[aba_selecionada]
    
    st.markdown(f"### Análise da aba: **{aba_selecionada}**")
    
    estrutura = identificar_estrutura_aba(df)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if estrutura['title_row'] is not None:
            st.success(f"✅ Título na linha {estrutura['title_row']}")
        else:
            st.info("ℹ️ Nenhum título identificado")
    
    with col2:
        if estrutura['header_row'] is not None:
            st.success(f"✅ Cabeçalho na linha {estrutura['header_row']}")
        else:
            st.info("ℹ️ Nenhum cabeçalho identificado")
    
    with col3:
        if estrutura['data_start'] is not None:
            st.success(f"✅ Dados a partir da linha {estrutura['data_start']}")
        else:
            st.info("ℹ️ Início dos dados não identificado")
    
    st.markdown("---")
    st.markdown("### 📊 Mapa de calor - Células preenchidas")
    
    # Cria mapa de calor mostrando onde há dados
    heat_data = df.notna().astype(int)
    
    fig, ax = plt.subplots(figsize=(12, 8))
    im = ax.imshow(heat_data.values[:50], cmap='YlGn', aspect='auto')
    
    ax.set_xlabel('Colunas')
    ax.set_ylabel('Linhas')
    ax.set_title(f'Mapa de preenchimento - {aba_selecionada}')
    
    plt.colorbar(im, ax=ax, label='Preenchido (1) / Vazio (0)')
    st.pyplot(fig)
    plt.close()
    
    st.markdown("---")
    st.markdown("### 🔍 Preview dos dados")
    st.dataframe(df.head(20), use_container_width=True)
