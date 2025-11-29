import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="Gestão de Usinas & UCs", page_icon="⚡", layout="wide")

# =========================
# 1. Encontrar automaticamente o arquivo Excel
# =========================

def encontrar_excel():
    """
    Procura arquivos .xlsx na pasta do app.
    Se houver mais de um, dá preferência a nomes que contenham 'Gest' ou 'Usina'.
    """
    arquivos = list(Path(".").glob("*.xlsx"))
    if not arquivos:
        return None

    preferidos = [
        f for f in arquivos
        if "gest" in f.name.lower() or "usina" in f.name.lower() or "uc" in f.name.lower()
    ]
    if preferidos:
        return preferidos[0]

    return arquivos[0]


EXCEL_PATH = encontrar_excel()

if EXCEL_PATH is None:
    st.error(
        "Nenhum arquivo .xlsx foi encontrado na pasta do app.\n\n"
        "Suba um arquivo Excel no repositório (por exemplo: Gestao_de_Usinas_e_UCs.xlsx)."
    )
    st.stop()

st.sidebar.success(f"Arquivo Excel encontrado: {EXCEL_PATH.name}")

# =========================
# 2. Carregar todas as abas da planilha
# =========================

@st.cache_data
def carregar_planilhas(path: Path):
    """
    Lê todas as abas do Excel em um dict: {nome_aba: DataFrame}.
    """
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")

try:
    sheets = carregar_planilhas(EXCEL_PATH)
except Exception as e:
    st.error(f"Erro ao ler o arquivo Excel: {e}")
    st.stop()

nomes_abas = list(sheets.keys())

# =========================
# 3. Interface básica
# =========================

st.sidebar.title("Gestão de Usinas & UCs")

aba_escolhida = st.sidebar.selectbox("Escolha a aba da planilha:", nomes_abas)

st.title("Gestão de Usinas & UCs – Visualização da planilha")
st.markdown(f"### Aba selecionada: **{aba_escolhida}**")

df = sheets[aba_escolhida]

st.dataframe(df, use_container_width=True)

st.markdown("---")
st.write("Colunas numéricas detectadas:")
st.write(df.select_dtypes("number").head())

st.info(
    "✅ App básico funcionando. Agora que a leitura da planilha está OK, "
    "podemos evoluir para dashboards (cards, gráficos e lógicas específicas) "
    "por usina / UC."
)
