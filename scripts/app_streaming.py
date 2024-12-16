import pandas as pd
import streamlit as st

# Função para carregar a planilha
@st.cache_data
def carregar_planilha(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo, engine="openpyxl")
        return df
    except Exception as e:
        st.error(f"Erro ao carregar a planilha: {e}")
        return None

# Função para processar os dados
def processar_dados(df):
    # Tabela 1: Resumo por OBRA
    tabela_obra = df.groupby(['ISRC', 'Track title', 'Sales Type']).agg(
        UNIDADES=('Quantity', 'sum'),
        VALOR=('Net Revenue', 'sum')
    ).reset_index()
    tabela_obra = tabela_obra.rename(columns={
        'ISRC': 'ISRC',
        'Track title': 'OBRA',
        'Sales Type': 'TIPO',
        'UNIDADES': 'UNIDADES',
        'VALOR': 'VALOR (€ Euro)'
    }).sort_values(by='OBRA')

    # Totais da tabela OBRA
    total_obra_unidades = tabela_obra['UNIDADES'].sum()
    total_obra_valor = tabela_obra['VALOR (€ Euro)'].sum()

    # Tabela 2: Resumo por TIPO
    tabela_tipo = df.groupby('Sales Type').agg(
        UNIDADES=('Quantity', 'sum'),
        VALOR=('Net Revenue', 'sum')
    ).reset_index()
    tabela_tipo = tabela_tipo.rename(columns={
        'Sales Type': 'TIPO',
        'UNIDADES': 'UNIDADES',
        'VALOR': 'VALOR (€ Euro)'
    })

    # Totais da tabela TIPO
    total_tipo_unidades = tabela_tipo['UNIDADES'].sum()
    total_tipo_valor = tabela_tipo['VALOR (€ Euro)'].sum()

    return tabela_obra, total_obra_unidades, total_obra_valor, tabela_tipo, total_tipo_unidades, total_tipo_valor

# Caminho fixo da planilha padrão
caminho_padrao = "data/extrato_streaming.xlsx"

# Interface do App
st.title("Análise de Streamings e Downloads")
st.sidebar.header("Configurações")

# Carregar a planilha padrão ou permitir upload manual
df = carregar_planilha(caminho_padrao)
if df is not None:
    st.sidebar.success("Planilha carregada automaticamente!")
else:
    st.sidebar.error("Não foi possível carregar a planilha padrão.")
    caminho_arquivo = st.sidebar.file_uploader("Envie sua planilha", type=["xlsx"])
    if caminho_arquivo:
        df = carregar_planilha(caminho_arquivo)

# Somente prosseguir se a planilha foi carregada
if df is not None:
    # Total Global da Planilha (antes dos filtros)
    total_global = df['Net Revenue'].sum()

    # Filtros
    st.sidebar.subheader("Filtros")

    # Filtro por Período
    inicio = st.sidebar.date_input("Data Inicial", value=pd.Timestamp("2024-01-01"))
    fim = st.sidebar.date_input("Data Final", value=pd.Timestamp("2024-12-31"))
    df['Reporting month'] = pd.to_datetime(df['Reporting month'])
    inicio = pd.to_datetime(inicio)
    fim = pd.to_datetime(fim)
    df_filtrado_periodo = df[(df['Reporting month'] >= inicio) & (df['Reporting month'] <= fim)]

    # Filtro por Artista
    artistas_disponiveis = df_filtrado_periodo['Artist Name'].unique()
    artistas_selecionados = st.sidebar.multiselect(
        "Filtrar por Artista(s)", options=artistas_disponiveis, default=artistas_disponiveis
    )
    df_filtrado = df_filtrado_periodo[df_filtrado_periodo['Artist Name'].isin(artistas_selecionados)]

    # Total Global após os filtros
    total_filtrado = df_filtrado['Net Revenue'].sum()

    # Período selecionado
    st.markdown("")  # Espaço
    st.subheader(f"Período Selecionado: {inicio.date()} a {fim.date()}")

    # Card Totalizador
    st.markdown("")  # Espaço
    st.subheader("Resumo Total")
    st.metric(
        label="Receita Líquida (Net Revenue)",
        value=f"R$ {total_filtrado:,.2f}",
        delta=f"Total Original: R$ {total_global:,.2f}"
    )

    # Processar os dados filtrados
    tabela_obra, total_obra_unidades, total_obra_valor, tabela_tipo, total_tipo_unidades, total_tipo_valor = processar_dados(df_filtrado)

    # Tabela 1: Resumo por OBRA
    st.markdown("")  # Espaço
    st.subheader("Resumo por OBRA")
    st.dataframe(tabela_obra)

    # Destacar totais gerais da tabela OBRA
    st.markdown(
        f"**Total de UNIDADES:** {total_obra_unidades:,} | **Total de VALOR (€ Euro):** € {total_obra_valor:,.2f}"
    )

    # Tabela 2: Resumo por TIPO
    st.markdown("")  # Espaço
    st.subheader("Resumo por TIPO")
    st.dataframe(tabela_tipo)

    # Destacar totais gerais da tabela TIPO
    st.markdown(
        f"**Total de UNIDADES:** {total_tipo_unidades:,} | **Total de VALOR (€ Euro):** € {total_tipo_valor:,.2f}"
    )
