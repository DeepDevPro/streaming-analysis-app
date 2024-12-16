import pandas as pd

# Função para carregar a planilha Excel
def carregar_planilha(caminho_arquivo):
    try:
        # Lê a planilha no formato Excel
        df = pd.read_excel(caminho_arquivo, engine="openpyxl")
        print("Planilha carregada com sucesso!")
        return df
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

# Processar os dados
def processar_dados(df):
    if df is not None:
        print("\nVisualização inicial dos dados:")
        print(df.head())  # Exibe as primeiras linhas para referência

        # Total de receita bruta por artista
        print("\nCalculando receita bruta por artista...")
        receita_bruta_por_artista = df.groupby('Artist Name')['Gross Revenue'].sum()

        # Total de receita líquida por artista
        print("\nCalculando receita líquida por artista...")
        receita_liquida_por_artista = df.groupby('Artist Name')['Net Revenue'].sum()

        # Receita líquida por plataforma
        print("\nCalculando receita líquida por plataforma...")
        receita_por_plataforma = df.groupby('Platform')['Net Revenue'].sum()

        # Receita líquida por país/região
        print("\nCalculando receita líquida por país...")
        receita_por_pais = df.groupby('Country / Region')['Net Revenue'].sum()

        # Receita total líquida
        total_receita_liquida = df['Net Revenue'].sum()

        # Exibindo os resultados
        print("\nReceita bruta por artista:")
        print(receita_bruta_por_artista)

        print("\nReceita líquida por artista:")
        print(receita_liquida_por_artista)

        print("\nReceita líquida por plataforma:")
        print(receita_por_plataforma)

        print("\nReceita líquida por país:")
        print(receita_por_pais)

        print(f"\nReceita total líquida: R$ {total_receita_liquida:,.2f}")

        return {
            'bruta_por_artista': receita_bruta_por_artista,
            'liquida_por_artista': receita_liquida_por_artista,
            'por_plataforma': receita_por_plataforma,
            'por_pais': receita_por_pais,
            'total_liquida': total_receita_liquida
        }
    else:
        print("Erro: DataFrame vazio. Certifique-se de que o arquivo foi carregado corretamente.")
        return None

# Exportar os resultados para um arquivo Excel consolidado
def exportar_resultados(resultados, caminho_exportacao):
    try:
        with pd.ExcelWriter(caminho_exportacao) as writer:
            resultados['bruta_por_artista'].to_excel(writer, sheet_name='Bruta por Artista')
            resultados['liquida_por_artista'].to_excel(writer, sheet_name='Líquida por Artista')
            resultados['por_plataforma'].to_excel(writer, sheet_name='Por Plataforma')
            resultados['por_pais'].to_excel(writer, sheet_name='Por País')
        print(f"Resultados exportados para {caminho_exportacao}")
    except Exception as e:
        print(f"Erro ao exportar os resultados: {e}")

# Função para filtrar dados por período
def filtrar_por_periodo(df, coluna_data, inicio, fim):
    try:
        df[coluna_data] = pd.to_datetime(df[coluna_data])
        df_filtrado = df[(df[coluna_data] >= inicio) & (df[coluna_data] <= fim)]
        print(f"\nDados filtrados de {inicio} a {fim}:")
        print(df_filtrado.head())
        return df_filtrado
    except Exception as e:
        print(f"Erro ao filtrar dados por período: {e}")
        return df

# Caminho da planilha
caminho_arquivo = "extrato_streaming.xlsx"  # Substitua pelo caminho correto do seu arquivo
caminho_exportacao = "resultados_consolidados.xlsx"

# Executa o fluxo
df = carregar_planilha(caminho_arquivo)

if df is not None:
    # Exemplo de filtragem por período (opcional)
    inicio = '2024-01-01'  # Substitua pelas datas desejadas
    fim = '2024-12-31'
    df_filtrado = filtrar_por_periodo(df, 'Sales Month', inicio, fim)

    # Processa os dados
    resultados = processar_dados(df_filtrado)

    # Exporta os resultados
    if resultados:
        exportar_resultados(resultados, caminho_exportacao)
