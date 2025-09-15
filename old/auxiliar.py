import pandas as pd

# --- CONFIGURAÇÃO ---
nome_arquivo_1 = 'arquivo1.csv'
nome_arquivo_novo_filtro = 'arquivo1-novo-filtro.csv'
#nome_arquivo_2 = 'arquivo2.xlsx'
#nome_arquivo_saida_excel = 'new-arquivo1.xlsx'
#nome_arquivo_saida_csv = 'new-arquivo1.csv'
#nome_arquivo_saida_autcom_csv = 'new-arquivo1-autcom.csv'

# --- LEITURA DAS PLANILHAS ---
try:
    df1 = pd.read_csv(nome_arquivo_1, sep=';', encoding='latin-1', decimal=',')
    #df2 = pd.read_excel(nome_arquivo_2, header=None)
    df3 = pd.read_csv(nome_arquivo_novo_filtro, sep=';', encoding='latin-1', decimal=',')
    print("Arquivo(s) lido(s) com sucesso!")
    print(df1.columns)
    #print(df2.columns)
    print(df3.columns)
except Exception as e:
    print(f"Ocorreu um erro ao ler o(s) arquivo(s): {e}")
    exit()