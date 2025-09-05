import pandas as pd

# --- FUNÇÃO DE ARREDONDAMENTO PERSONALIZADO ---
def arredondamento_personalizado(numero):
    # - para decimais de .01 a .49 -> arredonda para .50
    # - para decimais de .51 a .99 -> arredonda para o próximo inteiro
    # - para decimais .00 e .50 -> mantém o valor
    if not isinstance(numero, (int, float)):
        return numero

    # round() para evitar imprecisões de ponto flutuante (ex: 0.4999999...)
    parte_decimal = round(numero % 1, 2)
    parte_inteira = int(numero)

    if parte_decimal == 0.00 or parte_decimal == 0.50:
        return float(numero)
    elif parte_decimal > 0.00 and parte_decimal < 0.50:
        return parte_inteira + 0.50
    elif parte_decimal > 0.50:
        return float(parte_inteira + 1)
    else:
        return float(numero)

# --- CONFIGURAÇÃO ---
nome_arquivo_1 = 'arquivo1.csv'
nome_arquivo_2 = 'arquivo2.xlsx'
nome_arquivo_saida_excel = 'new-arquivo1.xlsx'
nome_arquivo_saida_autcom_csv = 'new-arquivo1-autcom.csv'

# --- LEITURA DAS PLANILHAS ---
try:
    df1 = pd.read_csv(nome_arquivo_1, sep=';', encoding='latin-1', decimal=',')
    df2 = pd.read_excel(nome_arquivo_2, header=None)
    print("Arquivos lidos com sucesso!")
except Exception as e:
    print(f"Ocorreu um erro ao ler os arquivos: {e}")
    exit()

# --- RASTREAMENTO E PROCESSAMENTO ---
indices_modificados = []
colunas_para_alterar_valor_c = [
    'Novo Pr. Venda 1', 'Novo Pr. Venda 2', 'Novo Pr. Venda 3', 
    'Novo Pr. Venda 4', 'Novo Pr. Venda 5', 'Novo Pr. Venda 6', 
    'Novo Pr. Venda 7'
]
coluna_para_alterar_valor_i = 'Novo IPI Entrada'

for index, linha_df1 in df1.iterrows():
    valor_a_procurar = linha_df1['Referência']
    if pd.isna(valor_a_procurar):
        continue
    linhas_encontradas = df2[df2[1] == valor_a_procurar]
    if not linhas_encontradas.empty:
        primeira_linha_encontrada = linhas_encontradas.iloc[0]
        valor_da_coluna_c_df2 = primeira_linha_encontrada[2]
        try:
            valor_como_numero = float(str(valor_da_coluna_c_df2).replace(',', '.'))
            
            # cálculo para os preços de venda (multiplica por 1.5 e arredonda com a nova função)
            valor_calculado_venda = valor_como_numero * 1.5
            valor_final_venda = arredondamento_personalizado(valor_calculado_venda)
            
            # cálculo para o preço de compra (multiplica por 0.67, sem arredondar)
            valor_final_compra = valor_como_numero * 0.67
            
        except (ValueError, TypeError):
            valor_final_venda = valor_da_coluna_c_df2
            valor_final_compra = valor_da_coluna_c_df2
            print(f"Aviso: Valor '{valor_da_coluna_c_df2}' na planilha 2 não é numérico. Não foi aplicado o cálculo para a referência '{valor_a_procurar}'.")
        
        valor_da_coluna_i_df2 = primeira_linha_encontrada[8]
        print(f"Valor '{valor_a_procurar}' encontrado! Atualizando linha {index + 2}...")
        
        df1.loc[index, colunas_para_alterar_valor_c] = valor_final_venda
        df1.loc[index, coluna_para_alterar_valor_i] = valor_da_coluna_i_df2
        df1.loc[index, 'Novo Pr.Compra'] = valor_final_compra
        
        indices_modificados.append(index)

# --- GERAÇÃO ARQUIVO.XLSX ---
def destacar_celulas(linha):
    colunas_alteradas = colunas_para_alterar_valor_c + [coluna_para_alterar_valor_i] + ['Novo Pr.Compra']
    if linha.name in indices_modificados:
        estilos = ['background-color: yellow' if col in colunas_alteradas else '' for col in linha.index]
        return estilos
    else:
        return ['' for _ in linha.index]

styled_df = df1.style.apply(destacar_celulas, axis=1)

try:
    styled_df.to_excel(nome_arquivo_saida_excel, index=False, engine='openpyxl')
    print(f"\nArquivo Excel visual (com vírgula) foi salvo como '{nome_arquivo_saida_excel}'.")
except Exception as e:
    print(f"\nOcorreu um erro ao salvar o arquivo Excel: {e}")

# --- GERAÇÃO ARQUIVO.CSV AUTCOM ---
try:
    print(f"Iniciando reestruturação para o arquivo '{nome_arquivo_saida_autcom_csv}'...")
    df_para_autcom = df1.copy()
    
    colunas_para_formatar = colunas_para_alterar_valor_c + [coluna_para_alterar_valor_i] + ['Novo Pr.Compra']
    for col in colunas_para_formatar:
        if col in df_para_autcom.columns:
            df_para_autcom[col] = df_para_autcom[col].apply(
                lambda x: f'{x:g}'.replace('.', ',') if isinstance(x, (int, float)) else x
            )
            
    colunas_para_apagar = ['Novo Departamento', 'Desc. Departamento', 'Unnamed: 14']
    colunas_existentes_para_apagar = [col for col in colunas_para_apagar if col in df_para_autcom.columns]
    df_para_autcom.drop(columns=colunas_existentes_para_apagar, inplace=True)

    mapeamento_posicoes = {
        'Cód.Item': 0, 'Descrição': 6, 'Referência': 9, 'Novo Pr.Compra': 34,
        'Novo IPI Entrada': 43, 'Novo Pr. Venda 1': 55, 'Novo Pr. Venda 2': 58,
        'Novo Pr. Venda 3': 61, 'Novo Pr. Venda 4': 64, 'Novo Pr. Venda 5': 67,
        'Novo Pr. Venda 6': 70, 'Novo Pr. Venda 7': 73,
    }
    df_reestruturado = pd.DataFrame()
    for nome_coluna, nova_posicao in mapeamento_posicoes.items():
        if nome_coluna in df_para_autcom.columns:
            df_reestruturado[nova_posicao] = df_para_autcom[nome_coluna]
    todas_as_colunas = range(74)
    df_reestruturado = df_reestruturado.reindex(columns=todas_as_colunas)

    cabecalho_final = [''] * 74
    for nome_coluna, nova_posicao in mapeamento_posicoes.items():
        cabecalho_final[nova_posicao] = nome_coluna
    df_reestruturado.columns = cabecalho_final

    df_reestruturado.to_csv(nome_arquivo_saida_autcom_csv, index=False, sep=';', encoding='latin-1', header=True)
    print(f"Arquivo CSV para Autcom (reestruturado e com cabeçalho) foi salvo como '{nome_arquivo_saida_autcom_csv}'.")

except Exception as e:
    print(f"\nOcorreu um erro ao salvar o arquivo CSV para Autcom: {e}")

print("\nProcesso concluído!")