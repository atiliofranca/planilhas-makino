import pandas as pd
import tkinter as tk
from tkinter import filedialog

# --- FUNÇÕES AUXILIARES ---

def arredondamento_personalizado(numero):
    if not isinstance(numero, (int, float)):
        return numero
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

def letra_para_indice(letra):
    letra = letra.upper()
    indice = 0
    for char in letra:
        indice = indice * 26 + (ord(char) - ord('A') + 1)
    return indice - 1

# --- CONFIGURAÇÃO ---
nome_arquivo_saida_excel = 'stihl/new-stihl17.xlsx'
nome_arquivo_saida_autcom_csv = 'stihl/new-stihl17-autcom.csv'
COLUNA_DE_BUSCA = 'B'

# --- MAPEAMENTO DAS ABAS E COLUNAS ---
mapeamento_abas = {
    'Lançamentos':                    {'referencia': 'F', 'preco': 'G', 'ipi': 'Q'},
    'MS':                             {'referencia': 'E', 'preco': 'F', 'ipi': 'U'},
    'SABRES CORRENTES PINHÕES LIMAS': {'referencia': 'C', 'preco': 'D', 'ipi': 'J'},
    'ROÇADEIRAS E IMPL':              {'referencia': 'F', 'preco': 'G', 'ipi': 'Q'},
    'CJ.CORTE FS':                    {'referencia': 'C', 'preco': 'D', 'ipi': 'K'},
    'Produtos a Bateria':             {'referencia': 'E', 'preco': 'F', 'ipi': 'S'},
    'OUTRAS MÁQUINAS':                {'referencia': 'E', 'preco': 'F', 'ipi': 'S'},
    'OUTROS':                         {'referencia': 'F', 'preco': 'G', 'ipi': 'P'},
    'PEÇAS':                          {'referencia': 'B', 'preco': 'C', 'ipi': 'I'},
    'ACESSÓRIOS':                     {'referencia': 'C', 'preco': 'D', 'ipi': 'J'},
    'Ferramentas':                    {'referencia': 'B', 'preco': 'C', 'ipi': 'H'},
    'Artigos da Marca':               {'referencia': 'B', 'preco': 'C', 'ipi': 'I'},
    'EPIs':                           {'referencia': 'C', 'preco': 'D', 'ipi': 'K'},
}

# --- SELEÇÃO DE ARQUIVOS COM POP-UP ---
root = tk.Tk()
root.withdraw()

print("Por favor, selecione os arquivos de entrada nas janelas pop-up...")

nome_arquivo_1 = filedialog.askopenfilename(
    title="Passo 1 de 2: Selecione o arquivo.csv importado do Autcom (Lista Base)",
    filetypes=(("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*"))
)

if not nome_arquivo_1:
    print("Seleção cancelada. O programa será encerrado.")
    exit()

nome_arquivo_2 = filedialog.askopenfilename(
    title="Passo 2 de 2: Selecione o arquivo.xlsx do fornecedor (Fonte de Dados com Abas)",
    filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
)

if not nome_arquivo_2:
    print("Seleção cancelada. O programa será encerrado.")
    exit()

print(f"\nArquivo 1 selecionado: {nome_arquivo_1}")
print(f"Arquivo 2 selecionado: {nome_arquivo_2}\n")

# --- LEITURA E PRÉ-PROCESSAMENTO ---
try:
    df1 = pd.read_csv(nome_arquivo_1, sep=';', encoding='latin-1', decimal=',')
    
    print("Lendo todas as abas do arquivo2.xlsx...")
    todas_as_abas_df2 = pd.read_excel(nome_arquivo_2, sheet_name=None, header=None, engine='openpyxl')
    print("Leitura concluída. Consolidando dados para busca rápida...")

    dados_consolidados = {}
    indice_busca = letra_para_indice(COLUNA_DE_BUSCA)

    for nome_aba, mapa_colunas in mapeamento_abas.items():
        if nome_aba in todas_as_abas_df2:
            df_aba = todas_as_abas_df2[nome_aba]
            
            indice_ref = letra_para_indice(mapa_colunas['referencia'])
            indice_preco = letra_para_indice(mapa_colunas['preco'])
            indice_ipi = letra_para_indice(mapa_colunas['ipi'])
            
            for index, linha in df_aba.iterrows():
                referencia = linha.get(indice_ref)
                if pd.notna(referencia) and referencia not in dados_consolidados:
                    dados_consolidados[referencia] = {
                        'preco': linha.get(indice_preco),
                        'ipi': linha.get(indice_ipi)
                    }
    
    print(f"{len(dados_consolidados)} referências únicas consolidadas. Arquivos lidos com sucesso!\n")

except Exception as e:
    print(f"Ocorreu um erro ao ler ou processar os arquivos: {e}")
    exit()

# --- PROCESSAMENTO PRINCIPAL ---
indices_modificados = []
df1['Preço Real Stihl'] = pd.NA

colunas_venda = [
    'Novo Pr. Venda 1', 'Novo Pr. Venda 2', 'Novo Pr. Venda 3', 
    'Novo Pr. Venda 4', 'Novo Pr. Venda 5', 'Novo Pr. Venda 6', 
    'Novo Pr. Venda 7'
]

for index, linha_df1 in df1.iterrows():
    valor_a_procurar = linha_df1['Referência']
    
    if pd.notna(valor_a_procurar) and valor_a_procurar in dados_consolidados:
        dados_encontrados = dados_consolidados[valor_a_procurar]
        
        valor_preco_encontrado = dados_encontrados['preco']
        valor_ipi_encontrado = dados_encontrados['ipi']
        
        try:
            valor_como_numero = float(str(valor_preco_encontrado).replace(',', '.'))
            
            valor_calculado_venda = valor_como_numero * 1.5
            valor_final_venda = arredondamento_personalizado(valor_calculado_venda)
            valor_final_compra = valor_como_numero * 0.67
            valor_final_frete = valor_como_numero * 0.0015
            
        except (ValueError, TypeError):
            valor_final_venda = valor_preco_encontrado
            valor_final_compra = valor_preco_encontrado
            valor_final_frete = valor_preco_encontrado
            print(f"Aviso: Valor de preço '{valor_preco_encontrado}' não é numérico para a referência '{valor_a_procurar}'.")
        
        print(f"Valor '{valor_a_procurar}' encontrado! Atualizando linha {index + 2}...")
        
        df1.loc[index, colunas_venda] = valor_final_venda
        df1.loc[index, 'Novo IPI Entrada'] = valor_ipi_encontrado
        df1.loc[index, 'Novo Pr.Compra'] = valor_final_compra
        df1.loc[index, 'Novo Frete Entrada'] = valor_final_frete
        df1.loc[index, 'Preço Real Stihl'] = valor_preco_encontrado
        
        indices_modificados.append(index)

# --- GERAÇÃO ARQUIVO.XLSX ---
def destacar_celulas(linha):
    colunas_alteradas = colunas_venda + ['Novo IPI Entrada', 'Novo Pr.Compra', 'Novo Frete Entrada']
    if linha.name in indices_modificados:
        estilos = ['background-color: yellow' if col in colunas_alteradas else '' for col in linha.index]
        return estilos
    else:
        return ['' for _ in linha.index]

styled_df = df1.style.apply(destacar_celulas, axis=1)

try:
    styled_df.to_excel(nome_arquivo_saida_excel, index=False, engine='openpyxl')
    print(f"\nArquivo Excel visual foi salvo como '{nome_arquivo_saida_excel}'.")
except Exception as e:
    print(f"\nOcorreu um erro ao salvar o arquivo Excel: {e}")

# --- GERAÇÃO ARQUIVO.CSV AUTCOM ---
try:
    print(f"Iniciando reestruturação para o arquivo '{nome_arquivo_saida_autcom_csv}'...")
    df_para_autcom = df1.copy()

    df_para_autcom['Cód.Item'] = df_para_autcom['Cód.Item'].apply(
        lambda x: str(int(x)).zfill(7) if pd.notna(x) and x != '' else ''
    )
    
    colunas_para_formatar = colunas_venda + ['Novo IPI Entrada', 'Novo Pr.Compra', 'Novo Frete Entrada', 'Preço Real Stihl']
    for col in colunas_para_formatar:
        if col in df_para_autcom.columns:
            df_para_autcom[col] = df_para_autcom[col].apply(
                lambda x: f'{x:g}'.replace('.', ',') if isinstance(x, (int, float)) else x
            )
            
    colunas_para_apagar = ['Novo Departamento', 'Desc. Departamento', 'Unnamed: 14', 'Preço Real Stihl']
    colunas_existentes_para_apagar = [col for col in colunas_para_apagar if col in df_para_autcom.columns]
    df_para_autcom.drop(columns=colunas_existentes_para_apagar, inplace=True)

    mapeamento_posicoes = {
        'Cód.Item': 0, 'Descrição': 6, 'Referência': 9, 'Novo Pr.Compra': 34,
        'Novo IPI Entrada': 43, 'Novo Frete Entrada': 44, 'Novo Pr. Venda 1': 55, 
        'Novo Pr. Venda 2': 58, 'Novo Pr. Venda 3': 61, 'Novo Pr. Venda 4': 64, 
        'Novo Pr. Venda 5': 67, 'Novo Pr. Venda 6': 70, 'Novo Pr. Venda 7': 73,
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
    print(f"Arquivo CSV para Autcom foi salvo como '{nome_arquivo_saida_autcom_csv}'.")

except Exception as e:
    print(f"\nOcorreu um erro ao salvar o arquivo CSV para Autcom: {e}")

print("\nProcesso concluído!")