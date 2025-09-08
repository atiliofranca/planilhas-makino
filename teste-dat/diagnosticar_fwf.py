import pandas as pd
import sys

# --- CONFIGURAÇÃO ---
# Coloque o caminho e nome exato do seu arquivo .dat aqui
ARQUIVO_DAT = 'teste-dat/SM 11POS.DAT' 
# --------------------

print(f"Iniciando diagnóstico para: {ARQUIVO_DAT}\n")

# --- PARTE 1: TENTATIVA DE DETECÇÃO AUTOMÁTICA ---
print("="*40)
print("--- 1. TENTATIVA DE DETECÇÃO AUTOMÁTICA ---")
print("="*40)

try:
    # Pede ao pandas para ler o arquivo e adivinhar as colunas baseado nas 100 primeiras linhas
    df_teste = pd.read_fwf(
        ARQUIVO_DAT, 
        infer_nrows=100,  # Analisa 100 linhas para adivinhar
        header=None,
        encoding='latin-1'
    )
    
    print("Pandas detectou automaticamente as seguintes colunas:")
    print(df_teste.head())
    print(f"\nSucesso: Pandas detectou {len(df_teste.columns)} colunas.")
    print("Verifique a tabela acima. Se as colunas estiverem separadas corretamente,")
    print("você pode simplesmente usar 'pd.read_fwf(nome_arquivo_2, header=None, encoding='latin-1')' no seu script principal!")

except FileNotFoundError:
    print(f"ERRO: Arquivo não encontrado em: {ARQUIVO_DAT}")
    print("Por favor, verifique o caminho e nome do arquivo na variável ARQUIVO_DAT.")
    sys.exit() # Para o script se não achar o arquivo
except Exception as e:
    print(f"Detecção automática falhou ou encontrou um erro: {e}")


# --- PARTE 2: AJUDANTE MANUAL (RÉGUA) ---
print("\n" + "="*40)
print("--- 2. AJUDANTE MANUAL (RÉGUA) ---")
print("="*40)
print("Use as réguas abaixo para contar as posições de INÍCIO e FIM de cada coluna.")
print("Lembre-se: a contagem começa em 0.\n")

try:
    # Cria duas réguas para facilitar a contagem
    regua_dezenas = ""
    regua_unidades = ""
    
    # Cria uma régua de 150 caracteres
    for i in range(15): 
        regua_dezenas += str(i).ljust(10) # Marca 0, 10, 20...
        regua_unidades += "0123456789"    # Marca 0..9 repetidamente
    
    print(regua_dezenas)
    print(regua_unidades)
    
    # Imprime as 10 primeiras linhas do arquivo logo abaixo da régua
    with open(ARQUIVO_DAT, 'r', encoding='latin-1') as f:
        for i in range(10):
            linha = f.readline().strip('\n') # Lê a linha e remove a quebra de linha
            if not linha: # Para se o arquivo acabar antes
                break
            print(linha)

except Exception as e:
    print(f"Ocorreu um erro ao ler o arquivo para a régua manual: {e}")