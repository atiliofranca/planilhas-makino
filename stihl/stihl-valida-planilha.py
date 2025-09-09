import pandas as pd
import tkinter as tk
from tkinter import filedialog

def comparar_planilhas(caminho_antigo, caminho_novo):
    """
    Compara a estrutura de duas planilhas Excel, verificando abas e colunas.
    """
    print("-" * 50)
    print(f"Comparando:\n  - ANTIGA: {caminho_antigo}\n  - NOVA:   {caminho_novo}")
    print("-" * 50)

    houve_mudanca = False

    try:
        excel_antigo = pd.ExcelFile(caminho_antigo, engine='openpyxl')
        excel_novo = pd.ExcelFile(caminho_novo, engine='openpyxl')
    except Exception as e:
        print(f"❌ ERRO: Não foi possível ler um dos arquivos. Detalhe: {e}")
        return

    nomes_abas_antigo = excel_antigo.sheet_names
    nomes_abas_novo = excel_novo.sheet_names

    # --- 1. COMPARAÇÃO DAS ABAS ---
    print("\n--- 1. Verificando Nomes e Ordem das Abas ---")
    
    set_abas_antigo = set(nomes_abas_antigo)
    set_abas_novo = set(nomes_abas_novo)

    if set_abas_antigo == set_abas_novo:
        print("✅ Nomes das abas são os mesmos.")
        if nomes_abas_antigo == nomes_abas_novo:
            print("✅ Ordem das abas está idêntica.")
        else:
            print("❌ ALERTA: A ORDEM das abas mudou!")
            houve_mudanca = True
    else:
        houve_mudanca = True
        print("❌ ALERTA: NOMES DAS ABAS FORAM ALTERADOS!")
        abas_removidas = set_abas_antigo - set_abas_novo
        abas_adicionadas = set_abas_novo - set_abas_antigo
        if abas_removidas:
            print(f"  - Abas Removidas: {list(abas_removidas)}")
        if abas_adicionadas:
            print(f"  - Abas Adicionadas: {list(abas_adicionadas)}")

    # --- 2. COMPARAÇÃO DAS COLUNAS (CABEÇALHOS) ---
    print("\n--- 2. Verificando Colunas (primeira linha) de Cada Aba ---")
    abas_comuns = sorted(list(set_abas_antigo.intersection(set_abas_novo)))
    
    if not abas_comuns:
        print("Nenhuma aba em comum para comparar colunas.")
    
    for nome_aba in abas_comuns:
        try:
            # Lê apenas a primeira linha de cada aba para pegar os cabeçalhos
            df_antigo = pd.read_excel(caminho_antigo, sheet_name=nome_aba, header=None, nrows=1, engine='openpyxl')
            df_novo = pd.read_excel(caminho_novo, sheet_name=nome_aba, header=None, nrows=1, engine='openpyxl')
            
            # Converte a primeira linha em uma lista de nomes de coluna
            colunas_antigo = df_antigo.iloc[0].astype(str).tolist()
            colunas_novo = df_novo.iloc[0].astype(str).tolist()
            
            print(f"\nVerificando Aba: '{nome_aba}'...")
            
            if set(colunas_antigo) == set(colunas_novo):
                if colunas_antigo == colunas_novo:
                    print("  ✅ Nomes e ordem das colunas estão idênticos.")
                else:
                    print(f"  ❌ ALERTA: A ORDEM das colunas na aba '{nome_aba}' mudou!")
                    houve_mudanca = True
            else:
                houve_mudanca = True
                print(f"  ❌ ALERTA: NOMES DAS COLUNAS na aba '{nome_aba}' mudaram!")
                colunas_removidas = set(colunas_antigo) - set(colunas_novo)
                colunas_adicionadas = set(colunas_novo) - set(colunas_antigo)
                if colunas_removidas:
                    print(f"    - Colunas Removidas: {list(colunas_removidas)}")
                if colunas_adicionadas:
                    print(f"    - Colunas Adicionadas: {list(colunas_adicionadas)}")
        except Exception as e:
            print(f"\n❌ ERRO ao tentar comparar a aba '{nome_aba}': {e}")
            houve_mudanca = True
            
    # --- 3. RESUMO FINAL ---
    print("\n" + "="*50)
    if houve_mudanca:
        print("🚨 RESUMO: Foram encontradas MUDANÇAS ESTRUTURAIS entre as planilhas!")
        print("   Revise os alertas acima antes de executar o script de processamento principal.")
    else:
        print("✅ RESUMO: Nenhuma mudança estrutural encontrada. As planilhas são compatíveis.")
    print("="*50)


# --- BLOCO PRINCIPAL PARA EXECUTAR O SCRIPT ---
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    print("Por favor, selecione os arquivos para comparação.")
    
    arquivo_antigo_path = filedialog.askopenfilename(
        title="Passo 1 de 2: Selecione a planilha do MÊS ANTERIOR",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if not arquivo_antigo_path:
        print("Seleção cancelada.")
        exit()

    arquivo_novo_path = filedialog.askopenfilename(
        title="Passo 2 de 2: Selecione a planilha do MÊS ATUAL",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if not arquivo_novo_path:
        print("Seleção cancelada.")
        exit()
        
    comparar_planilhas(arquivo_antigo_path, arquivo_novo_path)