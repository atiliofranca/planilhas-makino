#!/bin/bash
# ===================================================================
# Script Inteligente de Configuração e Ativação para Linux/macOS
# ===================================================================

# Define o nome da pasta do ambiente virtual
VENV_FOLDER="venv"

echo "Iniciando script de configuração do ambiente..."

# --- Passo 1: Verifica se a pasta do ambiente virtual existe ---
if [ ! -d "$VENV_FOLDER" ]; then
    # Se a pasta NÃO existe, executa a configuração inicial
    
    echo "Pasta '$VENV_FOLDER' não encontrada. Criando novo ambiente virtual..."
    
    # 1.1. Cria o ambiente virtual (usando python3, o padrão no Linux)
    python3 -m venv venv
    
    # 1.2. Verifica se o requirements.txt existe antes de instalar
    if [ -f "requirements.txt" ]; then
        echo "Ambiente criado. Instalando dependências de 'requirements.txt'..."
        # Chama o pip de dentro do venv recém-criado para instalar os pacotes
        ./venv/bin/pip install -r requirements.txt
    else
        echo "AVISO: Arquivo 'requirements.txt' não encontrado. Nenhuma dependência foi instalada."
    fi
    
else
    # Se a pasta JÁ existe, apenas informa o usuário
    echo "Ambiente virtual '$VENV_FOLDER' já existe."
fi

# --- Passo 2: Ativa o ambiente virtual ---
# Este comando será executado no seu terminal principal
# porque o script inteiro está sendo executado com "source"
echo "Ativando o ambiente virtual..."
source venv/bin/activate

echo ""
echo "Ambiente virtual (venv) ativado com sucesso!"