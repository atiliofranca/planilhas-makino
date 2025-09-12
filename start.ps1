# ===================================================================
# Script Inteligente de Configuração e Ativação do Ambiente Virtual
# ===================================================================

$VenvFolder = ".\venv"
Write-Host "Iniciando script de configuração do ambiente..." -ForegroundColor Cyan

# --- Passo 1: Verifica se a pasta do ambiente virtual existe ---
if (-not (Test-Path -Path $VenvFolder -PathType Container)) {
    Write-Host "Pasta '$VenvFolder' não encontrada. Criando novo ambiente virtual..." -ForegroundColor Yellow
    
    # Cria o ambiente virtual
    python -m venv venv
    
    # Instala as dependências
    if (Test-Path -Path ".\requirements.txt") {
        Write-Host "Ambiente criado. Instalando dependências de 'requirements.txt'..." -ForegroundColor Yellow
        & "$VenvFolder\Scripts\pip.exe" install -r requirements.txt
    } else {
        Write-Host "AVISO: Arquivo 'requirements.txt' não encontrado." -ForegroundColor Red
    }
    
} else {
    Write-Host "Ambiente virtual '$VenvFolder' já existe." -ForegroundColor Green
}

# --- Passo 2: Ativa o ambiente virtual ---
# Este comando agora será executado no seu terminal principal
# graças ao "dot sourcing"
Write-Host "Ativando o ambiente virtual..." -ForegroundColor Cyan
.\venv\Scripts\activate

Write-Host ""
Write-Host "Ambiente virtual (venv) ativado com sucesso!" -ForegroundColor Green