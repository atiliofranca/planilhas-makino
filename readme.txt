depois de clonar o projeto, na pasta do projeto:

Passo 1 - Crie um Ambiente Virtual:

python -m venv venv
ou
python3 -m venv venv

Depende da sua versão de python instalado

Passo 2 - Ative o Ambiente Virtual:

No Linux ou macOS:
source venv/bin/activate

No Windows (PowerShell/CMD):
.\venv\Scripts\activate

se no windows surgir um erro, impedindo que esse comando seja carregado porque a execução de scripts foi desabilitada no sistema, usar o seguinte comando antes:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

Depois de criado, você verá (venv) aparecer no início do seu prompt do terminal, indicando que o ambiente virtual está ativo

Passo 3 - Instale as bibliotecas dentro do seu ambiente virtual:

pip install -r requirements.txt
