#!/bin/bash

# Script para construir o executável do Gerador PPTX em Linux/macOS

PYTHON_CMD=python3 # Ou python, dependendo da sua instalação

# 1. Criar ambiente virtual (recomendado)
echo "Criando ambiente virtual 'venv'..."
$PYTHON_CMD -m venv venv
if [ $? -ne 0 ]; then
    echo "Erro ao criar ambiente virtual. Verifique se o Python 3 e o módulo venv estão instalados."
    exit 1
fi

# 2. Ativar ambiente virtual
echo "Ativando ambiente virtual..."
source venv/bin/activate
if [ $? -ne 0 ]; then
    echo "Erro ao ativar ambiente virtual."
    # Tenta remover o venv criado incorretamente
    rm -rf venv
    exit 1
fi

# 3. Instalar dependências
echo "Instalando dependências de requirements.txt..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "Erro ao instalar dependências. Verifique o ficheiro requirements.txt e a sua ligação à internet."
    deactivate
    exit 1
fi

# 4. Instalar PyInstaller
echo "Instalando PyInstaller..."
pip install pyinstaller
if [ $? -ne 0 ]; then
    echo "Erro ao instalar PyInstaller."
    deactivate
    exit 1
fi

# 5. Executar PyInstaller
echo "Construindo o executável com PyInstaller..."
# --onefile: Cria um único executável
# --windowed: Remove a janela de console no Windows (ignorado no Linux/macOS, mas não causa erro)
# --name: Nome do executável final
pyinstaller --onefile --windowed --name GeradorPPTX main.py
if [ $? -ne 0 ]; then
    echo "Erro ao construir o executável com PyInstaller."
    deactivate
    exit 1
fi

# 6. Desativar ambiente virtual
echo "Desativando ambiente virtual..."
deactivate

echo ""
echo "Build concluído com sucesso!"
echo "O executável 'GeradorPPTX' encontra-se na pasta 'dist'."
echo "Pode ser necessário dar permissão de execução: chmod +x dist/GeradorPPTX"

exit 0

