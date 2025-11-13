@echo off
REM Script para construir o executável do Gerador PPTX no Windows

set PYTHON_CMD=py -3
REM Ou apenas "python" se estiver no PATH e for a versão correta

REM 1. Verificar se o Python está disponível
%PYTHON_CMD% --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Erro: Python 3 nao encontrado. Verifique a sua instalacao e o PATH.
    echo Certifique-se de que o comando '%PYTHON_CMD%' funciona no seu terminal.
    pause
    exit /b 1
)

REM 2. Criar ambiente virtual (recomendado)
echo Criando ambiente virtual 'venv'...
%PYTHON_CMD% -m venv venv
if %errorlevel% neq 0 (
    echo Erro ao criar ambiente virtual. Verifique se o modulo venv esta instalado.
    pause
    exit /b 1
)

REM 3. Ativar ambiente virtual
echo Ativando ambiente virtual...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo Erro ao ativar ambiente virtual.
    rmdir /s /q venv
    pause
    exit /b 1
)

REM 4. Instalar dependências
echo Instalando dependencias de requirements.txt...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Erro ao instalar dependencias. Verifique o ficheiro requirements.txt e a sua ligacao a internet.
    call venv\Scripts\deactivate.bat
    pause
    exit /b 1
)

REM 5. Instalar PyInstaller
echo Instalando PyInstaller...
pip install pyinstaller
if %errorlevel% neq 0 (
    echo Erro ao instalar PyInstaller.
    call venv\Scripts\deactivate.bat
    pause
    exit /b 1
)

REM 6. Executar PyInstaller
echo Construindo o executavel com PyInstaller...
REM --onefile: Cria um único executável
REM --windowed: Remove a janela de console no Windows
REM --name: Nome do executável final
pyinstaller --onefile --windowed --name GeradorPPTX main.py
if %errorlevel% neq 0 (
    echo Erro ao construir o executavel com PyInstaller.
    call venv\Scripts\deactivate.bat
    pause
    exit /b 1
)

REM 7. Desativar ambiente virtual
echo Desativando ambiente virtual...
call venv\Scripts\deactivate.bat

echo.
echo Build concluido com sucesso!
echo O executavel 'GeradorPPTX.exe' encontra-se na pasta 'dist'.

pause
exit /b 0

