@echo off
setlocal

echo ===============================================
echo  Captador de Dados CRM - Gerador do executavel
echo ===============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERRO] Python nao foi encontrado no PATH.
    echo Instale o Python 3.10 ou superior em https://www.python.org/downloads/
    echo IMPORTANTE: marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo Criando ambiente virtual isolado (pasta "venv")...
python -m venv venv
if errorlevel 1 (
    echo [ERRO] Nao foi possivel criar o ambiente virtual.
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

echo.
echo Instalando dependencias (isso pode levar alguns minutos)...
python -m pip install --upgrade pip
pip install -r requisitos.txt
if errorlevel 1 (
    echo [ERRO] Falha ao instalar as dependencias. Verifique sua conexao com a internet.
    pause
    exit /b 1
)

echo.
echo Gerando o executavel com PyInstaller...
pyinstaller --noconsole --onefile ^
  --name "Captador de Dados CRM" ^
  --icon "assets\icone_captador.ico" ^
  --add-data "assets;assets" ^
  --collect-all customtkinter ^
  --collect-submodules selenium ^
  --collect-submodules webdriver_manager ^
  Captador_de_Dados_CRM.py

echo.
if exist "dist\Captador de Dados CRM.exe" (
    echo ===============================================
    echo  PRONTO! O executavel esta em:
    echo  dist\Captador de Dados CRM.exe
    echo ===============================================
) else (
    echo [ERRO] O executavel nao foi gerado. Veja as mensagens acima para o motivo.
)

echo.
pause
