@echo off
REM Script para criar executável do Extrator de Comprovantes para Windows

echo ====================================
echo Criando executavel para Windows...
echo ====================================
echo.

REM Verificar se ambiente virtual existe
if not exist "venv" (
    echo Criando ambiente virtual...
    python -m venv venv
)

REM Ativar ambiente virtual
echo Ativando ambiente virtual...
call venv\Scripts\activate.bat

REM Instalar dependências
echo Instalando dependencias...
pip install --upgrade pip
pip install pandas openpyxl xlrd PyPDF2 pdfplumber pyinstaller

REM Criar executável
echo Compilando executavel...
pyinstaller --onefile --windowed --name="Extrator_Comprovantes" get_proof.py

REM Desativar ambiente virtual
deactivate

echo.
echo ====================================
echo Executavel criado com sucesso!
echo ====================================
echo.
echo Localizacao: dist\Extrator_Comprovantes.exe
echo.
echo Para distribuir:
echo   - Copie o arquivo dist\Extrator_Comprovantes.exe para outras maquinas Windows
echo   - Duplo clique no .exe para executar
echo.
pause