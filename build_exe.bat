@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM Ir para a pasta deste script
cd /d "%~dp0"

echo Limpando pastas anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__

REM Detectar o nome exato do PDF (lida com variações antes do - SINDNORTE)
set "PDF_NAME="
for %%F in ("formulario.pdf") do (
  set "PDF_NAME=%%~nxF"
  goto :foundpdf
)
:foundpdf

if not defined PDF_NAME (
  echo.
  echo ERRO: Nao encontrei o PDF que comeca com:
  echo   "formulario.pdf"
  echo Arquivos .pdf no diretorio atual:
  dir /b *.pdf
  echo.
  echo Dica: renomeie o PDF para um nome simples, por exemplo: SINDNORTE_form.pdf
  echo e atualize o script, ou me informe o nome exato do arquivo.
  pause
  exit /b 1
)

REM Escolher o executavel do PyInstaller (usa o da venv se existir)
set "PYI=pyinstaller"
if exist ".venv\Scripts\pyinstaller.exe" (
  set "PYI=.venv\Scripts\pyinstaller.exe"
) else (
  where pyinstaller >nul 2>&1
  if errorlevel 1 (
    echo ERRO: PyInstaller nao encontrado. Ative sua virtualenv ou instale com:
    echo   python -m pip install pyinstaller
    pause
    exit /b 1
  )
)

echo Usando PDF: "!PDF_NAME!"
echo Iniciando build...

"%PYI%" --noconfirm --onefile --windowed --add-data "logo.png;." --add-data "empresas.db;." --add-data "cnpj_formatter.py;." --add-data "pdf_mapping.py;." --add-data "pdf_filler.py;." --add-data "!PDF_NAME!;." main.py

set "RC=%ERRORLEVEL%"
echo.
if %RC% neq 0 (
  echo ERRO: Build falhou com codigo %RC%.
  pause
  exit /b %RC%
)

if exist "dist\main.exe" (
  echo ===== Build concluido! =====
  echo Verifique: dist\main.exe
) else (
  echo ATENCAO: Build terminou, mas dist\main.exe nao foi encontrado. Conteudo de dist/:
  dir /b dist
)

pause
endlocal
