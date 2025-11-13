@echo off
chcp 65001 >nul
echo ========================================
echo  🏛️ Verificador INCRA Pro - Build
echo ========================================
echo.

REM Verificar se PyInstaller está instalado
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo ❌ PyInstaller não encontrado. Instalando...
    pip install pyinstaller
    echo.
)

REM Verificar se a pasta poppler existe
if not exist "poppler" (
    echo.
    echo ⚠️ ATENÇÃO: Pasta 'poppler' não encontrada!
    echo.
    echo Por favor, siga estas etapas:
    echo 1. Baixe o Poppler para Windows em:
    echo    https://github.com/oschwartz10612/poppler-windows/releases/
    echo.
    echo 2. Baixe o arquivo: Release-XX.XX.X-0.zip
    echo.
    echo 3. Extraia o conteúdo na pasta do projeto
    echo.
    echo 4. Renomeie a pasta extraída para 'poppler'
    echo    (deve conter: poppler/Library/bin/...)
    echo.
    echo 5. Execute este script novamente
    echo.
    pause
    exit /b 1
)

REM Verificar estrutura do Poppler
if not exist "poppler\Library\bin" (
    echo ❌ Estrutura do Poppler incorreta!
    echo Certifique-se que existe: poppler\Library\bin\
    pause
    exit /b 1
)

echo ✅ Poppler encontrado em: poppler\Library\bin
echo.

echo [1/4] 🧹 Limpando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec
echo ✅ Limpeza concluída
echo.

echo [2/4] 📦 Preparando dependências...
pip install --upgrade pillow google-generativeai openpyxl PyPDF2 pdf2image
echo ✅ Dependências atualizadas
echo.

echo [3/4] 🔨 Compilando executável com PyInstaller...
echo    (Isso pode demorar alguns minutos...)
echo.

pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "VerificadorINCRA" ^
    --add-binary "poppler/Library/bin;poppler/bin" ^
    --add-data "process_memorial_descritivo_v2.py;." ^
    --hidden-import=PIL ^
    --hidden-import=PIL._tkinter_finder ^
    --hidden-import=google.generativeai ^
    --hidden-import=openpyxl ^
    --hidden-import=PyPDF2 ^
    --hidden-import=pdf2image ^
    --exclude-module=matplotlib ^
    --exclude-module=numpy ^
    verificador_georreferenciamento.py

if errorlevel 1 (
    echo.
    echo ❌ Erro durante a compilação!
    pause
    exit /b 1
)

echo.
echo ✅ Compilação concluída com sucesso!
echo.

echo [4/4] 📊 Informações do Build:
echo.
if exist "dist\VerificadorINCRA.exe" (
    for %%A in ("dist\VerificadorINCRA.exe") do (
        echo    📁 Local: dist\VerificadorINCRA.exe
        echo    📏 Tamanho: %%~zA bytes
    )
    echo.
    echo ========================================
    echo  ✅ BUILD CONCLUÍDO COM SUCESSO!
    echo ========================================
    echo.
    echo O executável está pronto em: dist\VerificadorINCRA.exe
    echo.
    echo Você pode distribuir apenas esse arquivo .exe
    echo O Poppler já está incluído internamente.
    echo.
) else (
    echo ❌ Executável não foi criado!
)

pause
