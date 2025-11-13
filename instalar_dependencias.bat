@echo off
chcp 65001 >nul
echo ========================================
echo  📦 Instalação de Dependências
echo  Verificador INCRA Pro
echo ========================================
echo.

echo [1/12] Atualizando pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo ❌ Erro ao atualizar pip!
    pause
    exit /b 1
)
echo ✅ pip atualizado
echo.

echo [2/12] Instalando pdf2image...
pip install pdf2image==1.16.3
if errorlevel 1 (
    echo ❌ Erro ao instalar pdf2image!
    pause
    exit /b 1
)
echo ✅ pdf2image instalado
echo.

echo [3/12] Instalando PyPDF2...
pip install PyPDF2==3.0.1
if errorlevel 1 (
    echo ❌ Erro ao instalar PyPDF2!
    pause
    exit /b 1
)
echo ✅ PyPDF2 instalado
echo.

echo [4/12] Instalando Pillow...
pip install Pillow==10.2.0
if errorlevel 1 (
    echo ❌ Erro ao instalar Pillow!
    pause
    exit /b 1
)
echo ✅ Pillow instalado
echo.

echo [5/12] Instalando openpyxl...
pip install openpyxl==3.1.2
if errorlevel 1 (
    echo ❌ Erro ao instalar openpyxl!
    pause
    exit /b 1
)
echo ✅ openpyxl instalado
echo.

echo [6/12] Instalando python-docx...
pip install python-docx==1.1.0
if errorlevel 1 (
    echo ❌ Erro ao instalar python-docx!
    pause
    exit /b 1
)
echo ✅ python-docx instalado
echo.

echo [7/12] Instalando google-generativeai...
pip install google-generativeai==0.3.2
if errorlevel 1 (
    echo ❌ Erro ao instalar google-generativeai!
    pause
    exit /b 1
)
echo ✅ google-generativeai instalado
echo.

echo [8/12] Instalando google-api-core...
pip install google-api-core==2.15.0
if errorlevel 1 (
    echo ❌ Erro ao instalar google-api-core!
    pause
    exit /b 1
)
echo ✅ google-api-core instalado
echo.

echo [9/12] Instalando google-auth...
pip install google-auth==2.26.2
if errorlevel 1 (
    echo ❌ Erro ao instalar google-auth!
    pause
    exit /b 1
)
echo ✅ google-auth instalado
echo.

echo [10/12] Instalando googleapis-common-protos...
pip install googleapis-common-protos==1.62.0
if errorlevel 1 (
    echo ❌ Erro ao instalar googleapis-common-protos!
    pause
    exit /b 1
)
echo ✅ googleapis-common-protos instalado
echo.

echo [11/12] Instalando protobuf...
pip install protobuf==4.25.2
if errorlevel 1 (
    echo ❌ Erro ao instalar protobuf!
    pause
    exit /b 1
)
echo ✅ protobuf instalado
echo.

echo [12/12] Verificando instalação...
python -c "import pdf2image, PIL, google.generativeai, openpyxl, PyPDF2, docx; print('✅ TODAS AS DEPENDÊNCIAS INSTALADAS COM SUCESSO!')" 2>nul
if errorlevel 1 (
    echo ❌ Algumas dependências não foram instaladas corretamente!
    echo.
    echo Execute novamente ou instale manualmente seguindo INSTALACAO_MANUAL.md
    pause
    exit /b 1
)
echo.

echo ========================================
echo  ✅ INSTALAÇÃO CONCLUÍDA!
echo ========================================
echo.
echo Todas as dependências Python foram instaladas.
echo.
echo ⚠️ IMPORTANTE: Não se esqueça de instalar o Poppler!
echo    Veja as instruções em: INSTALACAO_MANUAL.md (Passo 3)
echo.
echo Para executar o aplicativo:
echo    python verificador_georreferenciamento.py
echo.
pause
