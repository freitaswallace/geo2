# 🚀 Instalação Rápida - Verificador INCRA Pro

## Para Usuários (Apenas Executar)

Se você recebeu o arquivo `VerificadorINCRA.exe`:
1. Simplesmente **execute o arquivo .exe**
2. Não precisa instalar nada!
3. Tudo já está incluído no executável

---

## Para Desenvolvedores (Executar o Código Python)

### 📋 Pré-requisitos

- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

### 🔧 Passo 1: Instalar Python

**Windows:**
1. Baixe em: https://www.python.org/downloads/
2. Durante a instalação, marque "Add Python to PATH"

**Linux:**
```bash
sudo apt-get update
sudo apt-get install python3 python3-pip
```

**macOS:**
```bash
brew install python3
```

### 📦 Passo 2: Instalar Dependências Python

Na pasta do projeto, execute:

```bash
pip install -r requirements.txt
```

### 🔨 Passo 3: Instalar Poppler

**Windows:**
1. Baixe: https://github.com/oschwartz10612/poppler-windows/releases/
2. Extraia em `C:\poppler`
3. Adicione `C:\poppler\Library\bin` ao PATH do sistema

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install poppler-utils
```

**macOS:**
```bash
brew install poppler
```

### ▶️ Passo 4: Executar o Aplicativo

```bash
python verificador_georreferenciamento.py
```

---

## 🏗️ Para Criar o Executável (.exe)

Se você quer criar o arquivo .exe:

### 1. Baixar Poppler

Baixe e extraia na pasta do projeto (veja `COMO_COMPILAR.txt`)

### 2. Instalar PyInstaller

```bash
pip install pyinstaller
```

### 3. Executar Build

```bash
build_exe.bat
```

O executável estará em `dist/VerificadorINCRA.exe`

---

## 📚 Documentação Completa

- **COMO_COMPILAR.txt** - Guia rápido para criar .exe
- **BUILD_INSTRUCTIONS.md** - Guia detalhado de compilação
- **requirements.txt** - Lista completa de dependências

---

## 🐛 Problemas Comuns

### ModuleNotFoundError

**Problema:** `ModuleNotFoundError: No module named 'X'`

**Solução:**
```bash
pip install -r requirements.txt --upgrade
```

### Poppler não encontrado

**Problema:** `PDFInfoNotInstalledError` ou similar

**Solução:** Instale o Poppler (veja Passo 3 acima)

### Erro de permissão no Windows

**Problema:** Antivírus bloqueia o .exe

**Solução:** Adicione uma exceção no antivírus para a pasta do projeto

---

## ✅ Verificar Instalação

Para verificar se tudo está correto, execute:

```bash
python -c "import pdf2image, PIL, google.generativeai, openpyxl, PyPDF2, docx; print('✅ Todas as dependências instaladas!')"
```

Se aparecer "✅ Todas as dependências instaladas!", está tudo pronto!

---

## 📞 Suporte

Se encontrar problemas, verifique:
1. Versão do Python: `python --version` (deve ser 3.8+)
2. pip atualizado: `pip install --upgrade pip`
3. Variáveis de ambiente (PATH)

---

**Última atualização:** 2024
**Versão:** 4.0
