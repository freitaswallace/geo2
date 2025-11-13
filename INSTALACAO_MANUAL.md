# 📦 Instalação Manual - Passo a Passo

## 🐍 PASSO 1: BAIXAR E INSTALAR O PYTHON

### Versão Recomendada: **Python 3.11.9**

**Link de Download:**
```
https://www.python.org/downloads/release/python-3119/
```

### Para Windows:
1. Baixe: **Windows installer (64-bit)**
   - Link direto: https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe

2. Durante a instalação:
   - ✅ **MARQUE**: "Add Python 3.11 to PATH"
   - ✅ **MARQUE**: "Install pip"
   - Clique em "Install Now"

3. Verificar instalação:
   Abra o CMD e digite:
   ```bash
   python --version
   ```
   Deve aparecer: `Python 3.11.9`

---

## 📦 PASSO 2: INSTALAR DEPENDÊNCIAS (Uma por Uma)

Abra o **Prompt de Comando (CMD)** e execute cada comando abaixo **separadamente**:

### 1. Atualizar o pip (instalador de pacotes)
```bash
python -m pip install --upgrade pip
```

### 2. Instalar pdf2image (processamento de PDF)
```bash
pip install pdf2image==1.16.3
```

### 3. Instalar PyPDF2 (manipulação de PDF)
```bash
pip install PyPDF2==3.0.1
```

### 4. Instalar Pillow (processamento de imagens)
```bash
pip install Pillow==10.2.0
```

### 5. Instalar openpyxl (arquivos Excel)
```bash
pip install openpyxl==3.1.2
```

### 6. Instalar python-docx (arquivos Word)
```bash
pip install python-docx==1.1.0
```

### 7. Instalar Google Generative AI (Gemini)
```bash
pip install google-generativeai==0.3.2
```

### 8. Instalar dependências do Google (4 pacotes)
```bash
pip install google-api-core==2.15.0
```

```bash
pip install google-auth==2.26.2
```

```bash
pip install googleapis-common-protos==1.62.0
```

```bash
pip install protobuf==4.25.2
```

### 9. (OPCIONAL) PyInstaller - Apenas se for criar o .exe
```bash
pip install pyinstaller==6.3.0
```

---

## 🔧 PASSO 3: INSTALAR POPPLER

O Poppler é necessário para converter PDFs em imagens.

### Para Windows:

1. **Baixar Poppler:**
   ```
   https://github.com/oschwartz10612/poppler-windows/releases/
   ```
   - Baixe o arquivo: `Release-XX.XX.X-0.zip` (versão mais recente)

2. **Extrair e Configurar:**
   - Extraia o arquivo ZIP
   - Mova a pasta extraída para `C:\poppler`
   - A estrutura deve ficar: `C:\poppler\Library\bin\`

3. **Adicionar ao PATH do Windows:**

   **Método 1 (Simples - via CMD como Administrador):**
   ```bash
   setx PATH "%PATH%;C:\poppler\Library\bin" /M
   ```

   **Método 2 (Manual):**
   - Clique com botão direito em "Este Computador" → "Propriedades"
   - "Configurações avançadas do sistema"
   - "Variáveis de Ambiente"
   - Em "Variáveis do sistema", selecione "Path" → "Editar"
   - Clique "Novo" e adicione: `C:\poppler\Library\bin`
   - Clique OK em todas as janelas
   - **Reinicie o computador**

4. **Verificar:**
   Abra um novo CMD e digite:
   ```bash
   pdftoppm -h
   ```
   Se aparecer uma mensagem de ajuda, está instalado corretamente!

---

## ✅ PASSO 4: VERIFICAR SE TUDO ESTÁ FUNCIONANDO

Cole este comando no CMD:

```bash
python -c "import pdf2image, PIL, google.generativeai, openpyxl, PyPDF2, docx; print('✅ TODAS AS DEPENDÊNCIAS INSTALADAS COM SUCESSO!')"
```

Se aparecer a mensagem de sucesso, está tudo pronto! 🎉

---

## 🚀 PASSO 5: EXECUTAR O APLICATIVO

1. Navegue até a pasta do projeto:
   ```bash
   cd C:\caminho\para\pasta\geo2
   ```

2. Execute o script:
   ```bash
   python verificador_georreferenciamento.py
   ```

---

## 🐛 SOLUÇÃO DE PROBLEMAS

### Erro: "python não é reconhecido como comando"

**Solução:** Python não foi adicionado ao PATH
1. Desinstale o Python
2. Reinstale marcando "Add Python to PATH"
3. Reinicie o computador

### Erro: "PDFInfoNotInstalledError"

**Solução:** Poppler não está instalado ou não está no PATH
1. Verifique se existe: `C:\poppler\Library\bin\pdftoppm.exe`
2. Adicione ao PATH (veja Passo 3)
3. Reinicie o computador

### Erro: "ModuleNotFoundError: No module named 'X'"

**Solução:** Biblioteca não foi instalada
1. Execute novamente o comando de instalação da biblioteca específica
2. Exemplo: `pip install pdf2image==1.16.3`

### Erro: "pip não é reconhecido como comando"

**Solução:**
```bash
python -m ensurepip --upgrade
python -m pip install --upgrade pip
```

---

## 📋 CHECKLIST FINAL

Antes de executar o aplicativo, verifique:

- [ ] Python 3.11.9 instalado
- [ ] Python adicionado ao PATH
- [ ] pip atualizado
- [ ] pdf2image instalado
- [ ] PyPDF2 instalado
- [ ] Pillow instalado
- [ ] openpyxl instalado
- [ ] google-generativeai instalado
- [ ] Poppler instalado em C:\poppler
- [ ] Poppler adicionado ao PATH
- [ ] Computador reiniciado após configurar PATH
- [ ] Comando de verificação executado com sucesso

---

## 🔄 RESUMO DOS COMANDOS (Para copiar e colar sequencialmente)

```bash
python -m pip install --upgrade pip
pip install pdf2image==1.16.3
pip install PyPDF2==3.0.1
pip install Pillow==10.2.0
pip install openpyxl==3.1.2
pip install python-docx==1.1.0
pip install google-generativeai==0.3.2
pip install google-api-core==2.15.0
pip install google-auth==2.26.2
pip install googleapis-common-protos==1.62.0
pip install protobuf==4.25.2
```

---

**Boa sorte! 🚀**
