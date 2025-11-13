# 📦 Como Compilar o Verificador INCRA Pro para .EXE

Este guia mostra como criar um executável Windows (.exe) com o Poppler incluído.

---

## 🔧 Pré-requisitos

1. **Python 3.8+** instalado
2. **PyInstaller** (será instalado automaticamente pelo script)
3. **Poppler para Windows**

---

## 📥 Passo 1: Baixar o Poppler

1. Acesse: https://github.com/oschwartz10612/poppler-windows/releases/

2. Baixe a versão mais recente (arquivo `.zip`):
   - Exemplo: `Release-24.07.0-0.zip`

3. Extraia o arquivo ZIP

4. Renomeie a pasta extraída para `poppler`

5. Mova a pasta `poppler` para a raiz do projeto (mesma pasta onde está o `verificador_georreferenciamento.py`)

### ✅ Estrutura Esperada:

```
geo2/
├── verificador_georreferenciamento.py
├── process_memorial_descritivo_v2.py
├── build_exe.bat
├── BUILD_INSTRUCTIONS.md
└── poppler/
    └── Library/
        └── bin/
            ├── pdftoppm.exe
            ├── pdfinfo.exe
            └── ... (outros arquivos)
```

---

## 🚀 Passo 2: Executar o Build

Simplesmente execute o script de build:

```batch
build_exe.bat
```

O script irá:
1. ✅ Verificar se o PyInstaller está instalado
2. ✅ Verificar se o Poppler está na pasta correta
3. ✅ Limpar builds anteriores
4. ✅ Atualizar dependências
5. ✅ Compilar o executável com Poppler embutido

---

## 📂 Resultado

Após a compilação bem-sucedida:

```
dist/
└── VerificadorINCRA.exe  👈 Este é o seu executável!
```

**Tamanho esperado:** ~150-250 MB (inclui Poppler e todas as dependências)

---

## 🎯 Distribuição

Você pode distribuir **apenas o arquivo `VerificadorINCRA.exe`**:

- ✅ Nenhuma instalação adicional necessária
- ✅ Poppler incluído internamente
- ✅ Todas as bibliotecas Python embutidas
- ✅ Funciona em qualquer Windows 10/11

---

## 🐛 Solução de Problemas

### ❌ Erro: "Pasta 'poppler' não encontrada"

**Solução:** Certifique-se de que a pasta `poppler` está na raiz do projeto com a estrutura:
```
poppler/Library/bin/
```

### ❌ Erro: "ModuleNotFoundError" ao executar o .exe

**Solução:** Execute o build novamente. O script inclui todos os módulos necessários via `--hidden-import`.

### ❌ Executável muito grande (>300 MB)

Isso é normal! O Poppler adiciona ~80-100 MB ao executável.

### ❌ Antivírus bloqueia o executável

Executáveis criados com PyInstaller podem ser sinalizados como falsos positivos. Adicione uma exceção no antivírus.

---

## 🔧 Build Manual (Avançado)

Se preferir executar manualmente sem o script `.bat`:

```bash
# 1. Instalar PyInstaller
pip install pyinstaller

# 2. Compilar
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
    verificador_georreferenciamento.py
```

---

## 📝 Notas Técnicas

### Como o Poppler é Detectado

O código foi modificado para detectar automaticamente se está rodando como executável:

```python
def get_poppler_path():
    if getattr(sys, 'frozen', False):
        # Rodando como .exe - usa Poppler embutido
        base_path = Path(sys._MEIPASS)
        poppler_path = base_path / 'poppler' / 'bin'
        return str(poppler_path)
    else:
        # Rodando como script - usa Poppler do sistema
        return None
```

Todas as chamadas `convert_from_path()` agora incluem `poppler_path=POPPLER_PATH`.

---

## 🎉 Pronto!

Seu executável está pronto para distribuição. Teste-o em diferentes máquinas Windows para garantir compatibilidade.

### Checklist Final:

- [ ] Executável abre sem erros
- [ ] Interface gráfica é exibida corretamente
- [ ] Modo automático funciona (PDFs são processados)
- [ ] Modo manual funciona
- [ ] Relatório HTML é gerado
- [ ] Backups são salvos
- [ ] Botão "Limpar Backups" funciona

---

**Dúvidas ou problemas?** Verifique o console de erros ao executar o .exe (execute via CMD para ver mensagens).
