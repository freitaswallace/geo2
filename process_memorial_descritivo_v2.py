#!/usr/bin/env python3
"""
Script para automatizar o processamento de Memoriais Descritivos

Modos de operação:
1. Modo Normal: Processa PDF fornecido pelo usuário
2. Modo Prenotação INCRA: Busca automática em rede e conversão TIFF→PDF

Requisitos:
- pip install google-generativeai openpyxl python-docx pillow pdf2image --break-system-packages
"""

import os
import sys
import json
import shutil
import math
from pathlib import Path
from typing import Optional, Dict, List
import subprocess
import platform

# Importações das bibliotecas necessárias
try:
    import google.generativeai as genai
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from PIL import Image
    from pdf2image import convert_from_path
except ImportError as e:
    print(f"❌ Erro: Biblioteca necessária não encontrada - {e}")
    print("\n📦 Instale as dependências com:")
    print("pip install google-generativeai openpyxl python-docx pillow pdf2image --break-system-packages")
    sys.exit(1)

# Configurar para esconder janelas do CMD no Windows
if platform.system() == 'Windows':
    # Monkey patch para pdf2image não mostrar janelas do CMD
    original_popen = subprocess.Popen

    def no_console_popen(*args, **kwargs):
        """Wrapper do Popen que esconde janelas do console no Windows."""
        if 'startupinfo' not in kwargs:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            kwargs['startupinfo'] = startupinfo
        return original_popen(*args, **kwargs)

    subprocess.Popen = no_console_popen


# ============================================================================
# CONFIGURAÇÕES GLOBAIS
# ============================================================================

# Configurações do INCRA
INCRA_CONFIG = {
    'base_path': r'\\192.168.20.100\trabalho\TRABALHO\IMAGENS\IMOVEIS\DOCUMENTOS - DIVERSOS',
    'folder_interval': 1000,  # Intervalo de agrupamento (1000 em 1000)
    'identificador_inicio': [
        'MINISTÉRIO DA AGRICULTURA, PECUÁRIA E ABASTECIMENTO',
        'INSTITUTO NACIONAL DE COLONIZAÇÃO E REFORMA AGRÁRIA',
        'MEMORIAL DESCRITIVO'
    ],
    'marcador_tabela': 'DESCRIÇÃO DA PARCELA',
    'marcador_azimutes': 'Azimutes: Azimutes geodésicos',
    'colunas_vertice': ['Código', 'Longitude', 'Latitude', 'Altitude'],
    'colunas_segmento': ['Código', 'Azimute', 'Dist.', 'Confrontações']
}


# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def formatar_prenotacao(numero: str) -> str:
    """
    Formata número de prenotação para 8 dígitos com zeros à esquerda
    
    Args:
        numero: Número da prenotação (com ou sem zeros à esquerda)
    
    Returns:
        Número formatado com 8 dígitos
    """
    # Remove espaços e zeros à esquerda, depois formata
    numero_limpo = str(int(numero.strip()))
    return numero_limpo.zfill(8)


def calcular_pasta_milhar(prenotacao: str) -> str:
    """
    Calcula a pasta de milhar onde o arquivo está armazenado
    
    Args:
        prenotacao: Número da prenotação formatado (8 dígitos)
    
    Returns:
        Nome da pasta de milhar (ex: '00230000')
    """
    numero = int(prenotacao)
    milhar_superior = math.ceil(numero / 1000) * 1000
    return str(milhar_superior).zfill(8)


def testar_acesso_rede() -> bool:
    """
    Testa se o caminho de rede do INCRA está acessível
    
    Returns:
        True se acessível, False caso contrário
    """
    base_path = INCRA_CONFIG['base_path']
    
    print(f"\n🔌 Testando acesso à rede...")
    print(f"📂 Caminho: {base_path}")
    
    try:
        # Tenta acessar diretamente com os.scandir (mais compatível com UNC)
        with os.scandir(base_path) as entries:
            # Conta quantas pastas existem
            dirs = [entry.name for entry in entries if entry.is_dir()]
            
            print(f"✅ Rede acessível!")
            print(f"📁 Encontradas {len(dirs)} pastas na rede")
            
            # Mostra algumas pastas como exemplo
            if dirs:
                exemplos = sorted(dirs)[:5]
                print(f"   Exemplos: {', '.join(exemplos)}")
            
            return True
            
    except PermissionError:
        print(f"❌ Acesso negado!")
        print(f"\n💡 Possíveis causas:")
        print(f"   1. Sem permissões de leitura")
        print(f"   2. Credenciais de rede necessárias")
        print(f"   3. Compartilhamento requer autenticação")
        print(f"\n🔧 Solução:")
        print(f"   Abra o Explorer e acesse primeiro:")
        print(f"   {base_path}")
        print(f"   Depois tente novamente o script.")
        return False
        
    except FileNotFoundError:
        print(f"❌ Caminho não encontrado!")
        print(f"\n💡 Possíveis causas:")
        print(f"   1. Servidor offline")
        print(f"   2. Caminho incorreto")
        print(f"   3. Rede desconectada")
        print(f"\n🔧 Teste no CMD:")
        print(f"   dir \"{base_path}\"")
        return False
        
    except OSError as e:
        print(f"❌ Erro ao acessar rede: {e}")
        print(f"\n💡 Possíveis causas:")
        print(f"   1. Timeout de rede")
        print(f"   2. Firewall bloqueando")
        print(f"   3. Protocolo SMB desabilitado")
        return False
    
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")
        return False


def buscar_arquivo_incra(prenotacao: str) -> Optional[Path]:
    """
    Busca arquivo TIFF da prenotação na rede do INCRA
    
    Args:
        prenotacao: Número da prenotação (com ou sem formatação)
    
    Returns:
        Path do arquivo se encontrado, None caso contrário
    """
    # Formata prenotação
    prenotacao_formatada = formatar_prenotacao(prenotacao)
    print(f"🔍 Buscando prenotação: {prenotacao_formatada}")
    
    # Calcula pasta de milhar
    pasta_milhar = calcular_pasta_milhar(prenotacao_formatada)
    print(f"📁 Pasta calculada: {pasta_milhar}")
    
    # Monta caminho completo como string (melhor para UNC)
    base_path = INCRA_CONFIG['base_path']
    
    # Garante que não tem barra no final
    if base_path.endswith('\\') or base_path.endswith('/'):
        base_path = base_path[:-1]
    
    # Monta caminho da pasta e do arquivo
    pasta_completa = f"{base_path}\\{pasta_milhar}"
    arquivo_completo = f"{pasta_completa}\\{prenotacao_formatada}.tif"
    
    print(f"📂 Caminho: {arquivo_completo}")
    
    # Verifica se arquivo existe
    try:
        # Método 1: Tenta acessar diretamente o arquivo
        if os.path.isfile(arquivo_completo):
            print(f"✅ Arquivo encontrado!")
            return Path(arquivo_completo)
        
        # Método 2: Se não encontrou, lista a pasta para debug
        print(f"❌ Arquivo não encontrado diretamente")
        
        if not os.path.isdir(pasta_completa):
            print(f"❌ Pasta não existe: {pasta_milhar}")
            return None
        
        print(f"📁 Pasta existe, listando arquivos...")
        
        # Lista arquivos .tif na pasta
        arquivos_tif = []
        with os.scandir(pasta_completa) as entries:
            for entry in entries:
                if entry.is_file() and entry.name.lower().endswith('.tif'):
                    arquivos_tif.append(entry.name)
        
        if arquivos_tif:
            print(f"   Encontrados {len(arquivos_tif)} arquivos .tif na pasta")
            # Mostra alguns exemplos
            for arq in sorted(arquivos_tif)[:5]:
                print(f"   - {arq}")
            if len(arquivos_tif) > 5:
                print(f"   ... e mais {len(arquivos_tif) - 5} arquivos")
            
            # Procura o arquivo específico na lista
            nome_procurado = f"{prenotacao_formatada}.tif"
            if nome_procurado.upper() in [a.upper() for a in arquivos_tif]:
                print(f"✅ Arquivo encontrado na listagem!")
                return Path(arquivo_completo)
        else:
            print(f"   ⚠️ Pasta vazia ou sem arquivos .tif")
        
        return None
        
    except PermissionError as e:
        print(f"❌ Acesso negado: {e}")
        print(f"💡 Abra a pasta no Explorer primeiro:")
        print(f"   {pasta_completa}")
        return None
        
    except Exception as e:
        print(f"❌ Erro ao buscar arquivo: {e}")
        return None


def copiar_para_downloads(arquivo_origem: Path, prenotacao: str) -> Path:
    """
    Copia arquivo TIFF para pasta Tabelas_Incra em Documentos

    Args:
        arquivo_origem: Path do arquivo original (pode ser string UNC)
        prenotacao: Número da prenotação formatado

    Returns:
        Path do arquivo copiado
    """
    # Determina pasta Documentos/Tabelas_Incra do usuário
    home = Path.home()
    tabelas_incra = home / 'Documents' / 'Tabelas_Incra'

    # Cria pasta base se não existir
    tabelas_incra.mkdir(parents=True, exist_ok=True)

    # Cria subpasta específica para esta prenotação
    pasta_prenotacao = tabelas_incra / f'Prenotacao_{prenotacao}'
    pasta_prenotacao.mkdir(parents=True, exist_ok=True)

    print(f"📁 Pasta criada: {pasta_prenotacao}")
    
    # Nome do arquivo
    nome_arquivo = os.path.basename(str(arquivo_origem))
    destino = pasta_prenotacao / nome_arquivo
    
    print(f"📋 Copiando arquivo...")
    print(f"   Origem: {arquivo_origem}")
    print(f"   Destino: {destino}")
    
    try:
        # Usa shutil.copy2 com strings para melhor compatibilidade UNC
        shutil.copy2(str(arquivo_origem), str(destino))
        print(f"✅ Arquivo copiado com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao copiar arquivo: {e}")
        print(f"⚠️ Tentando método alternativo...")
        
        # Método alternativo: lê e escreve byte a byte
        try:
            with open(str(arquivo_origem), 'rb') as f_origem:
                conteudo = f_origem.read()
            with open(str(destino), 'wb') as f_destino:
                f_destino.write(conteudo)
            print(f"✅ Arquivo copiado com método alternativo!")
        except Exception as e2:
            print(f"❌ Erro no método alternativo: {e2}")
            raise Exception(f"Não foi possível copiar o arquivo: {e2}")
    
    return destino


def converter_tiff_para_pdf(tiff_path: Path) -> Path:
    """
    Converte arquivo TIFF (multi-página) para PDF
    
    Args:
        tiff_path: Path do arquivo TIFF
    
    Returns:
        Path do arquivo PDF gerado
    """
    print(f"\n🔄 Convertendo TIFF para PDF...")
    
    pdf_path = tiff_path.with_suffix('.pdf')
    
    try:
        # Abre arquivo TIFF
        img = Image.open(tiff_path)
        
        # Lista para armazenar todas as páginas
        images = []
        
        # Itera por todas as páginas do TIFF
        try:
            page = 0
            while True:
                img.seek(page)
                # Converte para RGB se necessário
                if img.mode != 'RGB':
                    rgb_img = img.convert('RGB')
                else:
                    rgb_img = img.copy()
                images.append(rgb_img)
                page += 1
                print(f"  📄 Página {page} processada")
        except EOFError:
            pass  # Fim do arquivo TIFF
        
        # Salva como PDF multi-página
        if images:
            images[0].save(
                pdf_path,
                save_all=True,
                append_images=images[1:] if len(images) > 1 else [],
                resolution=100.0,
                quality=95
            )
            print(f"✅ PDF criado: {pdf_path.name} ({len(images)} páginas)")
            return pdf_path
        else:
            raise ValueError("Nenhuma página encontrada no TIFF")
            
    except Exception as e:
        print(f"❌ Erro ao converter TIFF: {e}")
        raise


def extrair_memorial_incra(pdf_path: Path, api_key: str) -> Dict:
    """
    Extrai tabela de Memorial Descritivo do INCRA
    
    Esta função procura especificamente pelo formato do INCRA e extrai
    a tabela de coordenadas que pode estar em múltiplas páginas
    
    Args:
        pdf_path: Path do arquivo PDF
        api_key: Chave da API do Gemini
    
    Returns:
        Dados estruturados da tabela
    """
    print(f"\n📊 Extraindo Memorial Descritivo do INCRA...")
    
    # Configura API
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash-lite')
    
    # Carrega PDF
    with open(pdf_path, 'rb') as f:
        pdf_data = f.read()
    
    # Prompt especializado para Memorial do INCRA
    prompt = f"""
Você está processando um Memorial Descritivo do INCRA (Instituto Nacional de Colonização e Reforma Agrária).

INSTRUÇÕES CRÍTICAS:

1. LOCALIZAÇÃO: Encontre o bloco que contém:
   - "MINISTÉRIO DA AGRICULTURA, PECUÁRIA E ABASTECIMENTO"
   - "INSTITUTO NACIONAL DE COLONIZAÇÃO E REFORMA AGRÁRIA"
   - "MEMORIAL DESCRITIVO"

2. TABELA: Após encontrar "DESCRIÇÃO DA PARCELA" e "Azimutes: Azimutes geodésicos", 
   localize a tabela com as seguintes colunas:

   VÉRTICE (4 colunas):
   - Código
   - Longitude
   - Latitude  
   - Altitude (m)

   SEGMENTO VANTE (4 colunas):
   - Código
   - Azimute
   - Dist. (m)
   - Confrontações

3. MULTI-PÁGINA: A tabela pode continuar em múltiplas páginas. Continue extraindo 
   até encontrar um novo cabeçalho de seção (como "CERTIFICAÇÃO") ou o fim da tabela.

4. FORMATO DE SAÍDA: Retorne APENAS o JSON neste formato exato:

{{
  "header_row1": ["VÉRTICE", "SEGMENTO VANTE"],
  "header_row2": ["Código", "Longitude", "Latitude", "Altitude (m)", "Código", "Azimute", "Dist. (m)", "Confrontações"],
  "data": [
    ["valor1", "valor2", "valor3", "valor4", "valor5", "valor6", "valor7", "valor8"],
    ...
  ]
}}

IMPORTANTE:
- Mantenha a formatação exata dos valores (graus, aspas, vírgulas)
- Se um campo estiver vazio, use ""
- Extraia TODAS as linhas da tabela de TODAS as páginas
- Retorne APENAS o JSON, sem texto adicional
"""
    
    print("🤖 Enviando para Gemini API...")
    response = model.generate_content([
        prompt,
        {"mime_type": "application/pdf", "data": pdf_data}
    ])
    
    print("✅ Resposta recebida")
    
    # Processa resposta
    response_text = response.text.strip()
    
    # Remove marcadores markdown se presentes
    if response_text.startswith("```json"):
        response_text = response_text[7:]
    if response_text.startswith("```"):
        response_text = response_text[3:]
    if response_text.endswith("```"):
        response_text = response_text[:-3]
    
    # Extrai JSON
    if '{' in response_text:
        response_text = response_text[response_text.find('{'):]
    if '}' in response_text:
        response_text = response_text[:response_text.rfind('}')+1]
    
    response_text = response_text.strip()
    
    # Parse JSON
    try:
        table_data = json.loads(response_text)
        num_linhas = len(table_data.get('data', []))
        print(f"✅ Tabela extraída: {num_linhas} linhas de dados")
        return table_data
    except json.JSONDecodeError as e:
        print(f"❌ Erro ao fazer parse do JSON: {e}")
        print(f"Resposta recebida (primeiros 500 chars): {response_text[:500]}")
        raise


# ============================================================================
# FUNÇÕES PRINCIPAIS DO MODO NORMAL
# ============================================================================

def configure_gemini_api(api_key):
    """Configura a API do Google Gemini"""
    genai.configure(api_key=api_key)


def extract_table_from_pdf(pdf_path, api_key):
    """Extrai dados da tabela do PDF usando Google Gemini API (modo normal)"""
    print(f"📄 Processando PDF: {pdf_path}")
    
    configure_gemini_api(api_key)
    model = genai.GenerativeModel('gemini-2.5-flash-lite')
    
    with open(pdf_path, 'rb') as f:
        pdf_data = f.read()
    
    prompt = """
Analise este Memorial Descritivo e extraia APENAS a tabela principal que contém informações de vértices.

A tabela tem a seguinte estrutura:
- Cabeçalho Linha 1: "VÉRTICE" (colunas A-D) e "SEGMENTO VANTE" (colunas E-H)
- Cabeçalho Linha 2: Código, Longitude, Latitude, Altitude (m), Código, Azimute, Dist. (m), Confrontações

Retorne os dados em formato JSON seguindo EXATAMENTE esta estrutura:
{
  "header_row1": ["VÉRTICE", "SEGMENTO VANTE"],
  "header_row2": ["Código", "Longitude", "Latitude", "Altitude (m)", "Código", "Azimute", "Dist. (m)", "Confrontações"],
  "data": [
    ["valor1", "valor2", "valor3", "valor4", "valor5", "valor6", "valor7", "valor8"],
    ...
  ]
}

IMPORTANTE: 
- Retorne APENAS o JSON, sem texto adicional
- Inclua TODOS os dados da tabela
- Mantenha a formatação original dos valores
- Se um campo estiver vazio, use ""
"""
    
    print("🤖 Enviando para Gemini API...")
    response = model.generate_content([prompt, {"mime_type": "application/pdf", "data": pdf_data}])
    
    print("✅ Resposta recebida da API")
    
    response_text = response.text.strip()
    
    if response_text.startswith("```json"):
        response_text = response_text[7:]
    if response_text.startswith("```"):
        response_text = response_text[3:]
    if response_text.endswith("```"):
        response_text = response_text[:-3]
    
    response_text = response_text.strip()
    
    try:
        table_data = json.loads(response_text)
        print(f"📊 Tabela extraída: {len(table_data.get('data', []))} linhas de dados")
        return table_data
    except json.JSONDecodeError as e:
        print(f"❌ Erro ao fazer parse do JSON: {e}")
        print(f"Resposta recebida: {response_text[:500]}...")
        sys.exit(1)


def create_excel_file(table_data, output_path):
    """Cria arquivo Excel com a tabela formatada"""
    print(f"\n📊 Criando arquivo Excel: {output_path}")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Memorial Descritivo"
    
    header_font = Font(bold=True, size=11)
    center_alignment = Alignment(horizontal='center', vertical='center')
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Linha 1: Cabeçalhos mesclados
    ws.merge_cells('A1:D1')
    cell_a1 = ws['A1']
    cell_a1.value = "VÉRTICE"
    cell_a1.font = header_font
    cell_a1.alignment = center_alignment
    cell_a1.border = border_style
    
    ws.merge_cells('E1:H1')
    cell_e1 = ws['E1']
    cell_e1.value = "SEGMENTO VANTE"
    cell_e1.font = header_font
    cell_e1.alignment = center_alignment
    cell_e1.border = border_style
    
    # Linha 2: Sub-cabeçalhos
    header_row2 = table_data.get('header_row2', [])
    for col_idx, header in enumerate(header_row2, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.alignment = center_alignment
        cell.border = border_style
    
    # Linhas 3+: Dados
    data_rows = table_data.get('data', [])
    for row_idx, row_data in enumerate(data_rows, start=3):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.border = border_style
            if col_idx in [1, 5]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Ajusta larguras
    column_widths = {
        'A': 15, 'B': 18, 'C': 18, 'D': 15,
        'E': 15, 'F': 15, 'G': 12, 'H': 30
    }
    
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
    
    wb.save(output_path)
    print(f"✅ Excel criado com sucesso!")
    return output_path


def create_word_file(table_data, output_path):
    """Cria arquivo Word com a tabela formatada"""
    print(f"\n📝 Criando arquivo Word: {output_path}")
    
    doc = Document()

    data_rows = table_data.get('data', [])
    num_rows = 2 + len(data_rows)
    num_cols = 8
    
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'
    
    # Linha 1: Mesclagem
    cell_a1 = table.rows[0].cells[0]
    cell_d1 = table.rows[0].cells[3]
    cell_a1.merge(cell_d1)
    cell_a1.text = "VÉRTICE"
    cell_a1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cell_a1.paragraphs[0].runs[0]
    run.bold = True
    run.font.size = Pt(11)
    
    cell_e1 = table.rows[0].cells[4]
    cell_h1 = table.rows[0].cells[7]
    cell_e1.merge(cell_h1)
    cell_e1.text = "SEGMENTO VANTE"
    cell_e1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cell_e1.paragraphs[0].runs[0]
    run.bold = True
    run.font.size = Pt(11)
    
    # Linha 2: Sub-cabeçalhos
    header_row2 = table_data.get('header_row2', [])
    for col_idx, header in enumerate(header_row2):
        cell = table.rows[1].cells[col_idx]
        cell.text = header
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(10)
    
    # Linhas 3+: Dados
    for row_idx, row_data in enumerate(data_rows, start=2):
        for col_idx, value in enumerate(row_data):
            cell = table.rows[row_idx].cells[col_idx]
            cell.text = str(value) if value else ""
            if col_idx in [0, 4]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].font.size = Pt(9)
    
    # AutoFit
    table.autofit = True
    table.allow_autofit = True
    
    # Larguras preferenciais
    preferred_widths = [
        Cm(2.5), Cm(3.5), Cm(3.5), Cm(2.5),
        Cm(2.5), Cm(2.5), Cm(2.0), Cm(4.0)
    ]
    
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx < len(preferred_widths):
                cell.width = preferred_widths[idx]
    
    doc.save(output_path)
    print(f"✅ Word criado com sucesso!")
    return output_path


# ============================================================================
# FUNÇÃO PRINCIPAL - MODO PRENOTAÇÃO INCRA
# ============================================================================

def modo_prenotacao_incra(api_key: str):
    """
    Modo de operação para prenotações do INCRA
    
    Fluxo completo:
    1. Testa acesso à rede
    2. Solicita número da prenotação
    3. Busca arquivo TIFF na rede
    4. Copia para Downloads
    5. Converte TIFF para PDF
    6. Extrai tabela do Memorial
    7. Oferece opção de gerar Excel e/ou Word
    """
    print("\n" + "="*70)
    print("🏛️  MODO PRENOTAÇÃO INCRA")
    print("="*70)
    
    # 0. Testa acesso à rede primeiro
    if not testar_acesso_rede():
        print("\n" + "="*70)
        print("❌ ERRO: Não foi possível acessar a rede do INCRA!")
        print("="*70)
        print("\n📝 Verifique:")
        print("  1. Conexão com a rede")
        print("  2. Caminho configurado (linha ~30 do script)")
        print(f"     Atual: {INCRA_CONFIG['base_path']}")
        print("  3. Permissões de acesso")
        print("  4. VPN ativa (se necessário)")
        return
    
    # 1. Solicita prenotação
    prenotacao = input("\n📋 Digite o número da Prenotação (ex: 229885 ou 00229885): ").strip()
    
    if not prenotacao:
        print("❌ Número de prenotação não fornecido!")
        return
    
    try:
        prenotacao_formatada = formatar_prenotacao(prenotacao)
    except ValueError:
        print("❌ Número de prenotação inválido!")
        return
    
    print(f"✅ Prenotação formatada: {prenotacao_formatada}")
    
    # 2. Busca arquivo TIFF
    print("\n" + "-"*70)
    arquivo_tiff = buscar_arquivo_incra(prenotacao_formatada)
    
    if not arquivo_tiff:
        print("❌ Arquivo não encontrado na rede do INCRA!")
        print(f"📂 Caminho esperado: {INCRA_CONFIG['base_path']}\\{calcular_pasta_milhar(prenotacao_formatada)}\\{prenotacao_formatada}.tif")
        return
    
    # 3. Copia para Downloads
    print("\n" + "-"*70)
    arquivo_local = copiar_para_downloads(arquivo_tiff, prenotacao_formatada)
    
    # 4. Converte TIFF para PDF
    print("\n" + "-"*70)
    try:
        arquivo_pdf = converter_tiff_para_pdf(arquivo_local)
    except Exception as e:
        print(f"❌ Erro na conversão: {e}")
        return
    
    # 5. Extrai tabela
    print("\n" + "-"*70)
    try:
        table_data = extrair_memorial_incra(arquivo_pdf, api_key)
    except Exception as e:
        print(f"❌ Erro na extração: {e}")
        return
    
    # 6. Oferece opções de geração
    print("\n" + "="*70)
    print("📊 EXTRAÇÃO CONCLUÍDA!")
    print(f"✅ {len(table_data.get('data', []))} linhas extraídas")
    print("="*70)
    
    escolher_arquivos_saida(table_data, arquivo_pdf.parent, prenotacao_formatada)


# ============================================================================
# FUNÇÃO PRINCIPAL - MODO NORMAL
# ============================================================================

def modo_normal(api_key: str):
    """Modo de operação normal (arquivo PDF fornecido pelo usuário)"""
    print("\n" + "="*70)
    print("📄 MODO NORMAL - Processar PDF")
    print("="*70)
    
    pdf_path = input("\n📂 Digite o caminho completo do arquivo PDF: ").strip()
    pdf_path = pdf_path.strip("'\"")
    
    if not os.path.exists(pdf_path):
        print(f"\n❌ Erro: Arquivo não encontrado: {pdf_path}")
        return
    
    try:
        table_data = extract_table_from_pdf(pdf_path, api_key)
    except Exception as e:
        print(f"\n❌ Erro ao processar PDF: {e}")
        return
    
    output_dir = Path(pdf_path).parent
    
    print("\n" + "="*70)
    print("📊 EXTRAÇÃO CONCLUÍDA!")
    print(f"✅ {len(table_data.get('data', []))} linhas extraídas")
    print("="*70)
    
    escolher_arquivos_saida(table_data, output_dir)


# ============================================================================
# FUNÇÃO DE ESCOLHA DE ARQUIVOS DE SAÍDA
# ============================================================================

def escolher_arquivos_saida(table_data: Dict, output_dir: Path, prefixo: str = "output"):
    """
    Permite ao usuário escolher quais arquivos gerar (Excel e/ou Word)
    
    Args:
        table_data: Dados extraídos da tabela
        output_dir: Diretório onde salvar os arquivos
        prefixo: Prefixo para nome dos arquivos
    """
    print("\n" + "="*70)
    print("💾 ESCOLHA OS ARQUIVOS DE SAÍDA")
    print("="*70)
    print("\nQuais arquivos você deseja gerar?")
    print("  1 - Apenas Excel (.xlsx)")
    print("  2 - Apenas Word (.docx)")
    print("  3 - Ambos (Excel + Word)")
    print("  0 - Cancelar (não gerar nenhum)")
    
    while True:
        escolha = input("\n👉 Digite sua escolha (0-3): ").strip()
        
        if escolha == '0':
            print("\n❌ Operação cancelada. Nenhum arquivo foi gerado.")
            return
        
        elif escolha == '1':
            # Apenas Excel
            excel_path = output_dir / f"{prefixo}.xlsx"
            create_excel_file(table_data, str(excel_path))
            print(f"\n✅ Arquivo gerado:")
            print(f"   📊 {excel_path}")
            break
        
        elif escolha == '2':
            # Apenas Word
            word_path = output_dir / f"{prefixo}.docx"
            create_word_file(table_data, str(word_path))
            print(f"\n✅ Arquivo gerado:")
            print(f"   📝 {word_path}")
            break
        
        elif escolha == '3':
            # Ambos
            excel_path = output_dir / f"{prefixo}.xlsx"
            word_path = output_dir / f"{prefixo}.docx"
            create_excel_file(table_data, str(excel_path))
            create_word_file(table_data, str(word_path))
            print(f"\n✅ Arquivos gerados:")
            print(f"   📊 {excel_path}")
            print(f"   📝 {word_path}")
            break
        
        else:
            print("❌ Opção inválida! Digite 0, 1, 2 ou 3.")


# ============================================================================
# FUNÇÃO MAIN
# ============================================================================

def main():
    """Função principal que coordena o fluxo completo"""
    print("="*70)
    print("🚀 PROCESSADOR DE MEMORIAL DESCRITIVO")
    print("="*70)
    
    # Configurar API Key (fixa)
    print("\n🔑 Configuração da API do Google Gemini")
    api_key = 'AIzaSyAdA_GO7cQ0m1ouie4wGwXf4a4SnHKjBh8'
    print(f"✅ Usando chave configurada")
    
    # Escolher modo de operação
    print("\n" + "="*70)
    print("🎯 ESCOLHA O MODO DE OPERAÇÃO")
    print("="*70)
    print("\n  1 - Modo Normal (fornecer arquivo PDF)")
    print("  2 - Modo Prenotação INCRA (busca automática)")
    
    while True:
        modo = input("\n👉 Digite sua escolha (1 ou 2): ").strip()
        
        if modo == '1':
            modo_normal(api_key)
            break
        elif modo == '2':
            modo_prenotacao_incra(api_key)
            break
        else:
            print("❌ Opção inválida! Digite 1 ou 2.")
    
    print("\n" + "="*70)
    print("✨ PROCESSAMENTO FINALIZADO!")
    print("="*70)


if __name__ == "__main__":
    main()