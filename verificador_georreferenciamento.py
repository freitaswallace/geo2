#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Verificador de Consistência de Documentos de Georreferenciamento
Aplicação GUI para cartórios - Análise multimodal com Gemini AI
Autor: Sistema Automatizado
Versão: 4.0 - Interface moderna com Modo Automático
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from pathlib import Path
import threading
from typing import List, Optional, Dict, Tuple
import json
import tempfile
import shutil
import webbrowser
import math
from datetime import datetime
import configparser
import subprocess
import platform

try:
    from pdf2image import convert_from_path
    from PIL import Image, ImageTk
    import google.generativeai as genai
    from openpyxl import load_workbook
    import PyPDF2
    # Importar funções de extração do script existente
    from process_memorial_descritivo_v2 import (
        extract_table_from_pdf,
        extrair_memorial_incra,
        create_excel_file
    )
except ImportError as e:
    print(f"❌ Erro: Biblioteca necessária não encontrada: {e}")
    print("\nInstale as dependências com:")

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
    print("pip install pdf2image Pillow google-generativeai openpyxl PyPDF2")
    print("\nNota: Também é necessário ter o 'poppler-utils' instalado no sistema.")
    sys.exit(1)


class ConfigManager:
    """Gerencia configurações persistentes da aplicação."""

    def __init__(self):
        self.config_dir = Path.home() / ".conferencia_geo"
        self.config_file = self.config_dir / "config.ini"
        self.config = configparser.ConfigParser()
        self._ensure_config_exists()

    def _ensure_config_exists(self):
        """Cria diretório e arquivo de configuração se não existir."""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        if not self.config_file.exists():
            self.config['API'] = {'gemini_key': ''}
            self.save()
        else:
            self.config.read(self.config_file)

    def save(self):
        """Salva configurações no arquivo."""
        with open(self.config_file, 'w') as f:
            self.config.write(f)

    def get_api_key(self) -> str:
        """Retorna a API key salva."""
        return self.config.get('API', 'gemini_key', fallback='')

    def set_api_key(self, key: str):
        """Salva a API key."""
        if 'API' not in self.config:
            self.config['API'] = {}
        self.config['API']['gemini_key'] = key
        self.save()


class VerificadorGeorreferenciamento:
    """Classe principal da aplicação de verificação de documentos."""

    def __init__(self, root):
        self.root = root
        self.root.title("✨ Verificador INCRA Pro v4.0")
        self.root.geometry("1450x980")

        # Maximizar janela ao abrir
        try:
            self.root.state('zoomed')  # Windows/Linux
        except:
            try:
                self.root.attributes('-zoomed', True)  # Alternativa
            except:
                pass  # Se não funcionar, mantém tamanho padrão

        # Gerenciador de configurações
        self.config_manager = ConfigManager()

        # Variáveis para armazenar caminhos dos arquivos
        self.incra_path = tk.StringVar()
        self.projeto_path = tk.StringVar()
        self.numero_prenotacao = tk.StringVar()
        self.modo_atual = tk.StringVar(value="automatico")

        # Variáveis para Sub-modo Por Páginas (Modo Automático)
        self.incra_paginas = tk.StringVar()  # Números de páginas do INCRA (ex: "1,2,3")
        self.projeto_paginas = tk.StringVar()  # Números de páginas do PROJETO (ex: "4,5")
        self.incra_anexo_path = tk.StringVar()  # Caminho do anexo manual do INCRA
        self.projeto_anexo_path = tk.StringVar()  # Caminho do anexo manual do PROJETO

        # Variáveis para armazenar dados extraídos
        self.incra_excel_path: Optional[str] = None
        self.projeto_excel_path: Optional[str] = None
        self.incra_data: Optional[Dict] = None
        self.projeto_data: Optional[Dict] = None

        # Variáveis para modo automático
        self.pdf_extraido_incra: Optional[str] = None
        self.pdf_extraido_projeto: Optional[str] = None
        self.preview_incra_image: Optional[Image.Image] = None
        self.preview_projeto_image: Optional[Image.Image] = None

        # Janela de progresso em tempo real
        self.progress_window: Optional[tk.Toplevel] = None
        self.progress_text: Optional[scrolledtext.ScrolledText] = None

        # Configurar estilo moderno
        self._configurar_estilo()

        # Criar interface
        self._criar_interface()

        # Carregar API key salva
        self._carregar_api_key()

        # Configurar evento de fechamento da janela
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_programa)

    @staticmethod
    def _get_startup_info():
        """Cria configuração para esconder janelas do CMD no Windows."""
        if platform.system() == 'Windows':
            import subprocess
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            return startupinfo
        return None

    def _configurar_estilo(self):
        """Configura tema moderno e profissional com cores vibrantes."""
        style = ttk.Style()
        style.theme_use('clam')

        # Paleta de cores moderna e agradável (inspirada em Material Design)
        self.colors = {
            'primary': '#6366F1',      # Indigo vibrante
            'primary_dark': '#4F46E5',
            'secondary': '#EC4899',    # Rosa vibrante
            'success': '#10B981',      # Verde esmeralda
            'warning': '#F59E0B',      # Âmbar
            'danger': '#EF4444',       # Vermelho
            'info': '#3B82F6',         # Azul
            'bg_light': '#F9FAFB',     # Cinza muito claro
            'bg_card': '#FFFFFF',
            'text_dark': '#1F2937',
            'text_medium': '#6B7280',
            'text_light': '#9CA3AF',
            'border': '#E5E7EB'
        }

        # Configurar background
        self.root.configure(bg=self.colors['bg_light'])

        # Estilos de labels
        style.configure('Title.TLabel',
            font=('Inter', 24, 'bold'),
            foreground=self.colors['primary'],
            background=self.colors['bg_light']
        )

        style.configure('Subtitle.TLabel',
            font=('Inter', 13, 'bold'),
            foreground=self.colors['text_dark'],
            background=self.colors['bg_light']
        )

        style.configure('Normal.TLabel',
            font=('Inter', 10),
            foreground=self.colors['text_medium'],
            background=self.colors['bg_light']
        )

        style.configure('Emoji.TLabel',
            font=('Segoe UI Emoji', 32),
            background=self.colors['bg_card']
        )

        # Estilos de botões
        style.configure('Primary.TButton',
            font=('Inter', 12, 'bold'),
            padding=(20, 15),
            borderwidth=0
        )

        style.map('Primary.TButton',
            background=[('active', self.colors['primary_dark']), ('!active', self.colors['primary'])],
            foreground=[('active', 'white'), ('!active', 'white')]
        )

        style.configure('Success.TButton',
            font=('Inter', 11, 'bold'),
            padding=(15, 12)
        )

        style.configure('Action.TButton',
            font=('Inter', 10, 'bold'),
            padding=(10, 8)
        )

        # Estilos de frames
        style.configure('Card.TFrame',
            background=self.colors['bg_card'],
            relief='flat'
        )

        style.configure('TFrame',
            background=self.colors['bg_light']
        )

        # Estilos de LabelFrame
        style.configure('Card.TLabelframe',
            background=self.colors['bg_card'],
            borderwidth=0
        )

        style.configure('Card.TLabelframe.Label',
            font=('Inter', 12, 'bold'),
            foreground=self.colors['primary'],
            background=self.colors['bg_card']
        )

    def _criar_interface(self):
        """Cria todos os elementos da interface gráfica."""

        # Container principal com scrollbar
        container = tk.Frame(self.root, bg=self.colors['bg_light'])
        container.pack(fill=tk.BOTH, expand=True)

        # Canvas e Scrollbar
        canvas = tk.Canvas(container, bg=self.colors['bg_light'], highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)

        # Frame scrollável dentro do canvas
        main_frame = tk.Frame(canvas, bg=self.colors['bg_light'])

        # Posicionar scrollbar e canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Criar window no canvas
        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")

        # Configurar scroll
        canvas.configure(yscrollcommand=scrollbar.set)

        # Atualizar região de scroll quando o conteúdo mudar
        def configure_scroll(event=None):
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Ajustar largura do main_frame para preencher o canvas
            canvas_width = canvas.winfo_width()
            if canvas_width > 1:  # Só atualizar se o canvas tiver largura válida
                canvas.itemconfig(canvas_window, width=canvas_width)

        main_frame.bind("<Configure>", configure_scroll)
        canvas.bind("<Configure>", configure_scroll)

        # Scroll com mouse wheel
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Forçar atualização inicial após 100ms
        self.root.after(100, configure_scroll)

        # ===== CABEÇALHO COM DESIGN MODERNO =====
        header_frame = tk.Frame(main_frame, bg=self.colors['bg_light'])
        header_frame.pack(fill=tk.X, pady=(0, 25), padx=25)

        # Título com emoji grande
        title_container = tk.Frame(header_frame, bg=self.colors['bg_light'])
        title_container.pack()

        tk.Label(
            title_container,
            text="🏛️",
            font=('Segoe UI Emoji', 48),
            bg=self.colors['bg_light']
        ).pack(side=tk.LEFT, padx=(0, 15))

        title_text_frame = tk.Frame(title_container, bg=self.colors['bg_light'])
        title_text_frame.pack(side=tk.LEFT)

        ttk.Label(
            title_text_frame,
            text="VERIFICADOR INCRA PRO",
            style='Title.TLabel'
        ).pack(anchor=tk.W)

        ttk.Label(
            title_text_frame,
            text="Sistema Inteligente de Análise e Conferência Georreferenciada",
            style='Normal.TLabel'
        ).pack(anchor=tk.W)

        # ===== BARRA DE FERRAMENTAS COM CARDS =====
        toolbar_card = self._criar_card(main_frame)
        toolbar_card.pack(fill=tk.X, pady=(0, 20), padx=25)

        toolbar_content = tk.Frame(toolbar_card, bg=self.colors['bg_card'])
        toolbar_content.pack(fill=tk.X, padx=20, pady=15)

        # Botão API Key estilizado
        api_frame = tk.Frame(toolbar_content, bg=self.colors['bg_card'])
        api_frame.pack(side=tk.LEFT, padx=(0, 20))

        tk.Button(
            api_frame,
            text="⚙️  Configurar API",
            command=self._abrir_config_api,
            font=('Inter', 10, 'bold'),
            bg=self.colors['info'],
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor='hand2',
            activebackground=self.colors['primary'],
            highlightthickness=2,
            highlightbackground=self.colors['info'],
            highlightcolor=self.colors['primary_dark']
        ).pack()

        # Status API
        self.api_status_label = tk.Label(
            api_frame,
            text="⭕ Não configurada",
            font=('Inter', 8),
            fg=self.colors['danger'],
            bg=self.colors['bg_card']
        )
        self.api_status_label.pack(pady=(5, 0))

        # Separador vertical
        tk.Frame(
            toolbar_content,
            width=2,
            bg=self.colors['border']
        ).pack(side=tk.LEFT, fill=tk.Y, padx=20)

        # Campo Prenotação estilizado
        prenotacao_frame = tk.Frame(toolbar_content, bg=self.colors['bg_card'])
        prenotacao_frame.pack(side=tk.LEFT)

        tk.Label(
            prenotacao_frame,
            text="📋",
            font=('Segoe UI Emoji', 20),
            bg=self.colors['bg_card']
        ).pack(side=tk.LEFT, padx=(0, 10))

        prenotacao_input_frame = tk.Frame(prenotacao_frame, bg=self.colors['bg_card'])
        prenotacao_input_frame.pack(side=tk.LEFT)

        tk.Label(
            prenotacao_input_frame,
            text="Nº Prenotação",
            font=('Inter', 11, 'bold'),
            fg=self.colors['text_dark'],
            bg=self.colors['bg_card']
        ).pack(anchor=tk.W)

        prenotacao_entry = tk.Entry(
            prenotacao_input_frame,
            textvariable=self.numero_prenotacao,
            font=('Inter', 13, 'bold'),
            width=15,
            relief=tk.SOLID,
            bg='#F3F4F6',
            fg=self.colors['primary'],
            insertbackground=self.colors['primary'],
            borderwidth=2,
            highlightthickness=0
        )
        prenotacao_entry.pack(pady=(5, 0), ipady=6, ipadx=8)

        vcmd = (self.root.register(self._validar_numero), '%P')
        prenotacao_entry.config(validate='key', validatecommand=vcmd)

        # ===== SELETOR DE MODO (CARDS GRANDES E BONITOS) =====
        modo_card = self._criar_card(main_frame)
        modo_card.pack(fill=tk.X, pady=(0, 20), padx=25)

        modo_content = tk.Frame(modo_card, bg=self.colors['bg_card'])
        modo_content.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            modo_content,
            text="Escolha o modo de operação:",
            font=('Inter', 13, 'bold'),
            fg=self.colors['text_dark'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 15))

        # Container para os cards de modo
        modos_container = tk.Frame(modo_content, bg=self.colors['bg_card'])
        modos_container.pack(fill=tk.X)

        # CARD MODO AUTOMÁTICO
        self.card_automatico = self._criar_modo_card(
            modos_container,
            "🤖",
            "MODO AUTOMÁTICO",
            "Busca inteligente na rede\nExtração automática com IA\nMais rápido e eficiente",
            self.colors['primary'],
            lambda: self._selecionar_modo("automatico")
        )
        self.card_automatico.pack(side=tk.LEFT, padx=(0, 15), expand=True, fill=tk.BOTH)

        # CARD MODO MANUAL
        self.card_manual = self._criar_modo_card(
            modos_container,
            "📝",
            "MODO MANUAL",
            "Selecione os arquivos manualmente\nMaior controle sobre os documentos\nRecomendado para casos especiais",
            self.colors['secondary'],
            lambda: self._selecionar_modo("manual")
        )
        self.card_manual.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        # ===== CONTEÚDO DO MODO SELECIONADO =====
        self.content_frame = tk.Frame(main_frame, bg=self.colors['bg_light'])
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20), padx=25)

        # Criar ambos os modos (esconder um deles)
        self._criar_modo_automatico_content()
        self._criar_modo_manual_content()

        # Selecionar modo inicial
        self._selecionar_modo("automatico")

        # ===== ÁREA DE RESULTADOS =====
        result_card = self._criar_card(main_frame)
        result_card.pack(fill=tk.BOTH, expand=True, padx=25)

        result_content = tk.Frame(result_card, bg=self.colors['bg_card'])
        result_content.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(
            result_content,
            text="📊  Relatório de Comparação",
            font=('Inter', 12, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg_card']
        ).pack(anchor=tk.W, pady=(0, 10))

        # ScrolledText com estilo
        self.resultado_text = scrolledtext.ScrolledText(
            result_content,
            font=('Consolas', 10),
            wrap=tk.WORD,
            relief=tk.SOLID,
            bg='#F9FAFB',
            fg=self.colors['text_dark'],
            insertbackground=self.colors['primary'],
            selectbackground=self.colors['primary'],
            selectforeground='white',
            borderwidth=2,
            highlightthickness=0
        )
        self.resultado_text.pack(fill=tk.BOTH, expand=True, ipady=10, ipadx=10)

        # ===== BARRA DE STATUS =====
        status_frame = tk.Frame(main_frame, bg=self.colors['bg_card'], height=40)
        status_frame.pack(fill=tk.X, pady=(15, 25), padx=25)

        # Label de status à esquerda
        self.status_label = tk.Label(
            status_frame,
            text="✨ Pronto para iniciar",
            font=('Inter', 10),
            fg=self.colors['success'],
            bg=self.colors['bg_card'],
            anchor=tk.W
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15, pady=10)

        # Botão "Limpar Tudo" à direita
        self.btn_limpar_tudo = tk.Button(
            status_frame,
            text="🗑️  Limpar Tudo",
            command=self._limpar_tudo,
            font=('Inter', 10, 'bold'),
            bg=self.colors['warning'],
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2',
            activebackground='#D97706',
            activeforeground='white'
        )
        self.btn_limpar_tudo.pack(side=tk.RIGHT, padx=15, pady=5)

    def _criar_card(self, parent):
        """Cria um card (frame com sombra e bordas arredondadas simuladas)."""
        card = tk.Frame(
            parent,
            bg=self.colors['bg_card'],
            highlightbackground=self.colors['border'],
            highlightthickness=1
        )
        return card

    def _criar_modo_card(self, parent, emoji, titulo, descricao, cor, comando):
        """Cria um card clicável para seleção de modo."""
        card = tk.Frame(
            parent,
            bg=self.colors['bg_card'],
            highlightbackground=self.colors['border'],
            highlightthickness=2,
            cursor='hand2'
        )

        # Conteúdo interno
        content = tk.Frame(card, bg=self.colors['bg_card'])
        content.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

        # Emoji grande
        tk.Label(
            content,
            text=emoji,
            font=('Segoe UI Emoji', 48),
            bg=self.colors['bg_card']
        ).pack(pady=(0, 15))

        # Título
        tk.Label(
            content,
            text=titulo,
            font=('Inter', 14, 'bold'),
            fg=cor,
            bg=self.colors['bg_card']
        ).pack()

        # Descrição
        tk.Label(
            content,
            text=descricao,
            font=('Inter', 9),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_card'],
            justify=tk.CENTER
        ).pack(pady=(10, 0))

        # Badge de status (inicialmente oculto)
        badge = tk.Label(
            content,
            text="✓ SELECIONADO",
            font=('Inter', 8, 'bold'),
            fg='white',
            bg=cor,
            padx=10,
            pady=4
        )

        # Evento de clique
        def on_click(event=None):
            comando()

        card.bind('<Button-1>', on_click)
        for widget in card.winfo_children():
            widget.bind('<Button-1>', on_click)
            for child in widget.winfo_children():
                child.bind('<Button-1>', on_click)

        # Guardar referências para atualização
        card.badge = badge
        card.cor = cor
        card.content_frame = content

        return card

    def _selecionar_modo(self, modo):
        """Alterna entre modos e atualiza visual dos cards."""
        self.modo_atual.set(modo)

        # Atualizar visual dos cards
        if modo == "automatico":
            # Destacar automático
            self.card_automatico.config(highlightbackground=self.colors['primary'], highlightthickness=3)
            self.card_automatico.badge.pack(pady=(15, 0))

            # Desmarcar manual
            self.card_manual.config(highlightbackground=self.colors['border'], highlightthickness=2)
            self.card_manual.badge.pack_forget()

            # Mostrar conteúdo
            self.manual_content.pack_forget()
            self.automatico_content.pack(fill=tk.BOTH, expand=True)

        else:  # manual
            # Destacar manual
            self.card_manual.config(highlightbackground=self.colors['secondary'], highlightthickness=3)
            self.card_manual.badge.pack(pady=(15, 0))

            # Desmarcar automático
            self.card_automatico.config(highlightbackground=self.colors['border'], highlightthickness=2)
            self.card_automatico.badge.pack_forget()

            # Mostrar conteúdo
            self.automatico_content.pack_forget()
            self.manual_content.pack(fill=tk.BOTH, expand=True)

    def _criar_modo_automatico_content(self):
        """Cria conteúdo do modo automático."""
        self.automatico_content = self._criar_card(self.content_frame)

        content = tk.Frame(self.automatico_content, bg=self.colors['bg_card'])
        content.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Descrição
        tk.Label(
            content,
            text="🚀  O sistema buscará automaticamente o arquivo na rede e processará tudo para você!",
            font=('Inter', 11),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 25))

        # ===== SUB-MODO POR PÁGINAS =====
        submodo_frame = tk.Frame(content, bg=self.colors['bg_card'])
        submodo_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        tk.Label(
            submodo_frame,
            text="📝  Sub-modo Por Páginas (Opcional)",
            font=('Inter', 11, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 10))

        tk.Label(
            submodo_frame,
            text="Deixe em branco para usar detecção automática por IA, ou especifique páginas/anexe manualmente",
            font=('Inter', 9),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 20))

        # Container para INCRA e PROJETO lado a lado
        docs_container = tk.Frame(submodo_frame, bg=self.colors['bg_card'])
        docs_container.pack(fill=tk.BOTH, expand=True)

        # ===== INCRA =====
        incra_submodo_frame = tk.Frame(docs_container, bg=self.colors['bg_light'],
                                        highlightbackground=self.colors['border'],
                                        highlightthickness=1)
        incra_submodo_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        incra_inner = tk.Frame(incra_submodo_frame, bg=self.colors['bg_light'])
        incra_inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        tk.Label(
            incra_inner,
            text="📄  Memorial INCRA",
            font=('Inter', 10, 'bold'),
            fg=self.colors['text_dark'],
            bg=self.colors['bg_light']
        ).pack(pady=(0, 10))

        # Input de páginas do INCRA
        tk.Label(
            incra_inner,
            text="Números das páginas (ex: 1,2,3):",
            font=('Inter', 9),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        ).pack(anchor=tk.W, pady=(5, 2))

        incra_paginas_entry = tk.Entry(
            incra_inner,
            textvariable=self.incra_paginas,
            font=('Inter', 10),
            width=20,
            relief=tk.SOLID,
            bg='white',
            fg=self.colors['text_dark'],
            insertbackground=self.colors['primary'],
            borderwidth=1
        )
        incra_paginas_entry.pack(fill=tk.X, pady=(0, 10), ipady=4)

        # Separador OU
        tk.Label(
            incra_inner,
            text="━━━━━━ OU ━━━━━━",
            font=('Inter', 8),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        ).pack(pady=10)

        # Botão de anexo manual do INCRA
        tk.Label(
            incra_inner,
            text="Anexar arquivo manualmente:",
            font=('Inter', 9),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        ).pack(anchor=tk.W, pady=(5, 2))

        incra_anexo_btn = tk.Button(
            incra_inner,
            text="📁 Selecionar PDF",
            command=lambda: self._selecionar_arquivo_anexo(self.incra_anexo_path, "INCRA"),
            font=('Inter', 9),
            bg=self.colors['secondary'],
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor='hand2'
        )
        incra_anexo_btn.pack(fill=tk.X, pady=(0, 5))

        # Label para mostrar arquivo selecionado
        self.incra_anexo_label = tk.Label(
            incra_inner,
            text="Nenhum arquivo selecionado",
            font=('Inter', 8),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light'],
            wraplength=200
        )
        self.incra_anexo_label.pack()

        # ===== PROJETO =====
        projeto_submodo_frame = tk.Frame(docs_container, bg=self.colors['bg_light'],
                                          highlightbackground=self.colors['border'],
                                          highlightthickness=1)
        projeto_submodo_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        projeto_inner = tk.Frame(projeto_submodo_frame, bg=self.colors['bg_light'])
        projeto_inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        tk.Label(
            projeto_inner,
            text="📐  Planta/Projeto",
            font=('Inter', 10, 'bold'),
            fg=self.colors['text_dark'],
            bg=self.colors['bg_light']
        ).pack(pady=(0, 10))

        # Input de páginas do PROJETO
        tk.Label(
            projeto_inner,
            text="Números das páginas (ex: 4,5,6):",
            font=('Inter', 9),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        ).pack(anchor=tk.W, pady=(5, 2))

        projeto_paginas_entry = tk.Entry(
            projeto_inner,
            textvariable=self.projeto_paginas,
            font=('Inter', 10),
            width=20,
            relief=tk.SOLID,
            bg='white',
            fg=self.colors['text_dark'],
            insertbackground=self.colors['primary'],
            borderwidth=1
        )
        projeto_paginas_entry.pack(fill=tk.X, pady=(0, 10), ipady=4)

        # Separador OU
        tk.Label(
            projeto_inner,
            text="━━━━━━ OU ━━━━━━",
            font=('Inter', 8),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        ).pack(pady=10)

        # Botão de anexo manual do PROJETO
        tk.Label(
            projeto_inner,
            text="Anexar arquivo manualmente:",
            font=('Inter', 9),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        ).pack(anchor=tk.W, pady=(5, 2))

        projeto_anexo_btn = tk.Button(
            projeto_inner,
            text="📁 Selecionar PDF",
            command=lambda: self._selecionar_arquivo_anexo(self.projeto_anexo_path, "PROJETO"),
            font=('Inter', 9),
            bg=self.colors['secondary'],
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor='hand2'
        )
        projeto_anexo_btn.pack(fill=tk.X, pady=(0, 5))

        # Label para mostrar arquivo selecionado
        self.projeto_anexo_label = tk.Label(
            projeto_inner,
            text="Nenhum arquivo selecionado",
            font=('Inter', 8),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light'],
            wraplength=200
        )
        self.projeto_anexo_label.pack()

        # Botão grande de iniciar
        self.btn_iniciar_automatico = tk.Button(
            content,
            text="🚀  INICIAR BUSCA AUTOMÁTICA",
            command=self._iniciar_modo_automatico,
            font=('Inter', 14, 'bold'),
            bg=self.colors['primary'],
            fg='white',
            relief=tk.FLAT,
            padx=40,
            pady=20,
            cursor='hand2',
            activebackground=self.colors['primary_dark'],
            activeforeground='white'
        )
        self.btn_iniciar_automatico.pack(pady=20)

        # Frame de preview (inicialmente oculto)
        self.preview_frame = tk.Frame(content, bg=self.colors['bg_card'])

        preview_title = tk.Label(
            self.preview_frame,
            text="👁️  Prévia dos Documentos Extraídos",
            font=('Inter', 12, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg_card']
        )
        preview_title.pack(pady=(20, 15))

        # Container para previews lado a lado
        preview_container = tk.Frame(self.preview_frame, bg=self.colors['bg_card'])
        preview_container.pack(fill=tk.BOTH, expand=True)

        # Preview INCRA
        incra_frame = tk.Frame(preview_container, bg=self.colors['bg_card'])
        incra_frame.pack(side=tk.LEFT, padx=15, expand=True)

        tk.Label(
            incra_frame,
            text="📄 Memorial INCRA",
            font=('Inter', 11, 'bold'),
            fg=self.colors['text_dark'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 10))

        self.incra_preview_label = tk.Label(
            incra_frame,
            bg=self.colors['bg_light'],
            relief=tk.FLAT,
            highlightthickness=2,
            highlightbackground=self.colors['border']
        )
        self.incra_preview_label.pack()

        # Preview Projeto
        projeto_frame = tk.Frame(preview_container, bg=self.colors['bg_card'])
        projeto_frame.pack(side=tk.LEFT, padx=15, expand=True)

        tk.Label(
            projeto_frame,
            text="📐 Planta/Projeto",
            font=('Inter', 11, 'bold'),
            fg=self.colors['text_dark'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 10))

        self.projeto_preview_label = tk.Label(
            projeto_frame,
            bg=self.colors['bg_light'],
            relief=tk.FLAT,
            highlightthickness=2,
            highlightbackground=self.colors['border']
        )
        self.projeto_preview_label.pack()

        # Botões de confirmação
        confirm_frame = tk.Frame(self.preview_frame, bg=self.colors['bg_card'])
        confirm_frame.pack(pady=25)

        tk.Button(
            confirm_frame,
            text="✅  CONTINUAR",
            command=self._confirmar_documentos_automaticos,
            font=('Inter', 12, 'bold'),
            bg=self.colors['success'],
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=12,
            cursor='hand2',
            activebackground='#059669'
        ).pack(side=tk.LEFT, padx=10)

        tk.Button(
            confirm_frame,
            text="✋  FAZER MANUAL",
            command=self._alternar_para_manual,
            font=('Inter', 12, 'bold'),
            bg=self.colors['warning'],
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=12,
            cursor='hand2',
            activebackground='#D97706'
        ).pack(side=tk.LEFT, padx=10)

    def _criar_modo_manual_content(self):
        """Cria conteúdo do modo manual."""
        self.manual_content = self._criar_card(self.content_frame)

        content = tk.Frame(self.manual_content, bg=self.colors['bg_card'])
        content.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Descrição
        tk.Label(
            content,
            text="📁  Selecione manualmente os arquivos PDF para comparação",
            font=('Inter', 11),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 25))

        # Seleção INCRA
        incra_card = tk.Frame(
            content,
            bg='#FEF3C7',
            highlightthickness=2,
            highlightbackground='#FCD34D'
        )
        incra_card.pack(fill=tk.X, pady=10)

        incra_content = tk.Frame(incra_card, bg='#FEF3C7')
        incra_content.pack(fill=tk.X, padx=20, pady=15)

        tk.Label(
            incra_content,
            text="📄  Memorial INCRA",
            font=('Inter', 12, 'bold'),
            fg='#92400E',
            bg='#FEF3C7'
        ).pack(anchor=tk.W, pady=(0, 10))

        incra_input_frame = tk.Frame(incra_content, bg='#FEF3C7')
        incra_input_frame.pack(fill=tk.X)

        tk.Entry(
            incra_input_frame,
            textvariable=self.incra_path,
            font=('Inter', 10),
            state='readonly',
            relief=tk.SOLID,
            bg='white',
            fg=self.colors['text_dark'],
            borderwidth=2,
            highlightthickness=0
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, ipadx=10)

        tk.Button(
            incra_input_frame,
            text="📁 Selecionar",
            command=lambda: self._selecionar_arquivo(self.incra_path, "INCRA"),
            font=('Inter', 10, 'bold'),
            bg='#F59E0B',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=(10, 0))

        # Seleção Projeto
        projeto_card = tk.Frame(
            content,
            bg='#DBEAFE',
            highlightthickness=2,
            highlightbackground='#60A5FA'
        )
        projeto_card.pack(fill=tk.X, pady=10)

        projeto_content = tk.Frame(projeto_card, bg='#DBEAFE')
        projeto_content.pack(fill=tk.X, padx=20, pady=15)

        tk.Label(
            projeto_content,
            text="📐  Planta/Projeto",
            font=('Inter', 12, 'bold'),
            fg='#1E40AF',
            bg='#DBEAFE'
        ).pack(anchor=tk.W, pady=(0, 10))

        projeto_input_frame = tk.Frame(projeto_content, bg='#DBEAFE')
        projeto_input_frame.pack(fill=tk.X)

        tk.Entry(
            projeto_input_frame,
            textvariable=self.projeto_path,
            font=('Inter', 10),
            state='readonly',
            relief=tk.SOLID,
            bg='white',
            fg=self.colors['text_dark'],
            borderwidth=2,
            highlightthickness=0
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, ipadx=10)

        tk.Button(
            projeto_input_frame,
            text="📁 Selecionar",
            command=lambda: self._selecionar_arquivo(self.projeto_path, "Projeto"),
            font=('Inter', 10, 'bold'),
            bg='#3B82F6',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=(10, 0))

        # Botão de comparação
        tk.Button(
            content,
            text="🔍  COMPARAR DOCUMENTOS",
            command=self._comparar_manual,
            font=('Inter', 14, 'bold'),
            bg=self.colors['secondary'],
            fg='white',
            relief=tk.FLAT,
            padx=40,
            pady=20,
            cursor='hand2',
            activebackground='#DB2777'
        ).pack(pady=30)

    def _validar_numero(self, valor):
        """Valida entrada para aceitar apenas números."""
        return valor == "" or valor.isdigit()

    def _carregar_api_key(self):
        """Carrega API key salva e atualiza interface."""
        api_key = self.config_manager.get_api_key()
        if api_key:
            self.api_status_label.config(
                text="✅ Configurada",
                fg=self.colors['success']
            )
        else:
            self.api_status_label.config(
                text="⭕ Não configurada",
                fg=self.colors['danger']
            )

    def _abrir_config_api(self):
        """Abre janela para configurar API key."""
        config_window = tk.Toplevel(self.root)
        config_window.title("⚙️ Configuração da API Key")
        config_window.geometry("650x300")
        config_window.configure(bg=self.colors['bg_card'])
        config_window.transient(self.root)
        config_window.grab_set()

        main_frame = tk.Frame(config_window, bg=self.colors['bg_card'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Título
        tk.Label(
            main_frame,
            text="🔑  Configuração da API Key do Gemini",
            font=('Inter', 16, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg_card']
        ).pack(pady=(0, 10))

        tk.Label(
            main_frame,
            text="Insira sua API key abaixo. Ela será salva de forma segura e não precisará ser inserida novamente.",
            font=('Inter', 10),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_card'],
            wraplength=550
        ).pack(pady=10)

        # Campo de entrada
        api_var = tk.StringVar(value=self.config_manager.get_api_key())

        entry_frame = tk.Frame(main_frame, bg=self.colors['bg_card'])
        entry_frame.pack(fill=tk.X, pady=20)

        tk.Label(
            entry_frame,
            text="API Key:",
            font=('Inter', 11, 'bold'),
            fg=self.colors['text_dark'],
            bg=self.colors['bg_card']
        ).pack(anchor=tk.W, pady=(0, 8))

        api_entry = tk.Entry(
            entry_frame,
            textvariable=api_var,
            font=('Inter', 11),
            show="●",
            relief=tk.SOLID,
            bg='#F3F4F6',
            fg=self.colors['text_dark'],
            insertbackground=self.colors['primary'],
            borderwidth=2,
            highlightthickness=0
        )
        api_entry.pack(fill=tk.X, ipady=10, ipadx=10)

        # Botões
        btn_frame = tk.Frame(main_frame, bg=self.colors['bg_card'])
        btn_frame.pack(pady=20)

        def salvar_api():
            key = api_var.get().strip()
            if key:
                self.config_manager.set_api_key(key)
                self._carregar_api_key()
                messagebox.showinfo("✅ Sucesso", "API Key salva com sucesso!")
                config_window.destroy()
            else:
                messagebox.showwarning("⚠️ Aviso", "Por favor, insira uma API Key válida.")

        tk.Button(
            btn_frame,
            text="💾  Salvar",
            command=salvar_api,
            font=('Inter', 11, 'bold'),
            bg=self.colors['success'],
            fg='white',
            relief=tk.FLAT,
            padx=25,
            pady=10,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            btn_frame,
            text="❌  Cancelar",
            command=config_window.destroy,
            font=('Inter', 11, 'bold'),
            bg=self.colors['text_medium'],
            fg='white',
            relief=tk.FLAT,
            padx=25,
            pady=10,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)

    def _selecionar_arquivo(self, variavel, tipo):
        """Abre diálogo para selecionar arquivo PDF."""
        filename = filedialog.askopenfilename(
            title=f"Selecionar arquivo {tipo}",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if filename:
            variavel.set(filename)

    def _selecionar_arquivo_anexo(self, variavel, tipo):
        """Abre diálogo para selecionar arquivo PDF para anexo manual e atualiza label."""
        filename = filedialog.askopenfilename(
            title=f"Selecionar arquivo {tipo}",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if filename:
            variavel.set(filename)
            # Atualizar label correspondente
            nome_arquivo = Path(filename).name
            if tipo == "INCRA":
                self.incra_anexo_label.config(
                    text=f"✓ {nome_arquivo}",
                    fg=self.colors['success']
                )
            elif tipo == "PROJETO":
                self.projeto_anexo_label.config(
                    text=f"✓ {nome_arquivo}",
                    fg=self.colors['success']
                )

    def _atualizar_status(self, mensagem: str):
        """Atualiza a barra de status e a janela de progresso."""
        # Detectar tipo de mensagem e ajustar cor
        if "✅" in mensagem or "sucesso" in mensagem.lower():
            cor = self.colors['success']
        elif "❌" in mensagem or "erro" in mensagem.lower():
            cor = self.colors['danger']
        elif "🔄" in mensagem or "processando" in mensagem.lower():
            cor = self.colors['info']
        else:
            cor = self.colors['text_dark']

        self.status_label.config(text=mensagem, fg=cor)
        self.root.update_idletasks()

        # Atualizar também a janela de progresso se estiver aberta
        if self.progress_window and self.progress_text:
            try:
                self.progress_text.insert(tk.END, f"{mensagem}\n")
                self.progress_text.see(tk.END)
                self.progress_window.update_idletasks()
            except:
                pass

    def _mostrar_janela_progresso(self):
        """Cria e exibe a janela de progresso em tempo real."""
        if self.progress_window:
            try:
                self.progress_window.destroy()
            except:
                pass

        # Criar janela toplevel
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("🔄 Progresso em Tempo Real")
        self.progress_window.geometry("700x500")

        # Centralizar janela
        self.progress_window.transient(self.root)

        # Frame principal
        main_frame = tk.Frame(self.progress_window, bg=self.colors['bg_light'], padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = tk.Label(
            main_frame,
            text="🔄  Processamento em Andamento",
            font=('Inter', 16, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg_light']
        )
        title_label.pack(pady=(0, 15))

        # Subtítulo
        subtitle_label = tk.Label(
            main_frame,
            text="Acompanhe cada etapa do processo abaixo:",
            font=('Inter', 10),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        )
        subtitle_label.pack(pady=(0, 20))

        # ScrolledText para mostrar progresso
        self.progress_text = scrolledtext.ScrolledText(
            main_frame,
            font=('Consolas', 10),
            bg='#1E293B',
            fg='#F8FAFC',
            insertbackground='white',
            relief=tk.FLAT,
            padx=15,
            pady=15,
            wrap=tk.WORD
        )
        self.progress_text.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Configurar tags para colorir mensagens
        self.progress_text.tag_config('success', foreground='#10B981')
        self.progress_text.tag_config('error', foreground='#EF4444')
        self.progress_text.tag_config('info', foreground='#3B82F6')
        self.progress_text.tag_config('warning', foreground='#F59E0B')

        # Barra de progresso animada
        progress_frame = tk.Frame(main_frame, bg=self.colors['bg_light'], height=8)
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        self.progress_bar = tk.Canvas(progress_frame, height=8, bg=self.colors['border'], highlightthickness=0)
        self.progress_bar.pack(fill=tk.X)

        # Label de status
        self.progress_status_label = tk.Label(
            main_frame,
            text="Iniciando...",
            font=('Inter', 9),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        )
        self.progress_status_label.pack()

        # Label informativo
        info_label = tk.Label(
            main_frame,
            text="⏱️  Esta janela fechará automaticamente ao concluir",
            font=('Inter', 8, 'italic'),
            fg=self.colors['text_medium'],
            bg=self.colors['bg_light']
        )
        info_label.pack(pady=(15, 0))

        # Mensagem inicial
        self.progress_text.insert(tk.END, "═" * 60 + "\n")
        self.progress_text.insert(tk.END, "🚀  INICIANDO PROCESSAMENTO\n")
        self.progress_text.insert(tk.END, "═" * 60 + "\n\n")

        # Posicionar no centro da tela
        self.progress_window.update_idletasks()
        x = (self.progress_window.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.progress_window.winfo_screenheight() // 2) - (500 // 2)
        self.progress_window.geometry(f"700x500+{x}+{y}")

    def _fechar_janela_progresso(self):
        """Fecha a janela de progresso."""
        if self.progress_window:
            try:
                self.progress_window.destroy()
                self.progress_window = None
                self.progress_text = None
            except:
                pass

    def _fechar_progresso_automatico(self, delay=2000):
        """Fecha a janela de progresso automaticamente após um delay.

        Args:
            delay: Tempo em milissegundos antes de fechar (padrão: 2000ms = 2 segundos)
        """
        if self.progress_window:
            try:
                self.root.after(delay, self._fechar_janela_progresso)
            except:
                pass

    def _desabilitar_botoes(self):
        """Desabilita botões durante o processamento."""
        self.btn_iniciar_automatico.config(state='disabled', bg=self.colors['text_light'])

    def _habilitar_botoes(self):
        """Reabilita botões após o processamento."""
        self.btn_iniciar_automatico.config(state='normal', bg=self.colors['primary'])

    # ========== MODO MANUAL ==========

    def _comparar_manual(self):
        """Executa comparação no modo manual."""
        if not self._validar_entrada_manual():
            return

        # Mostrar janela de progresso
        self._mostrar_janela_progresso()

        def executar():
            try:
                self._desabilitar_botoes()
                self._atualizar_status("🔄 Processando documentos...")

                # Extrair dados para Excel
                self._atualizar_status("📄 Extraindo dados do INCRA...")
                self.incra_excel_path, self.incra_data = self._extrair_pdf_para_excel(
                    self.incra_path.get(), "incra"
                )

                self._atualizar_status("📐 Extraindo dados do Projeto...")
                self.projeto_excel_path, self.projeto_data = self._extrair_pdf_para_excel(
                    self.projeto_path.get(), "normal"
                )

                # Gerar relatório
                self._atualizar_status("📊 Gerando relatório de comparação...")
                relatorio = self._construir_relatorio_comparacao(True, False)

                # Salvar e abrir relatório
                self._salvar_e_abrir_relatorio(relatorio)

                # Mostrar resumo
                self._mostrar_resumo_no_texto()

                self._atualizar_status("✅ Comparação concluída com sucesso!")
                self._atualizar_status("\n" + "═" * 60)
                self._atualizar_status("🎉  PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
                self._atualizar_status("═" * 60 + "\n")

                # Fechar janela de progresso automaticamente após 2 segundos
                self._fechar_progresso_automatico()

            except Exception as e:
                self._atualizar_status(f"❌ Erro: {str(e)}")
                self._atualizar_status("\n" + "═" * 60)
                self._atualizar_status("❌  PROCESSAMENTO FINALIZADO COM ERRO")
                self._atualizar_status("═" * 60 + "\n")

                # Fechar janela de progresso automaticamente após 3 segundos
                self._fechar_progresso_automatico(delay=3000)

                messagebox.showerror("Erro", f"Erro ao processar documentos:\n\n{str(e)}")
            finally:
                self._habilitar_botoes()

        # Executar em thread separada
        threading.Thread(target=executar, daemon=True).start()

    def _validar_entrada_manual(self) -> bool:
        """Valida entradas do modo manual."""
        api_key = self.config_manager.get_api_key()
        if not api_key:
            messagebox.showerror("Erro", "Por favor, configure a API Key primeiro.")
            return False

        if not self.incra_path.get():
            messagebox.showerror("Erro", "Por favor, selecione o arquivo INCRA.")
            return False

        if not self.projeto_path.get():
            messagebox.showerror("Erro", "Por favor, selecione o arquivo Projeto/Planta.")
            return False

        if not self.numero_prenotacao.get():
            messagebox.showerror("Erro", "Por favor, insira o Número de Prenotação.")
            return False

        return True

    # ========== MODO AUTOMÁTICO ==========

    def _iniciar_modo_automatico(self):
        """Inicia o processo automático."""
        if not self._validar_entrada_automatico():
            return

        # Mostrar janela de progresso
        self._mostrar_janela_progresso()

        def executar():
            try:
                self._desabilitar_botoes()

                # 1. Verificar se precisa buscar arquivo TIFF
                # (só busca se pelo menos um dos documentos não tiver anexo manual)
                pdf_path = None
                precisa_buscar = (
                    (not self.incra_anexo_path.get() and not self.incra_paginas.get()) or
                    (not self.projeto_anexo_path.get() and not self.projeto_paginas.get()) or
                    self.incra_paginas.get() or
                    self.projeto_paginas.get()
                )

                if precisa_buscar:
                    self._atualizar_status("🔍 Buscando arquivo TIFF na rede...")
                    tiff_path = self._buscar_arquivo_tiff()

                    if not tiff_path:
                        raise Exception("Arquivo TIFF não encontrado na rede.")

                    # 2. Copiar e converter para PDF
                    self._atualizar_status("📋 Copiando e convertendo TIFF para PDF...")
                    pdf_path = self._converter_tiff_para_pdf(tiff_path)

                # 3. Processar Memorial INCRA (com lógica de ignorar inputs vazios)
                self._atualizar_status("📄 Processando Memorial INCRA...")

                if self.incra_paginas.get().strip():
                    # Usuário especificou páginas manualmente
                    self._atualizar_status(f"📄 Extraindo páginas especificadas do INCRA: {self.incra_paginas.get()}")
                    if not pdf_path:
                        raise Exception("Número de prenotação necessário para buscar páginas específicas do INCRA")
                    self.pdf_extraido_incra = self._extrair_paginas_especificas(
                        pdf_path, self.incra_paginas.get(), 'incra'
                    )
                elif self.incra_anexo_path.get().strip():
                    # Usuário anexou arquivo manualmente
                    self._atualizar_status("📄 Usando anexo manual do INCRA")
                    # Copiar anexo para pasta temporária
                    output_dir = Path.home() / "Downloads" / "conferencia_geo_temp"
                    output_dir.mkdir(parents=True, exist_ok=True)
                    dest = output_dir / "memorial_incra_extraido.pdf"
                    shutil.copy2(self.incra_anexo_path.get(), dest)
                    self.pdf_extraido_incra = str(dest)
                else:
                    # Usar detecção automática por IA (comportamento original)
                    self._atualizar_status("📄 Usando detecção automática por IA para INCRA")
                    if not pdf_path:
                        raise Exception("Número de prenotação necessário para detecção automática do INCRA")
                    self.pdf_extraido_incra = self._extrair_memorial_incra_do_pdf(pdf_path)

                # 4. Processar Planta/Projeto (com lógica de ignorar inputs vazios)
                self._atualizar_status("📐 Processando Planta/Projeto...")

                if self.projeto_paginas.get().strip():
                    # Usuário especificou páginas manualmente
                    self._atualizar_status(f"📐 Extraindo páginas especificadas do PROJETO: {self.projeto_paginas.get()}")
                    if not pdf_path:
                        raise Exception("Número de prenotação necessário para buscar páginas específicas do PROJETO")
                    self.pdf_extraido_projeto = self._extrair_paginas_especificas(
                        pdf_path, self.projeto_paginas.get(), 'projeto'
                    )
                elif self.projeto_anexo_path.get().strip():
                    # Usuário anexou arquivo manualmente
                    self._atualizar_status("📐 Usando anexo manual do PROJETO")
                    # Copiar anexo para pasta temporária
                    output_dir = Path.home() / "Downloads" / "conferencia_geo_temp"
                    output_dir.mkdir(parents=True, exist_ok=True)
                    dest = output_dir / "projeto_extraido.pdf"
                    shutil.copy2(self.projeto_anexo_path.get(), dest)
                    self.pdf_extraido_projeto = str(dest)
                else:
                    # Usar detecção automática por IA (comportamento original)
                    self._atualizar_status("📐 Usando detecção automática por IA para PROJETO")
                    if not pdf_path:
                        raise Exception("Número de prenotação necessário para detecção automática do PROJETO")
                    self.pdf_extraido_projeto = self._extrair_projeto_do_pdf(pdf_path)

                # 5. Salvar backups
                self._atualizar_status("💾 Salvando backups...")
                self._salvar_backups_pdfs()

                # 6. Gerar previews
                self._atualizar_status("👁️ Gerando prévias...")
                self._gerar_previews()

                # 7. Mostrar frame de preview
                self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=20)

                self._atualizar_status("✅ Documentos extraídos! Verifique as prévias.")
                self._atualizar_status("\n" + "═" * 60)
                self._atualizar_status("🎉  PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
                self._atualizar_status("═" * 60 + "\n")

                # Fechar janela de progresso automaticamente após 2 segundos
                self._fechar_progresso_automatico()

            except Exception as e:
                self._atualizar_status(f"❌ Erro: {str(e)}")
                self._atualizar_status("\n" + "═" * 60)
                self._atualizar_status("❌  PROCESSAMENTO FINALIZADO COM ERRO")
                self._atualizar_status("═" * 60 + "\n")

                # Fechar janela de progresso automaticamente após 3 segundos (para dar tempo de ler o erro)
                self._fechar_progresso_automatico(delay=3000)

                messagebox.showerror("Erro", f"Erro no modo automático:\n\n{str(e)}")
                self._habilitar_botoes()

        # Executar em thread separada
        threading.Thread(target=executar, daemon=True).start()

    def _validar_entrada_automatico(self) -> bool:
        """Valida entradas do modo automático."""
        api_key = self.config_manager.get_api_key()
        if not api_key:
            messagebox.showerror("Erro", "Por favor, configure a API Key primeiro.")
            return False

        if not self.numero_prenotacao.get():
            messagebox.showerror("Erro", "Por favor, insira o Número de Prenotação.")
            return False

        return True

    def _buscar_arquivo_tiff(self) -> Optional[str]:
        """Busca arquivo TIFF na rede baseado no número de prenotação."""
        numero = int(self.numero_prenotacao.get())
        numero_formatado = f"{numero:08d}"

        # Calcular subpasta
        milhar = math.ceil(numero / 1000) * 1000
        subpasta_formatada = f"{milhar:08d}"

        # Montar caminho
        base_path = Path(r"\\192.168.20.100\trabalho\TRABALHO\IMAGENS\IMOVEIS\DOCUMENTOS - DIVERSOS")
        tiff_path = base_path / subpasta_formatada / f"{numero_formatado}.tif"

        self._atualizar_status(f"🔍 Buscando: {tiff_path}")

        if tiff_path.exists():
            return str(tiff_path)

        return None

    def _converter_tiff_para_pdf(self, tiff_path: str) -> str:
        """Copia TIFF para Downloads e converte para PDF."""
        downloads_dir = Path.home() / "Downloads" / "conferencia_geo_temp"
        downloads_dir.mkdir(parents=True, exist_ok=True)

        # Copiar TIFF
        tiff_filename = Path(tiff_path).name
        tiff_dest = downloads_dir / tiff_filename
        shutil.copy2(tiff_path, tiff_dest)

        # Converter para PDF
        pdf_path = downloads_dir / f"{Path(tiff_filename).stem}.pdf"

        # Abrir TIFF multi-página
        img = Image.open(tiff_dest)
        images = []

        try:
            while True:
                images.append(img.copy().convert('RGB'))
                img.seek(img.tell() + 1)
        except EOFError:
            pass

        # Salvar como PDF
        if images:
            images[0].save(
                pdf_path,
                save_all=True,
                append_images=images[1:],
                resolution=200.0
            )

        return str(pdf_path)

    def _extrair_paginas_especificas(self, pdf_path: str, paginas_str: str, tipo: str) -> str:
        """Extrai páginas específicas de um PDF baseado em string (ex: '1,2,3').

        Args:
            pdf_path: Caminho do PDF completo
            paginas_str: String com números de páginas separados por vírgula (ex: '1,2,3')
            tipo: 'incra' ou 'projeto' para nomear o arquivo de saída

        Returns:
            Caminho do PDF com as páginas extraídas
        """
        output_dir = Path.home() / "Downloads" / "conferencia_geo_temp"
        output_dir.mkdir(parents=True, exist_ok=True)

        if tipo == 'incra':
            output_pdf = output_dir / "memorial_incra_extraido.pdf"
        else:
            output_pdf = output_dir / "projeto_extraido.pdf"

        # Parsear string de páginas
        try:
            # Remove espaços e split por vírgula
            paginas_list = [int(p.strip()) for p in paginas_str.split(',') if p.strip()]

            # Converter para índice 0-based (usuário digita 1-based)
            paginas_indices = [p - 1 for p in paginas_list if p > 0]

            if not paginas_indices:
                raise ValueError("Nenhuma página válida especificada")

            # Extrair páginas
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                writer = PyPDF2.PdfWriter()

                total_paginas = len(reader.pages)

                for page_num in paginas_indices:
                    if page_num < total_paginas:
                        writer.add_page(reader.pages[page_num])
                    else:
                        self._atualizar_status(f"⚠️ Página {page_num + 1} não existe no PDF (total: {total_paginas})")

                with open(output_pdf, 'wb') as output_file:
                    writer.write(output_file)

            return str(output_pdf)

        except ValueError as e:
            raise ValueError(f"Formato inválido para números de páginas: {str(e)}")

    def _extrair_memorial_incra_do_pdf(self, pdf_path: str) -> str:
        """Extrai páginas do Memorial INCRA do PDF."""
        output_dir = Path.home() / "Downloads" / "conferencia_geo_temp"
        output_pdf = output_dir / "memorial_incra_extraido.pdf"

        # Usar Gemini
        api_key = self.config_manager.get_api_key()
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')

        images = convert_from_path(pdf_path, dpi=150)
        paginas_encontradas = []

        for i, img in enumerate(images):
            temp_img_path = output_dir / f"temp_page_{i}.jpg"
            img.save(temp_img_path, 'JPEG')

            prompt = """
            Analise esta imagem e responda apenas com 'SIM' ou 'NAO':
            Esta página contém o Memorial Descritivo do INCRA?

            Características do Memorial INCRA:
            - Texto: "MINISTÉRIO DA AGRICULTURA, PECUÁRIA E ABASTECIMENTO"
            - Texto: "INSTITUTO NACIONAL DE COLONIZAÇÃO E REFORMA AGRÁRIA"
            - Texto: "MEMORIAL DESCRITIVO"
            - Tabela com colunas: "VÉRTICE", "SEGMENTO VANTE", "Confrontações"

            Responda apenas: SIM ou NAO
            """

            try:
                img_upload = Image.open(temp_img_path)
                response = model.generate_content([prompt, img_upload])
                resposta = response.text.strip().upper()

                if 'SIM' in resposta:
                    paginas_encontradas.append(i)

            except Exception as e:
                print(f"Erro ao analisar página {i}: {e}")

            temp_img_path.unlink()

        # Extrair páginas
        if paginas_encontradas:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                writer = PyPDF2.PdfWriter()

                for page_num in paginas_encontradas:
                    writer.add_page(reader.pages[page_num])

                with open(output_pdf, 'wb') as output_file:
                    writer.write(output_file)

        return str(output_pdf)

    def _extrair_projeto_do_pdf(self, pdf_path: str) -> str:
        """Extrai páginas da Planta/Projeto do PDF."""
        output_dir = Path.home() / "Downloads" / "conferencia_geo_temp"
        output_pdf = output_dir / "projeto_extraido.pdf"

        # Usar Gemini
        api_key = self.config_manager.get_api_key()
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')

        images = convert_from_path(pdf_path, dpi=150)
        paginas_encontradas = []

        for i, img in enumerate(images):
            temp_img_path = output_dir / f"temp_page_{i}.jpg"
            img.save(temp_img_path, 'JPEG')

            prompt = """
            Analise esta imagem e responda apenas com 'SIM' ou 'NAO':
            Esta página contém a Planta/Projeto de Georreferenciamento?

            Características da Planta/Projeto:
            - Títulos: "PLANTA DO IMÓVEL GEORREFERENCIADO" ou "PLANTA DE SITUAÇÃO"
            - Identificadores: "Código INCRA:", "Matrícula nº:", "Responsável técnico:"
            - Tabela com coordenadas (colunas: "Código", "Longitude", "Latitude")

            Responda apenas: SIM ou NAO
            """

            try:
                img_upload = Image.open(temp_img_path)
                response = model.generate_content([prompt, img_upload])
                resposta = response.text.strip().upper()

                if 'SIM' in resposta:
                    paginas_encontradas.append(i)

            except Exception as e:
                print(f"Erro ao analisar página {i}: {e}")

            temp_img_path.unlink()

        # Extrair páginas
        if paginas_encontradas:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                writer = PyPDF2.PdfWriter()

                for page_num in paginas_encontradas:
                    writer.add_page(reader.pages[page_num])

                with open(output_pdf, 'wb') as output_file:
                    writer.write(output_file)

        return str(output_pdf)

    def _salvar_backups_pdfs(self):
        """Salva backups dos PDFs extraídos."""
        docs_dir = Path.home() / "Documentos" / "Relatórios INCRA"

        incra_dir = docs_dir / "PDF_INCRAS"
        projeto_dir = docs_dir / "PDF_PLANTAS"

        incra_dir.mkdir(parents=True, exist_ok=True)
        projeto_dir.mkdir(parents=True, exist_ok=True)

        numero = self.numero_prenotacao.get()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if self.pdf_extraido_incra:
            dest_incra = incra_dir / f"INCRA_{numero}_{timestamp}.pdf"
            shutil.copy2(self.pdf_extraido_incra, dest_incra)

        if self.pdf_extraido_projeto:
            dest_projeto = projeto_dir / f"PROJETO_{numero}_{timestamp}.pdf"
            shutil.copy2(self.pdf_extraido_projeto, dest_projeto)

    def _gerar_previews(self):
        """Gera thumbnails dos documentos extraídos."""
        if self.pdf_extraido_incra and Path(self.pdf_extraido_incra).exists():
            images = convert_from_path(self.pdf_extraido_incra, dpi=100, first_page=1, last_page=1)
            if images:
                self.preview_incra_image = images[0]
                self.preview_incra_image.thumbnail((300, 400))

                photo = ImageTk.PhotoImage(self.preview_incra_image)
                self.incra_preview_label.config(image=photo)
                self.incra_preview_label.image = photo

        if self.pdf_extraido_projeto and Path(self.pdf_extraido_projeto).exists():
            images = convert_from_path(self.pdf_extraido_projeto, dpi=100, first_page=1, last_page=1)
            if images:
                self.preview_projeto_image = images[0]
                self.preview_projeto_image.thumbnail((300, 400))

                photo = ImageTk.PhotoImage(self.preview_projeto_image)
                self.projeto_preview_label.config(image=photo)
                self.projeto_preview_label.image = photo

    def _confirmar_documentos_automaticos(self):
        """Usuário confirmou documentos - prosseguir com comparação."""
        self.incra_path.set(self.pdf_extraido_incra)
        self.projeto_path.set(self.pdf_extraido_projeto)

        self.preview_frame.pack_forget()

        self._comparar_manual()

    def _alternar_para_manual(self):
        """Usuário optou por fazer manual."""
        self.preview_frame.pack_forget()
        self._selecionar_modo("manual")
        self._habilitar_botoes()
        messagebox.showinfo(
            "Modo Manual",
            "Selecione manualmente os arquivos corretos no Modo Manual."
        )

    # ========== EXTRAÇÃO E COMPARAÇÃO ==========

    def _extrair_pdf_para_excel(self, pdf_path: str, tipo: str = "normal") -> tuple[str, Dict]:
        """Extrai dados de um PDF memorial para Excel."""
        try:
            api_key = self.config_manager.get_api_key()
            genai.configure(api_key=api_key)

            output_dir = Path(tempfile.gettempdir()) / "conferencia_geo"
            output_dir.mkdir(parents=True, exist_ok=True)

            if not output_dir.exists():
                raise RuntimeError(f"Não foi possível criar o diretório: {output_dir}")

            nome_base = Path(pdf_path).stem
            excel_path = output_dir / f"{nome_base}_extraido.xlsx"

            if tipo == "incra":
                dados = extrair_memorial_incra(pdf_path, api_key)
            else:
                dados = extract_table_from_pdf(pdf_path, api_key)

            if not dados or 'data' not in dados:
                raise ValueError("Nenhum dado foi extraído do PDF")

            create_excel_file(dados, str(excel_path))

            if not excel_path.exists():
                raise RuntimeError(f"Arquivo Excel não foi criado")

            if excel_path.stat().st_size == 0:
                raise RuntimeError(f"Arquivo Excel está vazio")

            return str(excel_path), dados

        except Exception as e:
            error_msg = f"Erro ao extrair PDF para Excel: {str(e)}"
            raise RuntimeError(error_msg) from e

    def _normalizar_coordenada(self, coord: str) -> str:
        """Normaliza coordenadas para comparação."""
        if not coord:
            return ""

        coord = str(coord).strip()
        coord = coord.replace("′", "'").replace("″", '"')

        if coord.startswith("-"):
            coord = coord[1:].strip()

        coord = coord.replace(" W", "").replace(" S", "").strip()
        coord = coord.strip().strip('"').strip("'").strip()

        return coord

    def _limpar_string(self, valor) -> str:
        """Limpa strings e converte pontos para vírgulas."""
        if valor is None:
            return ""

        valor_limpo = str(valor).strip()

        while "  " in valor_limpo:
            valor_limpo = valor_limpo.replace("  ", " ")

        valor_limpo = valor_limpo.replace(".", ",")

        return valor_limpo

    def _construir_relatorio_comparacao(self, incluir_projeto: bool, incluir_memorial: bool) -> str:
        """Constrói relatório HTML comparando dados estruturados."""
        html = []

        html.append("""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Conferência INCRA</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            padding: 30px;
            line-height: 1.6;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: #ffffff;
            padding: 50px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            border-bottom: 3px solid #333;
            padding-bottom: 25px;
            margin-bottom: 35px;
        }
        h1 {
            color: #1a1a1a;
            margin-bottom: 10px;
            font-size: 28px;
            font-weight: 600;
            letter-spacing: -0.5px;
        }
        .subtitle {
            color: #666;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .info-section {
            background: #fafafa;
            border-left: 4px solid #333;
            padding: 20px 25px;
            margin-bottom: 35px;
        }
        .info-section p {
            margin: 8px 0;
            color: #333;
            font-size: 14px;
        }
        .info-section strong {
            color: #000;
            font-weight: 600;
        }
        .section-title {
            color: #1a1a1a;
            font-size: 18px;
            font-weight: 600;
            margin: 45px 0 20px 0;
            padding-bottom: 12px;
            border-bottom: 2px solid #ddd;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .linha-grupo {
            background: #fafafa;
            border: 2px solid #ddd;
            border-radius: 8px;
            padding: 25px;
            margin-bottom: 30px;
        }
        .linha-titulo {
            color: #1a1a1a;
            font-size: 16px;
            font-weight: 700;
            margin-bottom: 20px;
            padding: 12px;
            background: #333;
            color: #fff;
            text-align: center;
            border-radius: 4px;
            letter-spacing: 1px;
        }
        .subtitulo-secao {
            color: #555;
            font-size: 13px;
            font-weight: 600;
            margin: 20px 0 10px 0;
            padding-left: 10px;
            border-left: 4px solid #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .tabela-linha {
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 40px;
            font-size: 13px;
            border: 1px solid #ddd;
        }
        thead {
            background: #333;
            color: #fff;
        }
        th {
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        tbody tr {
            border-bottom: 1px solid #e8e8e8;
        }
        tbody tr:last-child {
            border-bottom: none;
        }
        td {
            padding: 12px;
            color: #333;
        }
        tbody tr:hover {
            background-color: #f9f9f9;
        }
        .identico {
            background-color: #e8f5e9 !important;
            border-left: 3px solid #4caf50;
        }
        .identico:hover {
            background-color: #d4edda !important;
        }
        .diferente {
            background-color: #ffebee !important;
            border-left: 3px solid #f44336;
        }
        .diferente:hover {
            background-color: #f8d7da !important;
        }
        .diferente td {
            font-weight: 600;
        }
        .resumo-section {
            background: #fafafa;
            border: 2px solid #333;
            padding: 30px;
            margin-top: 45px;
            border-radius: 4px;
        }
        .resumo-section h2 {
            color: #1a1a1a;
            margin-bottom: 25px;
            font-size: 20px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .resumo-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 25px;
        }
        .resumo-card {
            background: #fff;
            border-left: 4px solid #666;
            padding: 18px;
        }
        .resumo-card h4 {
            color: #333;
            margin-bottom: 12px;
            font-size: 14px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .resumo-card p {
            margin: 6px 0;
            color: #666;
            font-size: 14px;
        }
        .resumo-card strong {
            color: #000;
            font-weight: 600;
            font-size: 16px;
        }
        .resumo-total {
            background: #333;
            color: #fff;
            padding: 20px;
            margin-top: 25px;
            border-radius: 4px;
        }
        .resumo-total h4 {
            color: #fff;
            margin-bottom: 12px;
            font-size: 14px;
        }
        .resumo-total p {
            color: #fff;
            font-size: 14px;
        }
        .resumo-total strong {
            color: #fff;
            font-size: 16px;
        }
        .footer {
            text-align: center;
            margin-top: 50px;
            padding-top: 25px;
            border-top: 2px solid #ddd;
            color: #999;
            font-size: 12px;
        }
        @media print {
            body {
                background: #fff;
                padding: 0;
            }
            .container {
                box-shadow: none;
                padding: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>RELATÓRIO DE CONFERÊNCIA INCRA</h1>
            <p class="subtitle">Sistema de Análise e Verificação Georreferenciamento</p>
        </div>
""")

        html.append(f"""
        <div class="info-section">
            <p><strong>Data:</strong> {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}</p>
            <p><strong>Nº Prenotação:</strong> {self.numero_prenotacao.get()}</p>
        </div>
""")

        # Carregar dados
        wb_incra = load_workbook(self.incra_excel_path)
        ws_incra = wb_incra.active
        dados_incra = list(ws_incra.iter_rows(values_only=True))

        wb_projeto = load_workbook(self.projeto_excel_path)
        ws_projeto = wb_projeto.active
        dados_projeto = list(ws_projeto.iter_rows(values_only=True))

        identicos_vertice = 0
        diferencas_vertice = 0
        identicos_segmento = 0
        diferencas_segmento = 0

        # VÉRTICE E SEGMENTO VANTE (Agrupados por linha)
        html.append('<h2 class="section-title">Comparação Completa por Linha</h2>')

        max_rows = max(len(dados_incra), len(dados_projeto))

        for i in range(1, max_rows):
            incra_row = dados_incra[i] if i < len(dados_incra) else []
            projeto_row = dados_projeto[i] if i < len(dados_projeto) else []

            # Cabeçalho da linha
            html.append(f'<div class="linha-grupo">')
            html.append(f'<h3 class="linha-titulo">LINHA {i}</h3>')

            # Tabela de Vértices para esta linha
            html.append('<h4 class="subtitulo-secao">Vértice</h4>')
            html.append('<table class="tabela-linha">')
            html.append('<thead><tr>')
            html.append('<th>Campo</th><th>INCRA</th><th>Projeto</th><th>Status</th>')
            html.append('</tr></thead><tbody>')

            # Dados do Vértice
            codigo_incra = self._limpar_string(incra_row[0] if len(incra_row) > 0 else "")
            codigo_projeto = self._limpar_string(projeto_row[0] if len(projeto_row) > 0 else "")

            long_incra = self._normalizar_coordenada(self._limpar_string(incra_row[1] if len(incra_row) > 1 else ""))
            long_projeto = self._normalizar_coordenada(self._limpar_string(projeto_row[1] if len(projeto_row) > 1 else ""))

            lat_incra = self._normalizar_coordenada(self._limpar_string(incra_row[2] if len(incra_row) > 2 else ""))
            lat_projeto = self._normalizar_coordenada(self._limpar_string(projeto_row[2] if len(projeto_row) > 2 else ""))

            alt_incra = self._limpar_string(incra_row[3] if len(incra_row) > 3 else "")
            alt_projeto = self._limpar_string(projeto_row[3] if len(projeto_row) > 3 else "")

            campos_vertice = [
                ("Código", codigo_incra, codigo_projeto),
                ("Longitude", long_incra, long_projeto),
                ("Latitude", lat_incra, lat_projeto),
                ("Altitude", alt_incra, alt_projeto)
            ]

            for campo, val_incra, val_projeto in campos_vertice:
                status_classe = "identico" if val_incra == val_projeto else "diferente"
                status_texto = "✅ Idêntico" if val_incra == val_projeto else "❌ Diferente"

                if val_incra == val_projeto:
                    identicos_vertice += 1
                else:
                    diferencas_vertice += 1

                html.append(f'<tr class="{status_classe}">')
                html.append(f'<td><strong>{campo}</strong></td>')
                html.append(f'<td>{val_incra}</td><td>{val_projeto}</td><td>{status_texto}</td>')
                html.append('</tr>')

            html.append('</tbody></table>')

            # Tabela de Segmento Vante para esta linha
            html.append('<h4 class="subtitulo-secao">Segmento Vante</h4>')
            html.append('<table class="tabela-linha">')
            html.append('<thead><tr>')
            html.append('<th>Campo</th><th>INCRA</th><th>Projeto</th><th>Status</th>')
            html.append('</tr></thead><tbody>')

            # Dados do Segmento
            cod_seg_incra = self._limpar_string(incra_row[4] if len(incra_row) > 4 else "")
            cod_seg_projeto = self._limpar_string(projeto_row[4] if len(projeto_row) > 4 else "")

            azim_incra = self._limpar_string(incra_row[5] if len(incra_row) > 5 else "")
            azim_projeto = self._limpar_string(projeto_row[5] if len(projeto_row) > 5 else "")

            dist_incra = self._limpar_string(incra_row[6] if len(incra_row) > 6 else "")
            dist_projeto = self._limpar_string(projeto_row[6] if len(projeto_row) > 6 else "")

            campos_segmento = [
                ("Código", cod_seg_incra, cod_seg_projeto),
                ("Azimute", azim_incra, azim_projeto),
                ("Distância", dist_incra, dist_projeto)
            ]

            for campo, val_incra, val_projeto in campos_segmento:
                status_classe = "identico" if val_incra == val_projeto else "diferente"
                status_texto = "✅ Idêntico" if val_incra == val_projeto else "❌ Diferente"

                if val_incra == val_projeto:
                    identicos_segmento += 1
                else:
                    diferencas_segmento += 1

                html.append(f'<tr class="{status_classe}">')
                html.append(f'<td><strong>{campo}</strong></td>')
                html.append(f'<td>{val_incra}</td><td>{val_projeto}</td><td>{status_texto}</td>')
                html.append('</tr>')

            html.append('</tbody></table>')
            html.append('</div>')  # Fecha linha-grupo

        # RESUMO
        identicos_total = identicos_vertice + identicos_segmento
        diferencas_total = diferencas_vertice + diferencas_segmento

        html.append(f"""
        <div class="resumo-section">
            <h2>Resumo da Comparação</h2>
            <div class="resumo-grid">
                <div class="resumo-card">
                    <h4>Vértices</h4>
                    <p>Idênticos: <strong>{identicos_vertice}</strong></p>
                    <p>Diferentes: <strong>{diferencas_vertice}</strong></p>
                </div>
                <div class="resumo-card">
                    <h4>Segmentos Vante</h4>
                    <p>Idênticos: <strong>{identicos_segmento}</strong></p>
                    <p>Diferentes: <strong>{diferencas_segmento}</strong></p>
                </div>
            </div>
            <div class="resumo-total">
                <h4>Total Geral</h4>
                <p>Total de campos idênticos: <strong>{identicos_total}</strong></p>
                <p>Total de campos diferentes: <strong>{diferencas_total}</strong></p>
            </div>
        </div>
        <div class="footer">
            <p>Relatório gerado automaticamente pelo Sistema de Verificação INCRA v4.0</p>
            <p>{datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>
""")

        return "".join(html)

    def _salvar_e_abrir_relatorio(self, conteudo_html: str):
        """Salva relatório automaticamente e abre no navegador."""
        relatorios_dir = Path.home() / "Documentos" / "Relatórios INCRA"
        relatorios_dir.mkdir(parents=True, exist_ok=True)

        numero = self.numero_prenotacao.get()
        nome_arquivo = f"Relatório_INCRA_{numero}.html"
        caminho_completo = relatorios_dir / nome_arquivo

        with open(caminho_completo, 'w', encoding='utf-8') as f:
            f.write(conteudo_html)

        webbrowser.open(f'file://{caminho_completo}')

        self._atualizar_status(f"✅ Relatório salvo: {caminho_completo}")

    def _limpar_pasta_temp(self):
        """Remove a pasta temporária conferencia_geo_temp."""
        try:
            # Pasta em Downloads
            pasta_downloads = Path.home() / "Downloads" / "conferencia_geo_temp"
            if pasta_downloads.exists():
                shutil.rmtree(pasta_downloads)
                self._atualizar_status("🗑️ Pasta temporária em Downloads removida")

            # Pasta em temp do sistema
            pasta_temp = Path(tempfile.gettempdir()) / "conferencia_geo"
            if pasta_temp.exists():
                shutil.rmtree(pasta_temp)
                self._atualizar_status("🗑️ Pasta temporária do sistema removida")

        except Exception as e:
            print(f"Erro ao limpar pasta temp: {e}")

    def _limpar_tudo(self):
        """Limpa todos os campos e reseta o estado da aplicação."""
        resposta = messagebox.askyesno(
            "Confirmar Limpeza",
            "Deseja limpar todos os dados e começar uma nova comparação?\n\n" +
            "Isso irá:\n" +
            "• Limpar todos os campos\n" +
            "• Remover arquivos temporários\n" +
            "• Resetar a interface"
        )

        if resposta:
            try:
                # Limpar variáveis
                self.incra_path.set("")
                self.projeto_path.set("")
                self.numero_prenotacao.set("")
                self.incra_paginas.set("")
                self.projeto_paginas.set("")
                self.incra_anexo_path.set("")
                self.projeto_anexo_path.set("")

                # Resetar labels de anexo
                if hasattr(self, 'incra_anexo_label'):
                    self.incra_anexo_label.config(
                        text="Nenhum arquivo selecionado",
                        fg=self.colors['text_medium']
                    )
                if hasattr(self, 'projeto_anexo_label'):
                    self.projeto_anexo_label.config(
                        text="Nenhum arquivo selecionado",
                        fg=self.colors['text_medium']
                    )

                # Limpar dados extraídos
                self.incra_excel_path = None
                self.projeto_excel_path = None
                self.incra_data = None
                self.projeto_data = None
                self.pdf_extraido_incra = None
                self.pdf_extraido_projeto = None
                self.preview_incra_image = None
                self.preview_projeto_image = None

                # Esconder preview frame se estiver visível
                if hasattr(self, 'preview_frame'):
                    self.preview_frame.pack_forget()

                # Limpar área de resultados
                self.resultado_text.delete(1.0, tk.END)
                self.resultado_text.insert(1.0, "Interface limpa. Pronto para nova comparação.")

                # Limpar pasta temporária
                self._limpar_pasta_temp()

                # Habilitar botões
                self._habilitar_botoes()

                # Atualizar status
                self._atualizar_status("✨ Interface limpa! Pronto para nova comparação.")

                messagebox.showinfo("Sucesso", "Interface limpa com sucesso!")

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao limpar interface:\n\n{str(e)}")

    def _ao_fechar_programa(self):
        """Executado ao fechar o programa - limpa arquivos temporários."""
        try:
            # Limpar pasta temporária
            self._limpar_pasta_temp()
        except:
            pass
        finally:
            # Fechar o programa
            self.root.destroy()

    def _mostrar_resumo_no_texto(self):
        """Mostra resumo simplificado na área de texto."""
        self.resultado_text.delete(1.0, tk.END)

        resumo = f"""
╔════════════════════════════════════════════════════════════════╗
║          COMPARAÇÃO CONCLUÍDA COM SUCESSO                      ║
╚════════════════════════════════════════════════════════════════╝

📋 Número de Prenotação: {self.numero_prenotacao.get()}
📅 Data: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}

✅ O relatório HTML completo foi gerado e aberto automaticamente.
📁 Local: Documentos\\Relatórios INCRA\\Relatório_INCRA_{self.numero_prenotacao.get()}.html

💡 Consulte o relatório HTML para ver todos os detalhes da comparação.
"""

        self.resultado_text.insert(1.0, resumo)


def main():
    """Função principal."""
    root = tk.Tk()
    app = VerificadorGeorreferenciamento(root)
    root.mainloop()


if __name__ == "__main__":
    main()