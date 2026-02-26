import customtkinter as ctk
import pandas as pd
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import time
import logging
from datetime import datetime
import os
from PIL import Image, ImageDraw
import traceback
import threading
import unicodedata
import re
from typing import Optional


def sanitizar_nome_arquivo(nome: str) -> str:
    """Remove caracteres inválidos para nomes de arquivo no Windows."""
    nome = unicodedata.normalize('NFKD', nome).encode('ascii', 'ignore').decode('ascii')
    nome = re.sub(r'[\\/:*?"<>|]', '', nome)
    nome = nome.strip()
    if len(nome) > 200:
        nome = nome[:200]
    return nome


# Handler de log separado da classe principal
class GUILogHandler(logging.Handler):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui

    def emit(self, record):
        msg = self.format(record)
        self.gui.adicionar_log(msg, record.levelno)


class AutomacaoGUI:
    # Cores do tema
    CORES = {
        'sucesso': '#2ECC71',
        'erro': '#E74C3C',
        'aviso': '#F39C12',
        'info': '#3498DB',
        'texto': '#ECF0F1',
        'fundo_card': '#2C3E50',
        'fundo_escuro': '#1A252F',
        'destaque': '#1ABC9C',
        'processando': '#9B59B6',
    }

    def __init__(self):
        # Configuração do tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Folha de Ponto")
        self.window.geometry("800x550")
        self.window.minsize(750, 500)
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        # Flag para controle de execução
        self.executando = False
        self.pausado = False
        self.thread_automacao = None

        # Estatísticas
        self.stats = {
            'processados': 0,
            'sucesso': 0,
            'erros': 0,
            'tempo_inicio': None
        }

        # DataFrame carregado
        self.df_carregado = None

        # Configurar ícone
        self.set_window_icon()

        # Criar diretório de logs se não existir
        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)

        # Configurar logging para arquivos
        self.setup_file_logging()

        # Variáveis
        self.arquivo_excel = ctk.StringVar()
        self.linha_inicial = ctk.StringVar(value="2")
        self.status_var = ctk.StringVar(value="Aguardando início...")
        self.empresa_atual_var = ctk.StringVar(value="-")

        # Logger
        self.logger = logging.getLogger('AutomacaoDominio')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        # Adicionar GUIHandler
        self.gui_handler = GUILogHandler(self)
        formatter = logging.Formatter('%(message)s')
        self.gui_handler.setFormatter(formatter)
        self.logger.addHandler(self.gui_handler)

        self.criar_interface()

    def setup_file_logging(self):
        """Configura o logging para arquivos"""
        data_atual = datetime.now().strftime("%Y-%m-%d")

        # Logger de sucesso
        self.success_logger = logging.getLogger('SuccessLog')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            success_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'success_{data_atual}.log'),
                encoding='utf-8'
            )
            success_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.success_logger.addHandler(success_handler)

        # Logger de erro
        self.error_logger = logging.getLogger('ErrorLog')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            error_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'error_{data_atual}.log'),
                encoding='utf-8'
            )
            error_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.error_logger.addHandler(error_handler)

    def set_window_icon(self):
        """Configura o ícone da janela"""
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "assets", "favicon.ico")
            if os.name == 'nt' and os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

    def criar_interface(self):
        # Frame principal com grid
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)

        # === HEADER ===
        self.criar_header(main_frame)

        # === PAINEL DE CONFIGURAÇÃO ===
        self.criar_painel_config(main_frame)

        # === PAINEL DE ESTATÍSTICAS ===
        self.criar_painel_estatisticas(main_frame)

        # === ÁREA DE CONTEÚDO (Abas) ===
        self.criar_area_conteudo(main_frame)

    def criar_header(self, parent):
        """Cria o cabeçalho com título e status"""
        header_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header_frame.grid_columnconfigure(1, weight=1)

        # Ícone/Logo com fundo branco circular
        logo_path = os.path.join(os.path.dirname(__file__), "assets", "DomBot_New.png")
        if os.path.exists(logo_path):
            size = 66
            circle_size = 44
            bg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            circle_mask = Image.new("L", (circle_size, circle_size), 0)
            ImageDraw.Draw(circle_mask).ellipse((0, 0, circle_size - 1, circle_size - 1), fill=255)
            circle = Image.new("RGBA", (circle_size, circle_size), (255, 255, 255, 255))
            circle_offset = (size - circle_size) // 2
            bg.paste(circle, (circle_offset, circle_offset), circle_mask)
            original = Image.open(logo_path).convert("RGBA")
            original = original.resize((size, size), Image.LANCZOS)
            bg.paste(original, (0, 0), original)
            logo_image = ctk.CTkImage(light_image=bg, dark_image=bg, size=(size, size))
            ctk.CTkLabel(header_frame, image=logo_image, text="").grid(row=0, column=0, padx=10, pady=8)
        else:
            logo_frame = ctk.CTkFrame(header_frame, fg_color=self.CORES['destaque'],
                                       width=44, height=44, corner_radius=22)
            logo_frame.grid(row=0, column=0, padx=10, pady=8)
            logo_frame.grid_propagate(False)
            ctk.CTkLabel(logo_frame, text="🤖", font=("Segoe UI Emoji", 18)).place(relx=0.5, rely=0.5, anchor="center")

        # Título
        ctk.CTkLabel(
            header_frame,
            text="DomBot - Folha de Ponto",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.CORES['texto']
        ).grid(row=0, column=1, sticky="w", padx=5)

        # Status indicator
        self.status_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        self.status_frame.grid(row=0, column=2, padx=10)

        self.status_indicator = ctk.CTkFrame(
            self.status_frame,
            fg_color="#7F8C8D",
            width=10, height=10,
            corner_radius=5
        )
        self.status_indicator.pack(side="left", padx=(0, 6))

        self.status_label = ctk.CTkLabel(
            self.status_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=11),
            text_color="#95A5A6"
        )
        self.status_label.pack(side="left")

    def criar_painel_config(self, parent):
        """Cria o painel de configuração"""
        config_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        config_frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        config_frame.grid_columnconfigure(1, weight=1)

        # Linha única com tudo
        inner_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        inner_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        inner_frame.grid_columnconfigure(1, weight=1)

        # Arquivo Excel
        ctk.CTkLabel(
            inner_frame, text="📁", font=ctk.CTkFont(size=14)
        ).grid(row=0, column=0, padx=(0, 5))

        self.entry_arquivo = ctk.CTkEntry(
            inner_frame,
            textvariable=self.arquivo_excel,
            placeholder_text="Selecione o arquivo Excel...",
            height=32,
            font=ctk.CTkFont(size=11)
        )
        self.entry_arquivo.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            inner_frame, text="Procurar", command=self.selecionar_arquivo,
            width=80, height=32, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9"
        ).grid(row=0, column=2, padx=(0, 15))

        # Linha inicial
        ctk.CTkLabel(
            inner_frame, text="Linha:", font=ctk.CTkFont(size=11), text_color="#BDC3C7"
        ).grid(row=0, column=3, padx=(0, 3))

        self.entry_linha = ctk.CTkEntry(
            inner_frame, textvariable=self.linha_inicial,
            width=50, height=32, font=ctk.CTkFont(size=11), justify="center"
        )
        self.entry_linha.grid(row=0, column=4, padx=(0, 15))

        # Botões de controle
        self.btn_iniciar = ctk.CTkButton(
            inner_frame, text="▶ Iniciar", command=self.iniciar_automacao_thread,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60"
        )
        self.btn_iniciar.grid(row=0, column=5, padx=3)

        self.btn_pausar = ctk.CTkButton(
            inner_frame, text="⏸ Pausar", command=self.pausar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['aviso'], hover_color="#E67E22", state="disabled"
        )
        self.btn_pausar.grid(row=0, column=6, padx=3)

        self.btn_parar = ctk.CTkButton(
            inner_frame, text="⏹ Parar", command=self.parar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['erro'], hover_color="#C0392B", state="disabled"
        )
        self.btn_parar.grid(row=0, column=7, padx=(3, 0))

    def criar_painel_estatisticas(self, parent):
        """Cria o painel de estatísticas"""
        stats_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        stats_frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))

        # Grid para os cards de estatísticas
        for i in range(5):
            stats_frame.grid_columnconfigure(i, weight=1)

        # Cards de estatísticas
        self.criar_stat_card(stats_frame, 0, "📊", "Total", "total_label", "0")
        self.criar_stat_card(stats_frame, 1, "✅", "Sucesso", "sucesso_label", "0", self.CORES['sucesso'])
        self.criar_stat_card(stats_frame, 2, "❌", "Erros", "erros_label", "0", self.CORES['erro'])
        self.criar_stat_card(stats_frame, 3, "🏢", "Empresa", "empresa_label", "-", self.CORES['info'])
        self.criar_stat_card(stats_frame, 4, "⏱", "Tempo", "tempo_label", "00:00:00", self.CORES['aviso'])

        # Barra de progresso
        progress_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
        progress_frame.grid(row=1, column=0, columnspan=5, sticky="ew", padx=10, pady=(2, 8))
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, height=6, corner_radius=3, progress_color=self.CORES['destaque']
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_frame, text="0%", font=ctk.CTkFont(size=10), text_color="#95A5A6"
        )
        self.progress_label.grid(row=0, column=1, padx=(8, 0))

    def criar_stat_card(self, parent, col, icon, titulo, attr_name, valor_inicial, cor=None):
        """Cria um card de estatística"""
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.grid(row=0, column=col, padx=5, pady=8)

        ctk.CTkLabel(
            card, text=f"{icon} {titulo}", font=ctk.CTkFont(size=10), text_color="#7F8C8D"
        ).pack()

        label = ctk.CTkLabel(
            card, text=valor_inicial, font=ctk.CTkFont(size=14, weight="bold"),
            text_color=cor if cor else self.CORES['texto']
        )
        label.pack()

        setattr(self, attr_name, label)

    def criar_area_conteudo(self, parent):
        """Cria a área de conteúdo com abas"""
        self.tabview = ctk.CTkTabview(
            parent, fg_color=self.CORES['fundo_card'],
            segmented_button_fg_color=self.CORES['fundo_escuro'],
            segmented_button_selected_color=self.CORES['destaque'],
            corner_radius=8, height=25
        )
        self.tabview.grid(row=3, column=0, sticky="nsew")

        tab_logs = self.tabview.add("📋 Logs")
        tab_preview = self.tabview.add("📊 Preview")

        self.criar_aba_logs(tab_logs)
        self.criar_aba_preview(tab_preview)

    def criar_aba_logs(self, parent):
        """Cria a aba de logs"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        log_container = ctk.CTkFrame(parent, fg_color="transparent")
        log_container.grid(row=0, column=0, sticky="nsew", padx=3, pady=3)
        log_container.grid_columnconfigure(0, weight=1)
        log_container.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            log_container, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Configurar tags de cores
        self.log_text._textbox.tag_config("sucesso", foreground=self.CORES['sucesso'])
        self.log_text._textbox.tag_config("erro", foreground=self.CORES['erro'])
        self.log_text._textbox.tag_config("aviso", foreground=self.CORES['aviso'])
        self.log_text._textbox.tag_config("info", foreground=self.CORES['info'])
        self.log_text._textbox.tag_config("processando", foreground=self.CORES['processando'])

        # Botões de controle do log
        btn_frame = ctk.CTkFrame(log_container, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(5, 0))

        ctk.CTkButton(
            btn_frame, text="🗑 Limpar", command=self.limpar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame, text="💾 Exportar", command=self.exportar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="left", padx=8)

    def criar_aba_preview(self, parent):
        """Cria a aba de preview do Excel"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        info_frame = ctk.CTkFrame(parent, fg_color="transparent")
        info_frame.grid(row=0, column=0, sticky="ew", padx=3, pady=3)

        self.preview_info_label = ctk.CTkLabel(
            info_frame, text="Nenhum arquivo carregado",
            font=ctk.CTkFont(size=11), text_color="#95A5A6"
        )
        self.preview_info_label.pack(side="left")

        ctk.CTkButton(
            info_frame, text="🔄 Recarregar", command=self.carregar_preview,
            width=85, height=24, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="right")

        self.preview_text = ctk.CTkTextbox(
            parent, font=ctk.CTkFont(family="Consolas", size=10),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6
        )
        self.preview_text.grid(row=1, column=0, sticky="nsew", padx=3, pady=(0, 3))

    def selecionar_arquivo(self):
        """Abre diálogo para selecionar arquivo Excel"""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o arquivo Excel"
        )
        if filename:
            self.arquivo_excel.set(filename)
            self.adicionar_log(f"Arquivo selecionado: {os.path.basename(filename)}", logging.INFO, "info")
            self.carregar_preview()

    def carregar_preview(self):
        """Carrega preview do arquivo Excel"""
        if not self.arquivo_excel.get():
            return

        try:
            self.df_carregado = pd.read_excel(self.arquivo_excel.get())
            total_linhas = len(self.df_carregado)

            # Atualizar info
            self.preview_info_label.configure(
                text=f"📄 {os.path.basename(self.arquivo_excel.get())} | {total_linhas} linhas | Colunas: {', '.join(self.df_carregado.columns[:5])}..."
            )

            # Atualizar estatística de total
            self.total_label.configure(text=str(total_linhas))

            # Mostrar preview
            self.preview_text.delete("1.0", "end")

            # Cabeçalho
            header = " | ".join([f"{col:^15}" for col in self.df_carregado.columns[:6]])
            self.preview_text.insert("end", f"{'─' * len(header)}\n")
            self.preview_text.insert("end", f"{header}\n")
            self.preview_text.insert("end", f"{'─' * len(header)}\n")

            # Dados (primeiras 50 linhas)
            for idx, row in self.df_carregado.head(50).iterrows():
                row_text = " | ".join([f"{str(val)[:15]:^15}" for val in row.values[:6]])
                self.preview_text.insert("end", f"{row_text}\n")

            if total_linhas > 50:
                self.preview_text.insert("end", f"\n... e mais {total_linhas - 50} linhas\n")

            self.adicionar_log(f"Preview carregado: {total_linhas} linhas encontradas", logging.INFO, "sucesso")

        except Exception as e:
            self.adicionar_log(f"Erro ao carregar preview: {str(e)}", logging.ERROR, "erro")

    def limpar_logs(self):
        """Limpa a área de logs"""
        self.log_text.delete("1.0", "end")
        self.adicionar_log("Log limpo", logging.INFO, "info")

    def exportar_logs(self):
        """Exporta logs para arquivo"""
        try:
            filename = ctk.filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfilename=f"logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get("1.0", "end"))
                self.adicionar_log(f"Logs exportados para: {filename}", logging.INFO, "sucesso")
        except Exception as e:
            self.adicionar_log(f"Erro ao exportar logs: {str(e)}", logging.ERROR, "erro")

    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso"""
        porcentagem = atual / total if total > 0 else 0
        self.progress_bar.set(porcentagem)
        self.progress_label.configure(text=f"{porcentagem * 100:.1f}%")
        self.status_var.set(f"Processando: {atual}/{total}")
        self.window.update_idletasks()

    def atualizar_estatisticas(self, sucesso=None, erro=None, empresa=None):
        """Atualiza os cards de estatísticas"""
        if sucesso is not None:
            self.stats['sucesso'] += 1 if sucesso else 0
            self.sucesso_label.configure(text=str(self.stats['sucesso']))

        if erro is not None:
            self.stats['erros'] += 1 if erro else 0
            self.erros_label.configure(text=str(self.stats['erros']))

        if empresa is not None:
            self.empresa_label.configure(text=str(empresa)[:20])

        self.stats['processados'] = self.stats['sucesso'] + self.stats['erros']

    def atualizar_tempo(self):
        """Atualiza o tempo decorrido"""
        if self.stats['tempo_inicio'] and self.executando:
            elapsed = datetime.now() - self.stats['tempo_inicio']
            hours, remainder = divmod(int(elapsed.total_seconds()), 3600)
            minutes, seconds = divmod(remainder, 60)
            self.tempo_label.configure(text=f"{hours:02d}:{minutes:02d}:{seconds:02d}")
            self.window.after(1000, self.atualizar_tempo)

    def atualizar_status_indicator(self, status):
        """Atualiza o indicador de status visual"""
        cores = {
            'aguardando': '#7F8C8D',
            'executando': self.CORES['sucesso'],
            'pausado': self.CORES['aviso'],
            'erro': self.CORES['erro'],
            'concluido': self.CORES['info']
        }
        self.status_indicator.configure(fg_color=cores.get(status, '#7F8C8D'))

    def adicionar_log(self, mensagem, level=logging.INFO, tag=None):
        """Adiciona mensagem ao log visual com cores"""
        timestamp = datetime.now().strftime('%H:%M:%S')

        # Determinar tag baseado no nível se não especificado
        if tag is None:
            if level >= logging.ERROR:
                tag = "erro"
            elif level >= logging.WARNING:
                tag = "aviso"
            elif "sucesso" in mensagem.lower() or "processad" in mensagem.lower():
                tag = "sucesso"
            else:
                tag = "info"

        # Prefixo visual
        prefixos = {
            "sucesso": "✅",
            "erro": "❌",
            "aviso": "⚠️",
            "info": "ℹ️",
            "processando": "⏳"
        }
        prefixo = prefixos.get(tag, "•")

        # Inserir mensagem
        self.log_text.insert("end", f"[{timestamp}] {prefixo} ", tag)
        self.log_text.insert("end", f"{mensagem}\n", tag)
        self.log_text.see("end")
        self.window.update_idletasks()

    def iniciar_automacao_thread(self):
        """Inicia a automação em uma thread separada"""
        if self.executando:
            self.adicionar_log("Automação já em execução", logging.WARNING, "aviso")
            return

        # Resetar estatísticas
        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'tempo_inicio': datetime.now()}
        self.sucesso_label.configure(text="0")
        self.erros_label.configure(text="0")

        self.thread_automacao = threading.Thread(target=self.iniciar_automacao)
        self.thread_automacao.daemon = True
        self.thread_automacao.start()

        # Atualiza interface
        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_parar.configure(state="normal")
        self.atualizar_status_indicator('executando')

        # Iniciar timer
        self.atualizar_tempo()

    def pausar_automacao(self):
        """Pausa/retoma a automação"""
        if self.executando:
            self.pausado = not self.pausado
            if self.pausado:
                self.btn_pausar.configure(text="▶  Retomar")
                self.status_var.set("Pausado")
                self.atualizar_status_indicator('pausado')
                self.adicionar_log("Automação pausada", logging.INFO, "aviso")
            else:
                self.btn_pausar.configure(text="⏸  Pausar")
                self.status_var.set("Em execução...")
                self.atualizar_status_indicator('executando')
                self.adicionar_log("Automação retomada", logging.INFO, "info")

    def parar_automacao(self):
        """Para a execução da automação"""
        if self.executando:
            self.executando = False
            self.pausado = False
            self.adicionar_log("Solicitação de parada enviada. Aguardando conclusão...", logging.INFO, "aviso")
            self.status_var.set("Interrompendo...")
            self.atualizar_status_indicator('erro')

    def ao_fechar(self):
        """Tratamento do fechamento da janela"""
        if self.executando:
            from tkinter import messagebox
            if messagebox.askyesno("Confirmação",
                                   "Existe uma automação em execução. Deseja realmente sair?"):
                self.executando = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()

    def iniciar_automacao(self):
        if not self.arquivo_excel.get():
            self.adicionar_log("Erro: Selecione um arquivo Excel", logging.ERROR, "erro")
            self.btn_iniciar.configure(state="normal")
            self.btn_pausar.configure(state="disabled")
            self.btn_parar.configure(state="disabled")
            return

        try:
            linha_inicial = int(self.linha_inicial.get())
            if linha_inicial < 2:
                raise ValueError("Linha inicial deve ser >= 2 (linha 1 é o cabeçalho)")
        except ValueError:
            self.adicionar_log("Erro: Linha inicial inválida (deve ser >= 2)", logging.ERROR, "erro")
            self.btn_iniciar.configure(state="normal")
            self.btn_pausar.configure(state="disabled")
            self.btn_parar.configure(state="disabled")
            return

        self.adicionar_log("Iniciando automação...", logging.INFO, "processando")
        self.status_var.set("Em execução...")
        self.executando = True

        try:
            # Carregar Excel
            df = pd.read_excel(self.arquivo_excel.get())
            total_linhas = len(df) - (linha_inicial - 2)
            self.adicionar_log(f"Arquivo Excel carregado com {total_linhas} linhas para processar", logging.INFO, "info")
            self.total_label.configure(text=str(total_linhas))

            # Resetar barra de progresso
            self.progress_bar.set(0)

            # Iniciar automação
            automacao = DominioAutomation(self.logger, self)

            # Conectar ao Domínio
            if not automacao.connect_to_dominio():
                erro_msg = "Erro: Não foi possível conectar ao Domínio"
                self.adicionar_log(erro_msg, logging.ERROR, "erro")
                self.error_logger.error(erro_msg)
                return

            # Processar linhas
            for idx, (index, row) in enumerate(df.iloc[linha_inicial - 2:].iterrows()):
                # Verificar se deve parar
                if not self.executando:
                    self.adicionar_log("Automação interrompida pelo usuário", logging.INFO, "aviso")
                    break

                # Verificar pausa
                while self.pausado and self.executando:
                    time.sleep(0.5)

                # Atualizar progresso
                self.atualizar_progresso(idx + 1, total_linhas)

                # Atualizar empresa atual
                empresa_nome = row.get('EMPRESA', row.get('Nº', 'N/A'))
                self.atualizar_estatisticas(empresa=empresa_nome)

                try:
                    log_msg = (f"Linha {index + 1} - Nº {row['Nº']} - "
                               f"EMPRESA: {row.get('EMPRESA', 'N/A')}")

                    success = automacao.processar_linha(row, index)

                    # Log do resultado
                    if success:
                        self.success_logger.info(f"{log_msg} - Enviado com sucesso")
                        self.adicionar_log(f"Linha {index + 1} processada com sucesso", logging.INFO, "sucesso")
                        self.atualizar_estatisticas(sucesso=True)
                    else:
                        self.error_logger.error(f"{log_msg} - Erro no envio")
                        self.adicionar_log(f"Processo interrompido na linha {index + 1}", logging.ERROR, "erro")
                        self.atualizar_estatisticas(erro=True)
                        break

                    time.sleep(2)

                except Exception as e:
                    erro_msg = f"{log_msg} - Erro: {str(e)}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg, logging.ERROR, "erro")
                    self.adicionar_log(f"Detalhes do erro: {traceback.format_exc()}", logging.ERROR, "erro")
                    self.atualizar_estatisticas(erro=True)
                    break

            self.status_var.set("Processamento concluído")
            self.progress_bar.set(1.0)
            self.progress_label.configure(text="100%")
            self.atualizar_status_indicator('concluido')
            self.adicionar_log(
                f"Concluído! Sucesso: {self.stats['sucesso']} | Erros: {self.stats['erros']}",
                logging.INFO, "sucesso"
            )

        except Exception as e:
            erro_msg = f"Erro crítico: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg, logging.ERROR, "erro")
            self.adicionar_log(f"Detalhes do erro: {traceback.format_exc()}", logging.ERROR, "erro")
            self.status_var.set("Erro no processamento")
            self.atualizar_status_indicator('erro')
        finally:
            self.executando = False
            self.pausado = False
            self.btn_iniciar.configure(state="normal")
            self.btn_pausar.configure(state="disabled", text="⏸  Pausar")
            self.btn_parar.configure(state="disabled")

    def executar(self):
        self.window.mainloop()


class DominioAutomation:
    def __init__(self, logger, gui):
        timings.Timings.window_find_timeout = 20
        self.app = None
        self.main_window = None
        self.logger = logger
        self.gui = gui

    def log(self, message):
        self.logger.info(message)

    def should_stop(self) -> bool:
        """Verifica se deve parar a execução"""
        return not self.gui.executando

    def check_pause(self):
        """Verifica e aguarda se pausado"""
        while self.gui.pausado and self.gui.executando:
            time.sleep(0.5)

    def smart_sleep(self, seconds: float) -> bool:
        """Sleep interruptível que verifica pausa/parada"""
        interval = 0.15
        elapsed = 0.0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            if self.should_stop():
                return False
            sleep_time = min(interval, seconds - elapsed)
            time.sleep(sleep_time)
            elapsed += sleep_time
        return True

    def wait_for_condition(self, condition_fn, timeout: float = 30, poll_interval: float = 0.15, description: str = "") -> bool:
        """Polls condition_fn() até retornar True, ou timeout."""
        start = time.time()
        while time.time() - start < timeout:
            if self.should_stop():
                return False
            self.check_pause()
            try:
                if condition_fn():
                    if description:
                        self.log(f"{description} - concluido em {time.time() - start:.1f}s")
                    return True
            except Exception:
                pass
            time.sleep(poll_interval)
        if description:
            self.log(f"{description} - timeout apos {timeout}s")
        return False

    def _window_exists(self, title: str, class_name: str) -> bool:
        """Verifica se janela com título/classe existe via win32gui (rápido)."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if (win32gui.GetWindowText(hwnd) == title and
                                win32gui.GetClassName(hwnd) == class_name):
                            result[0] = True
                            return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _any_error_dialog_visible(self) -> bool:
        """Verifica se há diálogo de erro visível via win32gui."""
        error_keywords = ("erro", "aviso", "atenção", "alerta", "warning", "error", "informação")
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        cls = win32gui.GetClassName(hwnd)
                        if cls == "#32770":
                            title = win32gui.GetWindowText(hwnd).lower()
                            for kw in error_keywords:
                                if kw in title:
                                    result[0] = True
                                    return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _is_connection_alive(self) -> bool:
        """Verifica se a conexão pywinauto ainda é válida."""
        if self.app is None or self.main_window is None:
            return False
        try:
            hwnd = self.main_window.handle
            if not win32gui.IsWindow(hwnd):
                return False
            win32gui.GetWindowText(hwnd)
            return True
        except Exception:
            return False

    def handle_error_dialogs(self) -> bool:
        """Trata diálogos de erro que podem aparecer.
        Retorna True se deve continuar, False se deve abortar."""
        try:
            error_titles_lower = {"erro", "erro léxico", "aviso", "atenção",
                                  "informação", "alerta", "warning", "error"}

            found_hwnd = None
            found_title = None

            def enum_callback(hwnd, _):
                nonlocal found_hwnd, found_title
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if not title:
                        return True
                    title_lower = title.strip().lower()
                    for err_title in error_titles_lower:
                        if title_lower == err_title or err_title in title_lower:
                            if win32gui.GetClassName(hwnd) == "#32770":
                                found_hwnd = hwnd
                                found_title = title
                                return False
                except Exception:
                    pass
                return True

            win32gui.EnumWindows(enum_callback, None)

            if found_hwnd is None:
                return True

            try:
                error_window = self.app.window(handle=found_hwnd)
            except Exception:
                win32gui.SetForegroundWindow(found_hwnd)
                send_keys('{ENTER}')
                time.sleep(0.3)
                return True

            message = ""
            try:
                message = error_window.window_text()
                try:
                    static_texts = error_window.children(class_name="Static")
                    for static in static_texts:
                        text = static.window_text()
                        if text:
                            message += " " + text
                except Exception:
                    pass
            except Exception:
                pass

            self.log(f"Diálogo detectado: '{found_title}' - {message[:100] if message else 'sem mensagem'}")

            mensagens_continuar = [
                "sem dados para emitir",
                "nenhum registro encontrado",
                "não há dados",
                "registro não encontrado"
            ]

            message_lower = message.lower()
            for msg in mensagens_continuar:
                if msg in message_lower:
                    self.log(f"Aviso não crítico: {msg}")
                    error_window.set_focus()
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    for _ in range(4):
                        send_keys('{ESC}')
                        time.sleep(0.5)
                    return False

            if "léxico" in found_title.lower():
                self.log("Erro léxico detectado, fechando...")
                error_window.set_focus()
                for _ in range(3):
                    send_keys('{ESC}')
                    time.sleep(0.5)
                return True

            self.log(f"Fechando diálogo '{found_title}'...")
            error_window.set_focus()
            time.sleep(0.2)

            try:
                ok_button = error_window.child_window(title="OK", class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    time.sleep(0.5)
                    if found_title.lower() in ("erro", "aviso"):
                        return False
                    return True
            except Exception:
                pass

            send_keys('{ENTER}')
            time.sleep(0.5)

            try:
                if error_window.exists():
                    send_keys('{ESC}')
                    time.sleep(0.3)
            except Exception:
                pass

            if found_title.lower() in ("erro", "aviso"):
                return False

            return True

        except Exception as e:
            self.log(f"Exceção ao verificar diálogos: {str(e)}")
            return True

    def cleanup_windows(self):
        """Limpa e fecha janelas abertas"""
        try:
            self.log("Limpando janelas")
            self.main_window.set_focus()

            for _ in range(4):
                send_keys('{ESC}')
                time.sleep(0.5)

            try:
                relatorio_window = self.main_window.child_window(
                    title="Gerenciador de Relatórios",
                    class_name="FNWND3190"
                )
                if relatorio_window.exists() and relatorio_window.is_visible():
                    self.log("Fechando Gerenciador de Relatórios restante")
                    send_keys('{ESC}')
                    time.sleep(0.5)
            except Exception:
                pass

        except Exception as e:
            self.log(f"Erro durante limpeza: {str(e)}")

    def find_dominio_window(self) -> Optional[int]:
        """Encontra a janela do Domínio Folha"""
        try:
            self.log("Procurando janela do Domínio Folha...")

            try:
                all_windows = findwindows.find_windows()
                self.log(f"Total de janelas abertas: {len(all_windows)}")

                for hwnd in all_windows:
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        if "Domínio" in title and title:
                            self.log(f"Janela encontrada: '{title}'")
                            if "Folha" in title:
                                self.log(f"Janela do Domínio Folha localizada!")
                                return hwnd
                    except Exception:
                        continue
            except Exception as e:
                self.log(f"Erro ao listar janelas: {str(e)}")

            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                self.log(f"Janela do Domínio encontrada via regex (total: {len(windows)})")
                return windows[0]

            self.log("Nenhuma janela do Domínio Folha encontrada")
            return None
        except Exception as e:
            self.log(f"Erro ao procurar janela do Domínio: {str(e)}")
            self.log(f"Traceback: {traceback.format_exc()}")
            return None

    def connect_to_dominio(self):
        try:
            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)

            self.app = Application(backend="uia").connect(handle=handle)
            self.main_window = self.app.window(handle=handle)
            return True
        except Exception as e:
            self.log(f"Erro ao conectar ao Domínio Folha: {str(e)}")
            return False

    def wait_for_window(self, titulo, timeout=30):
        """Espera por uma janela com o título especificado"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                window = self.app.window(title=titulo)
                if window.exists():
                    return window
            except Exception:
                pass
            time.sleep(0.5)
        raise TimeoutError(f"Timeout esperando pela janela: {titulo}")

    def wait_and_check_window_closed(self, window, window_title, timeout=30):
        """Espera até que uma janela seja fechada"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            if not window.exists() or not window.is_visible():
                self.log(f"Janela '{window_title}' fechada com sucesso")
                return True
            time.sleep(0.5)

        self.log(f"Aviso: Tempo máximo de espera atingido para fechamento da janela '{window_title}'")
        return False

    def processar_linha(self, row, index):
        try:
            self.log(f"Processando linha {index + 1}")

            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                if not self.smart_sleep(1):
                    return False

            win32gui.SetForegroundWindow(handle)
            if not self.smart_sleep(0.5):
                return False

            app = Application(backend="uia").connect(handle=handle)
            main_window = app.window(handle=handle)

            main_window.set_focus()
            if not self.smart_sleep(0.5):
                return False

            self.log("Enviando F8 para troca de empresas")
            send_keys('{F8}')
            if not self.smart_sleep(1.5):
                return False

            try:
                troca_empresas_window = None
                max_attempts = 3

                for attempt in range(max_attempts):
                    try:
                        troca_empresas_window = main_window.child_window(
                            title="Troca de empresas",
                            class_name="FNWND3190"
                        )

                        if troca_empresas_window.exists():
                            break

                        troca_empresas_windows = main_window.children(title="Troca de empresas")
                        if troca_empresas_windows:
                            troca_empresas_window = troca_empresas_windows[0]
                            break
                    except Exception:
                        if attempt == max_attempts - 1:
                            self.log("Janela 'Troca de empresas' não encontrada após várias tentativas.")
                            return False
                        if not self.smart_sleep(1):
                            return False

                if not troca_empresas_window:
                    self.log("Janela 'Troca de empresas' não encontrada.")
                    return False
            except Exception as e:
                self.log(f"Erro ao localizar janela 'Troca de empresas': {str(e)}")
                return False

            self.log("Janela 'Troca de empresas' visível")

            empresa_num = str(int(row['Nº']))
            self.log(f"Enviando código da empresa: {empresa_num}")
            send_keys(empresa_num)
            if not self.smart_sleep(0.3):
                return False

            send_keys('{ENTER}')
            if not self.smart_sleep(6):
                return False

            self.wait_and_check_window_closed(troca_empresas_window, "Troca de empresas")

            try:
                aviso_window = main_window.child_window(
                    title="Avisos de Vencimento",
                    class_name="FNWND3190"
                )

                if aviso_window.exists() and aviso_window.is_visible():
                    self.log("Janela 'Avisos de Vencimento' encontrada - executando fechamento")
                    aviso_window.set_focus()
                    send_keys('{ESC}')
                    self.smart_sleep(1)
                    send_keys('{ESC}')
                    self.log("ESCs executados para fechar 'Avisos de Vencimento'")
            except Exception:
                self.log("Nenhuma janela de 'Avisos de Vencimento' encontrada")

            self.log("Enviando comandos para acessar relatórios")
            main_window.set_focus()
            send_keys('%r')
            if not self.smart_sleep(0.5):
                return False
            send_keys('i')
            if not self.smart_sleep(0.5):
                return False
            send_keys('i')
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(1):
                return False

            try:
                max_attempts = 3
                relatorio_window = None

                for attempt in range(max_attempts):
                    try:
                        relatorio_window = main_window.child_window(
                            title="Gerenciador de Relatórios",
                            class_name="FNWND3190"
                        )

                        if relatorio_window.exists():
                            break
                    except Exception:
                        if attempt == max_attempts - 1:
                            self.log("Janela 'Gerenciador de Relatórios' não encontrada após várias tentativas.")
                            return False
                        if not self.smart_sleep(1):
                            return False

                if not relatorio_window or not relatorio_window.exists():
                    self.log("Janela 'Gerenciador de Relatórios' não encontrada.")
                    return False

                self.log("Gerenciador de Relatórios localizado")

                rel_app = Application(backend='uia').connect(handle=relatorio_window.handle)
                tree = rel_app.window(class_name="FNWND3190").child_window(class_name="PBTreeView32_100")

                try:
                    for _ in range(3):
                        send_keys('{f}')
                        time.sleep(0.5)

                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    for _ in range(3):
                        send_keys('{f}')
                        time.sleep(0.5)
                except Exception as e:
                    self.log(f"Erro ao navegar na árvore de relatórios: {str(e)}")
                    return False

                self.log("Preenchendo os campos de data")
                send_keys('{TAB}*')
                time.sleep(0.3)
                send_keys('{TAB}' + str(row['data inicio']))
                time.sleep(0.3)
                send_keys('{TAB}' + str(row['data final']))
                time.sleep(0.5)

                self.log("Clicando em executar relatório")
                button_executar = relatorio_window.child_window(auto_id="1007", class_name="Button")
                button_executar.click_input()
                if not self.smart_sleep(2):
                    return False

                # Verificar se apareceu erro durante a execução do relatório
                if self._any_error_dialog_visible():
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False

                # === PUBLICAÇÃO DO DOCUMENTO ===
                self.log("Clicando no ícone de publicação")
                main_window.set_focus()
                button_publicacao = main_window.child_window(auto_id="1006", class_name="FNUDO3190")
                button_publicacao.click_input()
                if not self.smart_sleep(1):
                    return False

                try:
                    # Aguardar janela de Publicação de Documentos com polling
                    if not self.wait_for_condition(
                        lambda: main_window.child_window(
                            title="Publicação de Documentos",
                            class_name="FNWNS3190"
                        ).exists(),
                        timeout=10,
                        poll_interval=0.2,
                        description="Aguardando janela Publicação de Documentos"
                    ):
                        self.log("Janela 'Publicação de Documentos' não encontrada.")
                        return False

                    pub_doc_window = main_window.child_window(
                        title="Publicação de Documentos",
                        class_name="FNWNS3190"
                    )

                    self.log("Janela de publicação localizada")

                    # Selecionar categoria no ComboBox
                    self.log("Selecionando categoria: Pessoal/Folha de Ponto")
                    combo_box = pub_doc_window.child_window(auto_id="1007", class_name="ComboBox")
                    combo_box.click_input()
                    time.sleep(0.5)
                    send_keys("Pessoal/Folha de Ponto{ENTER}")
                    time.sleep(0.5)

                    if self.should_stop():
                        return False

                    # Definir nome do documento
                    nome_pdf = sanitizar_nome_arquivo(str(row['nome pdf']))
                    self.log(f"Definindo nome do PDF: {nome_pdf}")
                    edit_field = pub_doc_window.child_window(auto_id="1014", class_name="Edit")
                    edit_field.set_text(nome_pdf)
                    time.sleep(0.3)

                    # Clicar em Gravar
                    self.log("Clicando em gravar")
                    button_gravar = pub_doc_window.child_window(auto_id="1016", class_name="Button")
                    button_gravar.click_input()

                    # Aguardar fechamento da janela de publicação
                    if not self.wait_for_condition(
                        lambda: not pub_doc_window.exists() or not pub_doc_window.is_visible(),
                        timeout=10,
                        poll_interval=0.2,
                        description="Aguardando fechamento da Publicação de Documentos"
                    ):
                        self.log("Timeout aguardando fechamento da publicação")
                        return False

                    self.log("Documento publicado com sucesso")

                    # === GERAÇÃO DO PDF ===
                    # Verificar e tratar janela de erro
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False

                    # Aguardar processamento da publicação
                    if not self.smart_sleep(3):
                        return False

                    self.log("Gerando PDF")
                    main_window.set_focus()
                    button_pdf = main_window.child_window(auto_id="1015", class_name="FNUDO3190")
                    button_pdf.click_input()

                    # Esperar janela "Salvar em PDF" ou diálogo de erro
                    if not self.wait_for_condition(
                        lambda: self._window_exists("Salvar em PDF", "#32770") or self._any_error_dialog_visible(),
                        timeout=10,
                        poll_interval=0.15,
                        description="Aguardando janela Salvar em PDF"
                    ):
                        self.log("Timeout aguardando janela de salvamento PDF")
                        return False

                    # Se apareceu diálogo de erro, apenas fechar e continuar
                    # (o Domínio pode mostrar erro de gravação mas a janela Salvar em PDF aparece depois)
                    if self._any_error_dialog_visible():
                        self.log("Diálogo de erro detectado após gerar PDF, fechando e continuando...")
                        try:
                            error_win = main_window.child_window(title="Erro", class_name="#32770")
                            if error_win.exists() and error_win.is_visible():
                                ok_btn = error_win.child_window(title="OK", class_name="Button")
                                if ok_btn.exists():
                                    ok_btn.click_input()
                                    time.sleep(0.5)
                                    self.log("Botão OK clicado no diálogo de erro")
                        except Exception as e:
                            self.log(f"Erro ao fechar diálogo: {e}")
                            send_keys('{ENTER}')
                            time.sleep(0.5)

                        # Aguardar a janela "Salvar em PDF" aparecer após fechar o erro
                        if not self.wait_for_condition(
                            lambda: self._window_exists("Salvar em PDF", "#32770"),
                            timeout=10,
                            poll_interval=0.15,
                            description="Aguardando janela Salvar em PDF após fechar erro"
                        ):
                            self.log("Janela Salvar em PDF não apareceu após fechar erro")
                            self.cleanup_windows()
                            return False

                    # Localizar janela de salvamento
                    self.log("Configurando salvamento do PDF")
                    try:
                        save_window = main_window.child_window(
                            title="Salvar em PDF",
                            class_name="#32770"
                        )

                        if not save_window.exists():
                            self.log("Janela de salvamento não encontrada")
                            return False

                        if self.should_stop():
                            return False
                        self.check_pause()

                        time.sleep(0.5)
                        name_field = save_window.child_window(auto_id="1148", class_name="Edit")
                        name_field.set_text(nome_pdf)
                        time.sleep(0.3)

                        if self.should_stop():
                            return False

                        self.log("Salvando PDF")
                        button_salvar = save_window.child_window(auto_id="1", class_name="Button")
                        button_salvar.click_input()

                        # Esperar janela de salvamento fechar
                        if not self.wait_for_condition(
                            lambda: not save_window.exists() or not save_window.is_visible(),
                            timeout=15,
                            poll_interval=0.2,
                            description="Aguardando salvamento do PDF"
                        ):
                            self.log("Timeout aguardando salvamento do PDF")
                            return False

                    except Exception as e:
                        self.log(f"Erro durante salvamento: {str(e)}")
                        return False

                    # Fechar visualização do relatório
                    self.log("Fechando janelas")
                    send_keys('{ESC}')
                    if not self.smart_sleep(1):
                        return False

                except Exception as e:
                    self.log(f"Erro na publicação: {str(e)}")
                    return False

                # Limpeza final - fechar janelas restantes
                self.cleanup_windows()

            except Exception as e:
                self.log(f"Erro ao interagir com o Gerenciador de Relatórios: {str(e)}")
                self.log(f"Detalhes do erro: {traceback.format_exc()}")
                return False

            self.log(f"Linha {index + 1} processada com sucesso")
            return True

        except Exception as e:
            self.log(f"Erro ao processar linha {index + 1}: {str(e)}")
            self.log(f"Detalhes do erro: {traceback.format_exc()}")
            return False


def main():
    gui = AutomacaoGUI()
    gui.executar()


if __name__ == "__main__":
    main()
