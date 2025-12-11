import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import face_recognition
import os
import datetime
from datetime import timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import re
import io
import threading
import queue
import traceback
import logging
import zipfile
import shutil
import tempfile
import numpy as np
import json
import subprocess
import sys
from collections import defaultdict
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import csv
from openpyxl import Workbook
from cairosvg import svg2png

# --- FUNÇÃO PARA CARREGAR ÍCONES ---
def load_icon(name, theme):
    color = THEMES[theme]["COLOR_TEXT_MAIN"]
    try:
        with open(f"assets/{name}.svg", "r") as f:
            svg_data = f.read()

        # Substitui a cor do ícone
        svg_data = svg_data.replace('stroke="currentColor"', f'stroke="{color}"')

        png_data = svg2png(bytestring=svg_data.encode('utf-8'))

        return ctk.CTkImage(Image.open(io.BytesIO(png_data)), size=(24, 24))
    except Exception as e:
        log_debug(f"Erro ao carregar ícone {name}: {e}")
        return None

# --- CONFIGURAÇÃO DE TEMAS ---
THEMES = {
    "Noturno": {
        "COLOR_BG": "#0B1120", # Azul Oxford
        "COLOR_CARD": "#1e293b",
        "COLOR_BORDER": "#334155",
        "COLOR_TEXT_MAIN": "#f1f5f9", # Branco Acinzentado/Bege
        "COLOR_TEXT_DIM": "#94a3b8",
        "COLOR_ACCENT": "#facc15", # Amarelo Estrela
        "COLOR_BTN_HOVER": "#eab308",
        "COLOR_DANGER": "#ef4444",
        "COLOR_INFO": "#3b82f6"
    },
    "Diurno": {
        "COLOR_BG": "#f8fafc", # Branco Nuvem
        "COLOR_CARD": "#ffffff",
        "COLOR_BORDER": "#e2e8f0",
        "COLOR_TEXT_MAIN": "#0f172a",
        "COLOR_TEXT_DIM": "#64748b",
        "COLOR_ACCENT": "#0ea5e9", # Azul Oceano
        "COLOR_BTN_HOVER": "#0284c7",
        "COLOR_DANGER": "#dc2626",
        "COLOR_INFO": "#2563eb"
    }
}

# LOG
logging.basicConfig(filename='log_debug.txt', level=logging.DEBUG, format='%(asctime)s - %(message)s', filemode='w')

def log_debug(msg):
    print(msg)
    logging.info(msg)

NOME_PLANILHA_GOOGLE = "PontoFuncionarios"
ARQUIVO_CONFIG = "config.json"
ARQUIVO_FUNCIONARIOS = "funcionarios.json"

# --- JANELA DE SELEÇÃO DE MÚLTIPLOS FUNCIONÁRIOS ---
class ToplevelSelecaoFuncionarios(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title("Seleção Personalizada")
        self.geometry("400x500")
        self.transient(master)
        self.grab_set()

        self.checkboxes = {}

        ctk.CTkLabel(self, text="Selecione os Funcionários", font=("Arial", 16, "bold")).pack(pady=10)

        scroll_frame = ctk.CTkScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=15, pady=5)

        funcionarios_ativos = sorted(
            [f for f in self.master.dados_funcionarios if f.get('status', 'ativo') == 'ativo'],
            key=lambda x: x['nome']
        )

        for func in funcionarios_ativos:
            frame_func = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            frame_func.pack(fill="x", pady=5)

            var = ctk.StringVar(value="off")
            checkbox = ctk.CTkCheckBox(frame_func, text=func['nome'], variable=var, onvalue=func['nome'], offvalue="off")
            checkbox.pack(side="left")
            self.checkboxes[func['nome']] = var

        btn_confirmar = ctk.CTkButton(self, text="Confirmar", command=self.confirmar)
        btn_confirmar.pack(pady=10)

    def confirmar(self):
        selecionados = [var.get() for var in self.checkboxes.values() if var.get() != "off"]

        if not selecionados:
            # Se nada for selecionado, reseta para "Todos"
            self.master.funcionarios_selecionados = ["Todos"]
            self.master.combo_funcionarios.set("Todos")
            self.destroy()
            return

        self.master.funcionarios_selecionados = selecionados

        # Atualiza o texto do combobox para refletir a seleção
        if len(selecionados) == 1:
            self.master.combo_funcionarios.set(selecionados[0])
        else:
            self.master.combo_funcionarios.set(f"{len(selecionados)} funcionários selecionados")

        self.destroy()


# --- JANELA DE CADASTRO/EDIÇÃO DE FUNCIONÁRIO ---
class ToplevelWindowFuncionario(ctk.CTkToplevel):
    def __init__(self, master, funcionario_data=None):
        super().__init__(master)

        self.master = master # Referência à janela principal
        self.funcionario_data = funcionario_data

        self.title("Cadastrar Novo Funcionário" if not funcionario_data else "Editar Funcionário")
        self.geometry("500x650")
        self.transient(master)
        self.grab_set()

        # --- VARIÁVEIS ---
        self.photo_path = None
        self.dados_extras_visiveis = False

        # --- LAYOUT ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # 1. ÁREA DA FOTO
        self.frame_foto = ctk.CTkFrame(self, fg_color="#2b2b2b", height=200)
        self.frame_foto.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        self.frame_foto.pack_propagate(False)
        self.frame_foto.drop_target_register(DND_FILES)
        self.frame_foto.dnd_bind('<<Drop>>', self.drop_image)

        self.lbl_foto_preview = ctk.CTkLabel(self.frame_foto, text="👤", font=("Arial", 100), text_color="gray")
        self.lbl_foto_preview.pack(expand=True)

        self.btn_add_foto = ctk.CTkButton(self.frame_foto, text="+", width=30, height=30, corner_radius=15, command=self.select_image)
        self.btn_add_foto.place(relx=0.95, rely=0.95, anchor="se")

        # 2. CAMPOS PRINCIPAIS
        frame_campos = ctk.CTkFrame(self, fg_color="transparent")
        frame_campos.grid(row=1, column=0, sticky="ew", padx=20)
        frame_campos.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame_campos, text="Nome:").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_nome = ctk.CTkEntry(frame_campos)
        self.entry_nome.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=5)

        ctk.CTkLabel(frame_campos, text="Salário (R$):").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_salario = ctk.CTkEntry(frame_campos)
        self.entry_salario.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)

        ctk.CTkLabel(frame_campos, text="Admissão:").grid(row=2, column=0, sticky="w", pady=5)
        self.entry_admissao = DateEntry(frame_campos, locale='pt_BR', date_pattern='dd/mm/yyyy')
        self.entry_admissao.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)

        # 3. BOTÃO DE DADOS EXTRAS
        self.btn_dados_extras = ctk.CTkButton(self, text="Mais Dados ▼", fg_color="transparent", command=self.toggle_dados_extras)
        self.btn_dados_extras.grid(row=2, column=0, sticky="w", padx=20, pady=10)

        # 4. FRAME DE DADOS EXTRAS (inicialmente oculto)
        self.frame_extra = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_extra.grid_columnconfigure(1, weight=1)
        # Não usamos .grid() aqui ainda

        ctk.CTkLabel(self.frame_extra, text="E-mail:").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_email = ctk.CTkEntry(self.frame_extra)
        self.entry_email.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=5)

        ctk.CTkLabel(self.frame_extra, text="Celular:").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_celular = ctk.CTkEntry(self.frame_extra)
        self.entry_celular.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)

        ctk.CTkLabel(self.frame_extra, text="CPF:").grid(row=2, column=0, sticky="w", pady=5)
        self.entry_cpf = ctk.CTkEntry(self.frame_extra)
        self.entry_cpf.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)

        ctk.CTkLabel(self.frame_extra, text="CTPS:").grid(row=3, column=0, sticky="w", pady=5)
        self.entry_ctps = ctk.CTkEntry(self.frame_extra)
        self.entry_ctps.grid(row=3, column=1, sticky="ew", padx=(10, 0), pady=5)

        # 5. BOTÕES DE AÇÃO
        frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        frame_botoes.grid(row=5, column=0, sticky="sew", padx=20, pady=20)
        frame_botoes.grid_columnconfigure(0, weight=1)
        frame_botoes.grid_columnconfigure(1, weight=1)

        self.btn_salvar = ctk.CTkButton(frame_botoes, text="Salvar", fg_color=COLOR_ACCENT, command=self.salvar)
        self.btn_salvar.grid(row=0, column=1, sticky="ew", padx=(5,0))

        self.btn_cancelar = ctk.CTkButton(frame_botoes, text="Cancelar", fg_color=COLOR_DANGER, command=self.destroy)
        self.btn_cancelar.grid(row=0, column=0, sticky="ew", padx=(0,5))

        # Preencher dados se estiver em modo de edição
        if self.funcionario_data:
            self._preencher_dados()

    def _preencher_dados(self):
        self.entry_nome.insert(0, self.funcionario_data.get("nome", ""))
        self.entry_salario.insert(0, self.funcionario_data.get("salario", ""))

        try:
            data_adm = datetime.datetime.strptime(self.funcionario_data.get("admissao"), '%d/%m/%Y').date()
            self.entry_admissao.set_date(data_adm)
        except (ValueError, TypeError):
            pass

        # Mostra os campos extras se houver dados neles
        if any(self.funcionario_data.get(k) for k in ["email", "celular", "cpf", "carteira_trabalho"]):
            self.toggle_dados_extras()

        self.entry_email.insert(0, self.funcionario_data.get("email", ""))
        self.entry_celular.insert(0, self.funcionario_data.get("celular", ""))
        self.entry_cpf.insert(0, self.funcionario_data.get("cpf", ""))
        self.entry_ctps.insert(0, self.funcionario_data.get("carteira_trabalho", ""))

        if self.funcionario_data.get("fotos"):
            primeira_foto = self.funcionario_data["fotos"][0]
            path_foto = os.path.join(self.master.pasta_funcionarios, primeira_foto)
            if os.path.exists(path_foto):
                self._carregar_imagem(path_foto)

    def drop_image(self, event):
        path = event.data.replace('{', '').replace('}', '')
        if path.lower().endswith(('.png', '.jpg', '.jpeg')):
            self.photo_path = path
            self._carregar_imagem(path)
        else:
            messagebox.showwarning("Formato Inválido", "Por favor, arraste um arquivo de imagem (PNG, JPG).", parent=self)

    def select_image(self):
        path = filedialog.askopenfilename(filetypes=[("Imagens", "*.jpg *.jpeg *.png")])
        if path:
            self.photo_path = path
            self._carregar_imagem(path)

    def _carregar_imagem(self, path):
        try:
            pil_image = Image.open(path)
            h, w = pil_image.height, pil_image.width
            ratio = min(180/w, 180/h)

            ctk_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(int(w*ratio), int(h*ratio)))
            self.lbl_foto_preview.configure(image=ctk_image, text="")
        except Exception as e:
            messagebox.showerror("Erro ao carregar imagem", str(e), parent=self)

    def toggle_dados_extras(self):
        self.dados_extras_visiveis = not self.dados_extras_visiveis
        if self.dados_extras_visiveis:
            self.frame_extra.grid(row=3, column=0, sticky="ew", padx=20, pady=10)
            self.btn_dados_extras.configure(text="Menos Dados ▲")
            self.geometry("500x800")
        else:
            self.frame_extra.grid_forget()
            self.btn_dados_extras.configure(text="Mais Dados ▼")
            self.geometry("500x650")

    def salvar(self):
        nome = self.entry_nome.get().strip()
        if not nome:
            messagebox.showerror("Campo Obrigatório", "O nome do funcionário é obrigatório.", parent=self)
            return

        salario = self.entry_salario.get().replace(',', '.')
        try:
            float(salario if salario else 0)
        except ValueError:
            messagebox.showerror("Valor Inválido", "O salário deve ser um número válido.", parent=self)
            return

        # Lógica para salvar/atualizar
        try:
            novo_nome_foto = None
            if self.photo_path:
                nome_limpo = re.sub(r'[^a-zA-Z0-9]', '', nome)
                novo_nome_foto = f"{nome_limpo}_{int(time.time())}.jpg"
                shutil.copy(self.photo_path, os.path.join(self.master.pasta_funcionarios, novo_nome_foto))

            if self.funcionario_data is None: # Novo Cadastro
                if not self.photo_path:
                    messagebox.showerror("Campo Obrigatório", "A foto é obrigatória para um novo cadastro.", parent=self)
                    return

                novo_funcionario = {
                    "id": int(time.time() * 1000),
                    "nome": nome,
                    "salario": salario,
                    "admissao": self.entry_admissao.get_date().strftime('%d/%m/%Y'),
                    "email": self.entry_email.get(),
                    "celular": self.entry_celular.get(),
                    "cpf": self.entry_cpf.get(),
                    "carteira_trabalho": self.entry_ctps.get(),
                    "status": "ativo",
                    "fotos": [novo_nome_foto]
                }
                self.master.dados_funcionarios.append(novo_funcionario)
            else: # Edição
                for func in self.master.dados_funcionarios:
                    if func['id'] == self.funcionario_data['id']:
                        func['nome'] = nome
                        func['salario'] = salario
                        func['admissao'] = self.entry_admissao.get_date().strftime('%d/%m/%Y')
                        func['email'] = self.entry_email.get()
                        func['celular'] = self.entry_celular.get()
                        func['cpf'] = self.entry_cpf.get()
                        func['carteira_trabalho'] = self.entry_ctps.get()
                        if novo_nome_foto:
                            func['fotos'].append(novo_nome_foto)
                        break

            self.master.salvar_dados_funcionarios()
            self.master.carregar_lista_funcionarios()
            self.destroy()

        except Exception as e:
            messagebox.showerror("Erro ao Salvar", str(e), parent=self)


# --- CLASSE PRINCIPAL COM SUPORTE A DND ---
class AppPonto(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()

        # Inicializa Drag & Drop
        self.TkdndVersion = TkinterDnD._require(self)

        self.title("Sistema Ponto Neural v16.2 - Estável")
        self.geometry("1200x850")

        self.current_theme = "Noturno"
        self.configure(fg_color=self._get_color("COLOR_BG"))

        # Tenta carregar o ícone (se existir)
        try:
            self.iconbitmap("icone.ico")
        except:
            pass

        self.config = self.load_config()
        self.parar_execucao = False
        self.queue = queue.Queue()
        self.after(100, self.verificar_fila)

        self.temp_dir = None
        self.pasta_funcionarios = "funcionarios"
        if not os.path.exists(self.pasta_funcionarios): os.makedirs(self.pasta_funcionarios)

        self.dados_funcionarios = []
        self.carregar_dados_funcionarios()

        self.historico_relatorios = []
        self.dados_consolidados = []
        self.caminho_zip = ""
        self.conhecidos_nom = []

        # --- HEADER ---
        self.header = ctk.CTkFrame(self, height=60, fg_color="transparent")
        self.header.pack(fill="x", padx=20, pady=(10, 0))
        self.header.grid_columnconfigure(0, weight=1)

        self.lbl_title = ctk.CTkLabel(self.header, text="Sistema Ponto Neural", font=("Roboto Medium", 24))
        self.lbl_title.grid(row=0, column=0, sticky="w", padx=10)

        frame_switch = ctk.CTkFrame(self.header, fg_color="transparent")
        frame_switch.grid(row=0, column=1, sticky="e")

        self.theme_switch = ctk.CTkSwitch(frame_switch, text="Tema Claro", command=self._toggle_theme)
        self.theme_switch.pack(side="left", padx=10)

        # --- LAYOUT PRINCIPAL ---
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.main_container.grid_columnconfigure(1, weight=1)
        self.main_container.grid_rowconfigure(0, weight=1)

        # --- BARRA DE NAVEGAÇÃO ---
        self.nav_frame = ctk.CTkFrame(self.main_container, width=200, corner_radius=15)
        self.nav_frame.grid(row=0, column=0, sticky="nsw", padx=(0, 10))

        self.nav_buttons = {}
        self.frames = {}
        self.icons = {}

        # --- CONTEÚDO PRINCIPAL ---
        self.content_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew")

        for F, name, icon_name in zip(
            (AbaProcessamento, AbaFuncionarios, AbaRelatorios),
            ("Processamento", "Funcionários", "Relatórios"),
            ("rocket", "users", "clipboard-list")
        ):
            frame = F(self.content_frame, self)
            self.frames[name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

            button = ctk.CTkButton(
                self.nav_frame,
                text=f" {name}",
                anchor="w",
                height=50,
                corner_radius=8,
                command=lambda n=name: self.select_frame_by_name(n),
                compound="left"
            )
            button.pack(fill="x", padx=10, pady=5)
            self.nav_buttons[name] = button

        self.select_frame_by_name("Processamento")
        self._apply_theme()

# ===================================================================
# --- CLASSES DAS ABAS ---
# ===================================================================

class AbaProcessamento(ctk.CTkFrame):
    def __init__(self, master, app_ref):
        super().__init__(master, fg_color="transparent")
        self.app = app_ref
        self.pack(fill="both", expand=True)

        self.grid_columnconfigure(0, weight=4)
        self.grid_columnconfigure(1, weight=5)
        self.grid_rowconfigure(1, weight=1)

        # ESQUERDA
        self.frame_left = ctk.CTkFrame(self, corner_radius=15, border_width=1)
        self.frame_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)

        # Datas
        self.app.create_section_label(self.frame_left, "1. PERÍODO")
        self.frame_datas = ctk.CTkFrame(self.frame_left, fg_color="transparent")
        self.frame_datas.pack(fill="x", padx=20)

        style_cal = {'background': '#1e293b', 'foreground': 'white', 'borderwidth': 0, 'date_pattern': 'dd/mm/yyyy'}

        self.lbl_de = ctk.CTkLabel(self.frame_datas, text="De:")
        self.lbl_de.grid(row=0, column=0, sticky="w")
        self.app.cal_inicio = DateEntry(self.frame_datas, width=12, font=("Arial", 11), **style_cal, locale='pt_BR')
        self.app.cal_inicio.set_date(datetime.date.today().replace(day=1))
        self.app.cal_inicio.grid(row=1, column=0, padx=5, pady=5)

        self.lbl_ate = ctk.CTkLabel(self.frame_datas, text="Até:")
        self.lbl_ate.grid(row=0, column=1, sticky="w")
        self.app.cal_fim = DateEntry(self.frame_datas, width=12, font=("Arial", 11), **style_cal, locale='pt_BR')
        self.app.cal_fim.set_date(datetime.date.today())
        self.app.cal_fim.grid(row=1, column=1, padx=5, pady=5)

        # Slider
        self.separator1 = ctk.CTkFrame(self.frame_left, height=1)
        self.separator1.pack(fill="x", padx=20, pady=20)
        self.app.create_section_label(self.frame_left, "2. SENSIBILIDADE")

        self.frame_slider = ctk.CTkFrame(self.frame_left, fg_color="transparent")
        self.frame_slider.pack(fill="x", padx=20, pady=5)
        self.frame_slider.grid_columnconfigure(0, weight=1)

        self.app.slider = ctk.CTkSlider(self.frame_slider, from_=0.35, to=0.60, number_of_steps=25, command=self.app.update_slider_label)
        self.app.slider.set(self.app.config.get('tolerancia', 0.45))
        self.app.slider.grid(row=0, column=0, sticky="ew", padx=(5, 15))

        self.app.lbl_slider_value = ctk.CTkLabel(self.frame_slider, text=f"{self.app.slider.get():.2f}", font=("Consolas", 12), width=40)
        self.app.lbl_slider_value.grid(row=0, column=1)

        # Arquivo (Drag & Drop)
        self.separator2 = ctk.CTkFrame(self.frame_left, height=1)
        self.separator2.pack(fill="x", padx=20, pady=20)
        self.app.create_section_label(self.frame_left, "3. ARQUIVO ZIP (Arraste aqui)")

        self.app.frame_file = ctk.CTkFrame(self.frame_left, corner_radius=8, border_width=2)
        self.app.frame_file.pack(fill="x", padx=20, pady=5, ipady=10)

        self.app.frame_file.drop_target_register(DND_FILES)
        self.app.frame_file.dnd_bind('<<Drop>>', self.app.drop_file)

        self.app.lbl_arquivo = ctk.CTkLabel(self.app.frame_file, text="Arraste o ZIP ou clique na pasta", font=("Consolas", 11))
        self.app.lbl_arquivo.pack(side="left", padx=10, pady=10)

        self.btn_select = ctk.CTkButton(self.app.frame_file, text="📂", width=40, command=self.app.selecionar_zip)
        self.btn_select.pack(side="right", padx=10)

        # Ações
        self.frame_actions = ctk.CTkFrame(self.frame_left, fg_color="transparent")
        self.frame_actions.pack(side="bottom", fill="x", padx=20, pady=30)
        self.app.btn_iniciar = ctk.CTkButton(self.frame_actions, text="INICIAR", height=50, font=("Arial", 14, "bold"), command=self.app.iniciar_thread)
        self.app.btn_iniciar.pack(fill="x", pady=(0, 10))
        self.app.btn_pdf = ctk.CTkButton(self.frame_actions, text="GERAR RELATÓRIO PDF", height=40, fg_color="transparent", border_width=1, state="disabled", command=self.app.gerar_pdf_acao_wrapper)
        self.app.btn_pdf.pack(fill="x")

        # DIREITA
        self.frame_right = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_right.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)
        self.frame_right.grid_rowconfigure(1, weight=1)

        self.card_status = ctk.CTkFrame(self.frame_right, corner_radius=15, border_width=1)
        self.card_status.grid(row=0, column=0, sticky="ew")

        self.lbl_status_title = ctk.CTkLabel(self.card_status, text="STATUS", font=("Arial", 11, "bold"))
        self.lbl_status_title.pack(anchor="w", padx=20, pady=(15, 5))
        self.app.progress_bar = ctk.CTkProgressBar(self.card_status, height=10)
        self.app.progress_bar.set(0)
        self.app.progress_bar.pack(fill="x", padx=20, pady=5)
        self.app.lbl_status_txt = ctk.CTkLabel(self.card_status, text="Pronto.")
        self.app.lbl_status_txt.pack(side="left", padx=20, pady=(0, 15))
        self.app.lbl_estimativa = ctk.CTkLabel(self.card_status, text="--:--")
        self.app.lbl_estimativa.pack(side="right", padx=20, pady=(0, 15))

        self.app.txt_log = ctk.CTkTextbox(self.frame_right, font=("Consolas", 11), corner_radius=10)
        self.app.txt_log.grid(row=1, column=0, sticky="nsew", pady=(15,0))
        self.app.txt_log.configure(state="disabled")

        self.app.btn_parar = ctk.CTkButton(self.app.header, text="PARAR", width=80, state="disabled", command=self.app.solicitar_parada)

class AbaFuncionarios(ctk.CTkFrame):
    def __init__(self, master, app_ref):
        super().__init__(master, fg_color="transparent")
        self.app = app_ref
        self.pack(fill="both", expand=True)

        # --- Toolbar Superior ---
        self.toolbar = ctk.CTkFrame(self, fg_color="transparent", height=50)
        self.toolbar.pack(fill="x", padx=10, pady=10)

        self.btn_add = ctk.CTkButton(self.toolbar, text="+ Adicionar Novo", command=self.app.add_funcionario)
        self.btn_add.pack(side="left")
        self.btn_reload = ctk.CTkButton(self.toolbar, text="🔄 Recarregar", width=100, command=self.app.carregar_lista_funcionarios)
        self.btn_reload.pack(side="left", padx=10)

        # --- Frame de Filtros ---
        self.filter_frame = ctk.CTkFrame(self, border_width=1)
        self.filter_frame.pack(fill="x", padx=10, pady=(0, 10), ipady=10)

        self.lbl_filter_name = ctk.CTkLabel(self.filter_frame, text="Filtrar por nome:")
        self.lbl_filter_name.pack(side="left", padx=(15, 5))
        self.app.filtro_nome = ctk.CTkEntry(self.filter_frame, placeholder_text="Digite um nome...")
        self.app.filtro_nome.pack(side="left", padx=5, fill="x", expand=True)
        self.app.filtro_nome.bind("<KeyRelease>", lambda e: self.app.carregar_lista_funcionarios())

        self.lbl_sort_by = ctk.CTkLabel(self.filter_frame, text="Ordenar por:")
        self.lbl_sort_by.pack(side="left", padx=(15, 5))
        self.app.filtro_ordem_campo = ctk.CTkComboBox(self.filter_frame, values=["Nome", "Data de Admissão", "Salário"], command=lambda e: self.app.carregar_lista_funcionarios())
        self.app.filtro_ordem_campo.set("Nome")
        self.app.filtro_ordem_campo.pack(side="left", padx=5)

        self.app.filtro_ordem_dir = ctk.CTkComboBox(self.filter_frame, values=["Crescente", "Decrescente"], width=120, command=lambda e: self.app.carregar_lista_funcionarios())
        self.app.filtro_ordem_dir.set("Crescente")
        self.app.filtro_ordem_dir.pack(side="left", padx=5)

        self.app.scroll_func = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.app.scroll_func.pack(fill="both", expand=True, padx=10, pady=5)

        self.app.carregar_lista_funcionarios()

class AbaRelatorios(ctk.CTkFrame):
    def __init__(self, master, app_ref):
        super().__init__(master, fg_color="transparent")
        self.app = app_ref
        self.pack(fill="both", expand=True)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Frame dos Controles
        self.frame_controles = ctk.CTkFrame(self, corner_radius=15, border_width=1)
        self.frame_controles.grid(row=0, column=0, sticky="new", padx=10, pady=10)

        # Seção de Filtros
        self.app.create_section_label(self.frame_controles, "FILTROS")

        # Frame para os filtros
        self.frame_filtros = ctk.CTkFrame(self.frame_controles, fg_color="transparent")
        self.frame_filtros.pack(fill="x", padx=20, pady=10)

        # Filtro de Data
        self.lbl_periodo = ctk.CTkLabel(self.frame_filtros, text="Período:")
        self.lbl_periodo.grid(row=0, column=0, sticky="w", padx=(0, 10))

        style_cal = {'background': '#1e293b', 'foreground': 'white', 'borderwidth': 0, 'date_pattern': 'dd/mm/yyyy'}

        self.app.cal_relatorio_inicio = DateEntry(self.frame_filtros, width=12, font=("Arial", 11), **style_cal, locale='pt_BR')
        self.app.cal_relatorio_inicio.set_date(datetime.date.today().replace(day=1))
        self.app.cal_relatorio_inicio.grid(row=0, column=1, padx=5)

        self.lbl_ate = ctk.CTkLabel(self.frame_filtros, text="até")
        self.lbl_ate.grid(row=0, column=2, padx=5)

        self.app.cal_relatorio_fim = DateEntry(self.frame_filtros, width=12, font=("Arial", 11), **style_cal, locale='pt_BR')
        self.app.cal_relatorio_fim.set_date(datetime.date.today())
        self.app.cal_relatorio_fim.grid(row=0, column=3, padx=5)

        # Filtro de Funcionário
        self.lbl_funcionario = ctk.CTkLabel(self.frame_filtros, text="Funcionário:")
        self.lbl_funcionario.grid(row=0, column=4, sticky="w", padx=(20, 10))

        self.app.funcionarios_selecionados = ["Todos"]
        self.app.combo_funcionarios = ctk.CTkComboBox(self.frame_filtros, values=["Todos", "Personalizar..."], width=200, command=self.app.on_selecionar_funcionario)
        self.app.combo_funcionarios.set("Todos")
        self.app.combo_funcionarios.grid(row=0, column=5, padx=5)

        self.app.frames["Relatorios"].bind("<<Map>>", self.app.atualizar_lista_funcionarios_relatorio)

        # Botão para Gerar Relatório
        self.btn_gerar_relatorio = ctk.CTkButton(self.frame_controles, text="Gerar Relatório", command=self.app.gerar_relatorio)
        self.btn_gerar_relatorio.pack(pady=(10, 20))

        # Área de Preview do Relatório
        self.app.txt_relatorio = ctk.CTkTextbox(self, font=("Consolas", 11), corner_radius=10)
        self.app.txt_relatorio.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.app.txt_relatorio.configure(state="disabled")

        # Frame para os botões de exportação
        self.frame_export = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_export.grid(row=2, column=0, sticky="e", padx=10, pady=10)

        self.app.btn_export_csv = ctk.CTkButton(self.frame_export, text="Exportar CSV", state="disabled", command=self.app.exportar_csv)
        self.app.btn_export_csv.pack(side="left", padx=5)

        self.app.btn_export_excel = ctk.CTkButton(self.frame_export, text="Exportar Excel", state="disabled", command=self.app.exportar_excel)
        self.app.btn_export_excel.pack(side="left", padx=5)

        self.app.btn_export_pdf = ctk.CTkButton(self.frame_export, text="Exportar PDF", state="disabled", command=self.app.exportar_pdf)
        self.app.btn_export_pdf.pack(side="left", padx=5)

        # --- Histórico de Relatórios ---
        self.lbl_hist_title = ctk.CTkLabel(self, text="HISTÓRICO DE RELATÓRIOS GERADOS", font=("Arial", 11, "bold"))
        self.lbl_hist_title.grid(row=3, column=0, sticky="w", padx=10, pady=(10,0))
        self.app.scroll_historico = ctk.CTkScrollableFrame(self, height=150)
        self.app.scroll_historico.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.app.lbl_historico_vazio = ctk.CTkLabel(self.app.scroll_historico, text="Nenhum relatório gerado nesta sessão.")
        self.app.lbl_historico_vazio.pack(pady=20)

    def on_selecionar_funcionario(self, choice):
        if choice == "Personalizar...":
            self.open_selecao_funcionarios_window()
        else:
            self.funcionarios_selecionados = [choice]

    def open_selecao_funcionarios_window(self):
        # Garante que a janela não seja aberta múltiplas vezes
        if hasattr(self, 'toplevel_selecao') and self.toplevel_selecao.winfo_exists():
            self.toplevel_selecao.focus()
            return
        self.toplevel_selecao = ToplevelSelecaoFuncionarios(self)
        self.wait_window(self.toplevel_selecao) # Pausa a execução até a janela ser fechada


    def atualizar_lista_funcionarios_relatorio(self, event=None):
        nomes_ativos = sorted([
            f['nome'] for f in self.dados_funcionarios if f.get('status', 'ativo') == 'ativo'
        ])
        self.combo_funcionarios.configure(values=["Todos"] + nomes_ativos + ["Personalizar..."])

    def gerar_relatorio(self):
        self.txt_relatorio.configure(state="normal")
        self.txt_relatorio.delete("1.0", "end")

        if not self.dados_consolidados:
            self.txt_relatorio.insert("end", "Nenhum dado processado para gerar relatório.\nPor favor, processe um arquivo ZIP primeiro na aba 'Processamento'.")
            self.txt_relatorio.configure(state="disabled")
            return

        self.txt_relatorio.insert("end", "Gerando relatório...\n")

        data_inicio = self.cal_relatorio_inicio.get_date()
        data_fim = self.cal_relatorio_fim.get_date()

        # Filtrar dados com base na nova lista de seleção
        dados_filtrados = []
        for registro in self.dados_consolidados:
            try:
                data_registro = datetime.datetime.strptime(registro['data'], "%d/%m/%Y").date()
                if data_inicio <= data_registro <= data_fim:
                    if "Todos" in self.funcionarios_selecionados or registro['nome'] in self.funcionarios_selecionados:
                        dados_filtrados.append(registro)
            except (ValueError, TypeError):
                continue

        if not dados_filtrados:
            self.txt_relatorio.insert("end", "Nenhum registro encontrado para os filtros selecionados.\n")
            self.txt_relatorio.configure(state="disabled")
            return

        # Agregar dados
        registros_por_dia = defaultdict(lambda: defaultdict(list))
        for registro in dados_filtrados:
            registros_por_dia[registro['nome']][registro['data']].append(registro['hora'])

        self.dados_relatorio = []
        for nome, dias in registros_por_dia.items():
            for data, horas in dias.items():
                horas.sort()
                entrada = horas[0]
                saida = horas[-1] if len(horas) > 1 else "--:--"
                self.dados_relatorio.append({
                    "Nome": nome,
                    "Data": data,
                    "Entrada": entrada,
                    "Saída": saida
                })

        # Ordenar relatório por data
        self.dados_relatorio.sort(key=lambda x: datetime.datetime.strptime(x['Data'], "%d/%m/%Y"))

        # Exibir no preview
        self.txt_relatorio.delete("1.0", "end")
        header = f"{'Nome':<20} {'Data':<12} {'Entrada':<10} {'Saída':<10}\n"
        self.txt_relatorio.insert("end", header)
        self.txt_relatorio.insert("end", "-" * 60 + "\n")

        for linha in self.dados_relatorio:
            self.txt_relatorio.insert("end", f"{linha['Nome']:<20} {linha['Data']:<12} {linha['Entrada']:<10} {linha['Saída']:<10}\n")

        self.txt_relatorio.configure(state="disabled")
        # Habilitar botões de exportação
        self.btn_export_csv.configure(state="normal")
        self.btn_export_excel.configure(state="normal")
        self.btn_export_pdf.configure(state="normal")

        self.adicionar_relatorio_ao_historico()

    def adicionar_relatorio_ao_historico(self):
        nome_relatorio = self.combo_funcionarios.get()
        if len(self.funcionarios_selecionados) > 1:
            nome_relatorio = f"{len(self.funcionarios_selecionados)} funcionários"

        novo_historico = {
            "nome": nome_relatorio,
            "timestamp": datetime.datetime.now(),
            "dados": self.dados_relatorio.copy()
        }
        self.historico_relatorios.insert(0, novo_historico) # Adiciona no início
        self.atualizar_visualizacao_historico()

    def atualizar_visualizacao_historico(self):
        for widget in self.scroll_historico.winfo_children():
            widget.destroy()

        if not self.historico_relatorios:
            self.lbl_historico_vazio = ctk.CTkLabel(self.scroll_historico, text="Nenhum relatório gerado nesta sessão.", text_color=COLOR_TEXT_DIM)
            self.lbl_historico_vazio.pack(pady=20)
            return

        for item in self.historico_relatorios:
            card = ctk.CTkButton(
                self.scroll_historico,
                text=f"{item['nome']} - {item['timestamp'].strftime('%H:%M:%S')}",
                fg_color="#1e293b",
                hover_color="#3b82f6",
                anchor="w",
                command=lambda i=item: self.recarregar_relatorio_do_historico(i)
            )
            card.pack(fill="x", padx=5, pady=2)

    def recarregar_relatorio_do_historico(self, item_historico):
        self.dados_relatorio = item_historico['dados']

        self.txt_relatorio.configure(state="normal")
        self.txt_relatorio.delete("1.0", "end")

        header = f"{'Nome':<20} {'Data':<12} {'Entrada':<10} {'Saída':<10}\n"
        self.txt_relatorio.insert("end", header)
        self.txt_relatorio.insert("end", "-" * 60 + "\n")

        for linha in self.dados_relatorio:
            self.txt_relatorio.insert("end", f"{linha['Nome']:<20} {linha['Data']:<12} {linha['Entrada']:<10} {linha['Saída']:<10}\n")

        self.txt_relatorio.configure(state="disabled")
        self.btn_export_csv.configure(state="normal")
        self.btn_export_excel.configure(state="normal")
        self.btn_export_pdf.configure(state="normal")


    def exportar_pdf(self):
        if not self.dados_relatorio:
            messagebox.showwarning("Aviso", "Gere um relatório primeiro.")
            return

        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not filepath:
            return

        try:
            self.gerar_pdf(filepath, self.dados_relatorio)
            messagebox.showinfo("Sucesso", "Relatório exportado para PDF com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar para PDF: {e}")

    def exportar_csv(self):
        if not self.dados_relatorio:
            return

        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not filepath:
            return

        try:
            with open(filepath, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=self.dados_relatorio[0].keys())
                writer.writeheader()
                writer.writerows(self.dados_relatorio)
            messagebox.showinfo("Sucesso", "Relatório exportado para CSV com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar para CSV: {e}")

    def exportar_excel(self):
        if not self.dados_relatorio:
            return

        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not filepath:
            return

        try:
            workbook = Workbook()
            sheet = workbook.active

            headers = list(self.dados_relatorio[0].keys())
            sheet.append(headers)

            for row in self.dados_relatorio:
                sheet.append(list(row.values()))

            workbook.save(filepath)
            messagebox.showinfo("Sucesso", "Relatório exportado para Excel com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar para Excel: {e}")

    def carregar_lista_funcionarios(self, event=None):
        for widget in self.scroll_func.winfo_children():
            widget.destroy()

        # 1. Filtrar
        texto_filtro = self.filtro_nome.get().lower()
        funcionarios_filtrados = [
            f for f in self.dados_funcionarios
            if f.get('status', 'ativo') == 'ativo' and texto_filtro in f['nome'].lower()
        ]

        # 2. Ordenar
        campo_ordem = self.filtro_ordem_campo.get()
        direcao_ordem = self.filtro_ordem_dir.get() == "Decrescente"

        def get_sort_key(f):
            if campo_ordem == "Data de Admissão":
                try:
                    return datetime.datetime.strptime(f.get("admissao", "01/01/1900"), '%d/%m/%Y')
                except ValueError:
                    return datetime.datetime.min
            elif campo_ordem == "Salário":
                try:
                    return float(str(f.get("salario", "0")).replace(',', '.'))
                except ValueError:
                    return 0.0
            return f['nome'].lower()

        funcionarios_filtrados.sort(key=get_sort_key, reverse=direcao_ordem)

        # 3. Exibir
        if not funcionarios_filtrados:
            ctk.CTkLabel(self.scroll_func, text="Nenhum funcionário encontrado com os filtros atuais.").pack(pady=20)
            return

        for funcionario_data in funcionarios_filtrados:
            self.criar_grupo_funcionario(funcionario_data)

    def criar_grupo_funcionario(self, funcionario_data):
        nome = funcionario_data['nome']
        fotos = funcionario_data['fotos']

        card = ctk.CTkFrame(self.scroll_func, fg_color=COLOR_CARD, corner_radius=10, border_color=COLOR_BORDER, border_width=1)
        card.pack(fill="x", padx=5, pady=5)

        header = ctk.CTkFrame(card, fg_color="transparent", height=50)
        header.pack(fill="x", padx=10, pady=5)

        try:
            path = os.path.join(self.pasta_funcionarios, fotos[0]) if fotos else ""
            if path and os.path.exists(path):
                pil = Image.open(path)
                pil.thumbnail((40, 40))
                tk_img = ctk.CTkImage(light_image=pil, dark_image=pil, size=(40, 40))
                lbl_img = ctk.CTkLabel(header, image=tk_img, text="")
                lbl_img.pack(side="left")
            else:
                ctk.CTkLabel(header, text="👤", font=("Arial", 20)).pack(side="left", padx=10)
        except:
            ctk.CTkLabel(header, text="👤", font=("Arial", 20)).pack(side="left", padx=10)

        ctk.CTkLabel(header, text=nome, font=("Arial", 14, "bold"), text_color="white").pack(side="left", padx=15)
        ctk.CTkLabel(header, text=f"{len(fotos)} variações", font=("Arial", 11), text_color=COLOR_TEXT_DIM).pack(side="left", padx=5)

        btn_expand = ctk.CTkButton(header, text="▼", width=30, fg_color="transparent", text_color=COLOR_INFO, hover_color="#1e293b")
        btn_expand.pack(side="right")

        body = ctk.CTkFrame(card, fg_color="#0b1120")

        # Ações do funcionário
        frame_acoes = ctk.CTkFrame(body, fg_color="transparent")
        frame_acoes.pack(fill="x", padx=10, pady=10)

        ctk.CTkButton(frame_acoes, text="Editar", width=80, command=lambda f=funcionario_data: self.open_funcionario_window(f)).pack(side="left", padx=5)
        ctk.CTkButton(frame_acoes, text="Adicionar Variação", width=120, command=lambda f=funcionario_data: self.open_adicionar_variacao_window(f)).pack(side="left", padx=5)
        ctk.CTkButton(frame_acoes, text="Histórico", width=80, state="disabled").pack(side="left", padx=5)
        btn_excluir = ctk.CTkButton(frame_acoes, text="Excluir", fg_color=COLOR_DANGER, width=80, command=lambda f=funcionario_data: self.excluir_funcionario(f))
        btn_excluir.pack(side="right", padx=5)

        # Lista de fotos
        frame_fotos = ctk.CTkFrame(body, fg_color="transparent")
        frame_fotos.pack(fill="x", padx=10, pady=5)

        for f in fotos:
            row = ctk.CTkFrame(frame_fotos, fg_color="transparent", height=30)
            row.pack(fill="x", padx=10, pady=2)
            ctk.CTkLabel(row, text=f"📄 {f}", font=("Consolas", 11), text_color="gray").pack(side="left")
            btn_del_foto = ctk.CTkButton(row, text="🗑️", width=30, height=20, fg_color="transparent", hover_color=COLOR_DANGER,
                                    command=lambda file=f: self.delete_foto_funcionario(funcionario_data['id'], file))
            btn_del_foto.pack(side="right")

        def toggle():
            if body.winfo_ismapped():
                body.pack_forget(); btn_expand.configure(text="▼")
            else:
                body.pack(fill="x", padx=10, pady=(0, 10)); btn_expand.configure(text="▲")

        btn_expand.configure(command=toggle)

    def add_funcionario(self):
        self.open_funcionario_window()

    def open_funcionario_window(self, funcionario_data=None):
        if hasattr(self, 'toplevel_window') and self.toplevel_window.winfo_exists():
            self.toplevel_window.focus()
            return
        self.toplevel_window = ToplevelWindowFuncionario(self, funcionario_data)

    def open_adicionar_variacao_window(self, funcionario):
        path = filedialog.askopenfilename(
            title=f"Selecione uma nova foto para {funcionario['nome']}",
            filetypes=[("Imagens", "*.jpg *.jpeg *.png")]
        )
        if not path:
            return

        try:
            nome_limpo = re.sub(r'[^a-zA-Z0-9]', '', funcionario['nome'])
            novo_nome_foto = f"{nome_limpo}_{int(time.time())}.jpg"

            # Copia a nova foto
            shutil.copy(path, os.path.join(self.pasta_funcionarios, novo_nome_foto))

            # Atualiza a lista de fotos do funcionário
            for func in self.dados_funcionarios:
                if func['id'] == funcionario['id']:
                    func['fotos'].append(novo_nome_foto)
                    break

            self.salvar_dados_funcionarios()
            self.carregar_lista_funcionarios()
            messagebox.showinfo("Sucesso", "Nova variação de foto adicionada com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao adicionar a variação: {e}")

    def excluir_funcionario(self, funcionario):
        if not messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o funcionário '{funcionario['nome']}'?\n\nIsso irá apenas desativá-lo da interface, mantendo seus dados no histórico. A ação pode ser revertida manualmente no arquivo 'funcionarios.json'."):
            return

        try:
            # Encontra o funcionário e muda o status
            for func in self.dados_funcionarios:
                if func['id'] == funcionario['id']:
                    func['status'] = 'inativo'
                    break

            self.salvar_dados_funcionarios()
            self.carregar_lista_funcionarios()
            messagebox.showinfo("Sucesso", f"Funcionário '{funcionario['nome']}' desativado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao excluir o funcionário: {e}")

    def delete_foto_funcionario(self, funcionario_id, filename):
        if not messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir a foto '{filename}'?\nEsta ação não pode ser desfeita."):
            return

        try:
            # Encontra o funcionário e remove a foto da lista
            for func in self.dados_funcionarios:
                if func['id'] == funcionario_id:
                    if filename in func['fotos']:
                        func['fotos'].remove(filename)
                        break

            # Salva a alteração no JSON
            self.salvar_dados_funcionarios()

            # Remove o arquivo físico
            filepath = os.path.join(self.pasta_funcionarios, filename)
            if os.path.exists(filepath):
                os.remove(filepath)

            # Recarrega a interface
            self.carregar_lista_funcionarios()

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao excluir a foto: {e}")

    # --- DRAG & DROP HANDLER ---
    def drop_file(self, event):
        path = event.data.replace('{', '').replace('}', '')
        if path.lower().endswith('.zip'):
            self.caminho_zip = path
            self.lbl_arquivo.configure(text=os.path.basename(path), text_color=COLOR_ACCENT)
            self.frame_file.configure(border_color=COLOR_ACCENT)
            self.config['last_dir'] = os.path.dirname(path)
            self.save_config()
        else: messagebox.showwarning("Arquivo Inválido", "Arraste um ZIP!")

    # --- PERSISTÊNCIA ---
    def load_config(self):
        if os.path.exists(ARQUIVO_CONFIG):
            try:
                with open(ARQUIVO_CONFIG, 'r') as f: return json.load(f)
            except: pass
        return {'tolerancia': 0.45, 'last_dir': '/'}

    def save_config(self):
        try:
            self.config['tolerancia'] = self.slider.get()
            with open(ARQUIVO_CONFIG, 'w') as f: json.dump(self.config, f)
        except: pass

    def carregar_dados_funcionarios(self):
        if not os.path.exists(ARQUIVO_FUNCIONARIOS):
            self.migrar_dados_antigos()

        try:
            with open(ARQUIVO_FUNCIONARIOS, 'r', encoding='utf-8') as f:
                self.dados_funcionarios = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.dados_funcionarios = []

    def salvar_dados_funcionarios(self):
        try:
            with open(ARQUIVO_FUNCIONARIOS, 'w', encoding='utf-8') as f:
                json.dump(self.dados_funcionarios, f, indent=4)
        except Exception as e:
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar os dados dos funcionários: {e}")

    def migrar_dados_antigos(self):
        log_debug("Arquivo 'funcionarios.json' não encontrado. Tentando migrar do sistema antigo...")
        funcionarios_migrados = []
        if not os.path.exists(self.pasta_funcionarios):
            self.dados_funcionarios = []
            self.salvar_dados_funcionarios()
            return

        arquivos = [f for f in os.listdir(self.pasta_funcionarios) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        funcionarios_dict = defaultdict(list)

        for arq in arquivos:
            try:
                nome = os.path.splitext(arq)[0].split('_')[0].capitalize()
                funcionarios_dict[nome].append(arq)
            except:
                continue

        for nome, fotos in funcionarios_dict.items():
            funcionarios_migrados.append({
                "id": int(time.time() * 1000) + len(funcionarios_migrados),
                "nome": nome,
                "salario": "",
                "admissao": "",
                "email": "",
                "celular": "",
                "cpf": "",
                "carteira_trabalho": "",
                "status": "ativo",
                "fotos": fotos
            })

        self.dados_funcionarios = funcionarios_migrados
        self.salvar_dados_funcionarios()
        log_debug(f"Migração concluída. {len(self.dados_funcionarios)} funcionários salvos em 'funcionarios.json'.")

    # --- UI HELPERS ---
    def update_slider_label(self, value):
        self.lbl_slider_value.configure(text=f"{value:.2f}")

    def create_section_label(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=("Arial", 11, "bold"), text_color=COLOR_TEXT_DIM).pack(anchor="w", padx=20, pady=(0, 5))

    def log_tela(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"> {msg}\n"); self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def selecionar_zip(self):
        init_dir = self.config.get('last_dir', '/')
        f = filedialog.askopenfilename(initialdir=init_dir, filetypes=[("Arquivo ZIP", "*.zip")])
        if f:
            self.caminho_zip = f
            self.lbl_arquivo.configure(text=os.path.basename(f), text_color="white")
            self.config['last_dir'] = os.path.dirname(f)
            self.save_config()

    def solicitar_parada(self):
        if messagebox.askyesno("Parar", "Interromper?"): self.parar_execucao = True

    def restaurar_botoes(self):
        self.btn_iniciar.configure(state="normal"); self.btn_parar.configure(state="disabled")
        self.progress_bar.set(0); self.lbl_status_txt.configure(text="Pronto.")

    def verificar_fila(self):
        try:
            while True:
                task = self.queue.get_nowait()
                acao = task.get('acao')
                if acao == 'log': self.log_tela(task['texto'])
                elif acao == 'progresso':
                    self.progress_bar.set(task['valor'] / task['max'])
                    self.lbl_status_txt.configure(text=task['status'])
                    self.lbl_estimativa.configure(text=task['estimativa'])
                elif acao == 'corrigir': self.abrir_corretor_visual(task['lista'])
                elif acao == 'salvar_final': threading.Thread(target=self.salvar_dados, args=(self.dados_temporarios,)).start()
                elif acao == 'msg_fim':
                    messagebox.showinfo("Fim", task['texto'])
                    self.btn_pdf.configure(state="normal", fg_color=COLOR_CARD)
                    self.restaurar_botoes()
                elif acao == 'msg_erro': messagebox.showerror("ERRO", task['texto']); self.restaurar_botoes()
                self.queue.task_done()
        except queue.Empty: pass
        finally: self.after(100, self.verificar_fila)

    def abrir_corretor_visual(self, lista):
        janela = ctk.CTkToplevel(self); janela.title("Confirmação Manual"); janela.geometry("1000x750"); janela.attributes('-topmost', True); janela.transient(self); janela.grab_set()
        janela.grid_columnconfigure(0, weight=3); janela.grid_columnconfigure(1, weight=2); janela.grid_rowconfigure(0, weight=1)
        frame_img = ctk.CTkFrame(janela, fg_color="black", corner_radius=0); frame_img.grid(row=0, column=0, sticky="nsew")
        lbl_foto = ctk.CTkLabel(frame_img, text=""); lbl_foto.place(relx=0.5, rely=0.5, anchor="center")
        frame_info = ctk.CTkFrame(frame_img, fg_color="#202020", corner_radius=8); frame_info.place(relx=0.5, rely=0.9, anchor="center")
        lbl_info = ctk.CTkLabel(frame_info, text="Carregando...", font=("Consolas", 14), text_color="white"); lbl_info.pack(padx=15, pady=8)
        frame_decisao = ctk.CTkFrame(janela, fg_color=COLOR_CARD, corner_radius=0); frame_decisao.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(frame_decisao, text="QUEM É?", font=("Arial", 20, "bold"), text_color=COLOR_ACCENT).pack(pady=(40, 20))
        scroll_btns = ctk.CTkScrollableFrame(frame_decisao, fg_color="transparent"); scroll_btns.pack(fill="both", expand=True, padx=20)
        idx = [0]
        def finalizar(): janela.destroy(); self.queue.put({'acao': 'salvar_final'})
        def show():
            if idx[0] >= len(lista): finalizar(); return
            item = lista[idx[0]]
            lbl_foto.configure(image=None, text="Carregando..."); janela.update()
            lbl_info.configure(text=f"Foto {idx[0]+1}/{len(lista)}\n{item['data']} às {item['hora']}")
            try:
                p = item['caminho_completo']
                if os.path.exists(p):
                    pil = Image.open(p); ratio = min(550/pil.width, 550/pil.height); tk_i = ctk.CTkImage(light_image=pil, dark_image=pil, size=(int(pil.width*ratio), int(pil.height*ratio)))
                    lbl_foto.configure(image=tk_i, text=""); lbl_foto.image = tk_i
                else: lbl_foto.configure(text="Arquivo não encontrado", image=None)
            except Exception as e: lbl_foto.configure(text=f"Erro: {e}", image=None)
        def set_n(n):
            lista[idx[0]]['nome'] = n
            if n != "Desconhecido": threading.Thread(target=self.aprender_rosto, args=(lista[idx[0]]['caminho_completo'], n)).start()
            idx[0] += 1; show()
        nomes = sorted(list(set(self.conhecidos_nom)))
        for n in nomes: ctk.CTkButton(scroll_btns, text=n, height=45, fg_color="#1e293b", hover_color="#3b82f6", command=lambda x=n: set_n(x)).pack(fill="x", pady=4)
        ctk.CTkButton(frame_decisao, text="IGNORAR", height=50, fg_color=COLOR_DANGER, hover_color="#991b1b", command=lambda: set_n("Desconhecido")).pack(fill="x", padx=20, pady=30)
        show()

    # --- MOTOR ---
    def iniciar_thread(self):
        if not self.caminho_zip: messagebox.showwarning("Aviso", "Selecione ZIP!"); return
        if not os.path.exists(self.pasta_funcionarios): os.makedirs(self.pasta_funcionarios)
        self.save_config()
        self.parar_execucao = False
        self.btn_iniciar.configure(state="disabled"); self.btn_parar.configure(state="normal"); self.btn_pdf.configure(state="disabled")
        threading.Thread(target=self.wrapper_processar).start()

    def wrapper_processar(self):
        try: self.processar()
        except Exception as e: log_debug(f"ERRO: {e}"); logging.error(traceback.format_exc()); self.queue.put({'acao': 'msg_erro', 'texto': f"Erro: {e}"})
        finally:
            if self.temp_dir and os.path.exists(self.temp_dir): pass

    def preparing_arquivos(self):
        self.queue.put({'acao': 'log', 'texto': "📦 Extraindo ZIP..."})
        self.temp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(self.caminho_zip, 'r') as zip_ref: zip_ref.extractall(self.temp_dir)
        caminho_txt = None; media_files = []
        for root, dirs, files in os.walk(self.temp_dir):
            for file in files:
                if file.endswith(".txt") and "chat" in file.lower(): caminho_txt = os.path.join(root, file)
                elif file.endswith(".txt") and not caminho_txt: caminho_txt = os.path.join(root, file)
                if "-WA" in file or file.lower().endswith(('.jpg','.opus','.mp4','.webp')): media_files.append(os.path.join(root, file))
        if not caminho_txt: raise Exception("ZIP inválido.")
        media_files.sort()
        return self.temp_dir, caminho_txt, media_files

    def obter_horarios_validos(self, caminho_txt, d_ini, d_fim):
        horarios = []
        padrao = re.compile(r'^(\d{2}/\d{2}/\d{4})\s(\d{2}:\d{2})')
        try:
            with open(caminho_txt, 'r', encoding='utf-8') as f:
                for linha in f:
                    if "<Mídia oculta>" in linha or "(arquivo anexado)" in linha or "(anexado)" in linha or ".jpg" in linha or ".opus" in linha:
                        m = padrao.search(linha)
                        if m:
                            d, h = m.groups()
                            try:
                                dt = datetime.datetime.strptime(d, "%d/%m/%Y").date()
                                if d_ini <= dt <= d_fim: horarios.append({'data': d, 'hora': h})
                            except: pass
            return horarios
        except: raise Exception("Erro TXT")

    def aprender_rosto(self, origem, nome):
        try:
            ts = int(time.time()); dest = os.path.join(self.pasta_funcionarios, f"{nome}_auto_{ts}.jpg")
            shutil.copy(origem, dest); self.queue.put({'acao': 'log', 'texto': f"🧠 Aprendido: {nome}"})
        except: pass

    def salvar_dados(self, dados):
        if not dados: self.queue.put({'acao': 'msg_fim', 'texto': 'Nada para salvar.'}); return
        self.dados_consolidados = dados
        self.queue.put({'acao': 'log', 'texto': f"📤 Salvando {len(dados)} registros..."})
        lista = [[d['nome'], d['data'], d['hora'], os.path.basename(d['caminho_completo'])] for d in dados]
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
            sheet = gspread.authorize(creds).open(NOME_PLANILHA_GOOGLE).sheet1
            sheet.append_rows(lista)
            self.queue.put({'acao': 'msg_fim', 'texto': "Processo Finalizado com Sucesso!"})
        except: self.queue.put({'acao': 'msg_fim', 'texto': "Erro Google Sheets (PDF Disponível)."})
        finally:
            if self.temp_dir and os.path.exists(self.temp_dir):
                try: shutil.rmtree(self.temp_dir); log_debug("Temp limpo.")
                except: pass

    def processar(self):
        try: pasta_temp, caminho_txt, todos_arquivos = self.preparing_arquivos()
        except Exception as e: self.queue.put({'acao': 'msg_erro', 'texto': str(e)}); return
        d_ini = self.cal_inicio.get_date(); d_fim = self.cal_fim.get_date()
        self.queue.put({'acao': 'log', 'texto': f"--- Iniciando {d_ini} a {d_fim} ---"})
        lista_horarios = self.obter_horarios_validos(caminho_txt, d_ini, d_fim)
        limite = min(len(lista_horarios), len(todos_arquivos))
        if limite == 0: self.queue.put({'acao': 'msg_fim', 'texto': "Sem correspondência."}); return
        fila_para_reconhecer = []; pular = 0
        for i in range(limite):
            arq = todos_arquivos[i]; hr = lista_horarios[i]
            if arq.lower().endswith(('.jpg', '.jpeg', '.png')): fila_para_reconhecer.append({'caminho': arq, 'data': hr['data'], 'hora': hr['hora']})
            else: pular += 1
        self.queue.put({'acao': 'log', 'texto': f"ℹ️ Sincronia: {pular} mídias não-foto ignoradas."})
        self.conhecidos_enc = []; self.conhecidos_nom = []
        self.queue.put({'acao': 'log', 'texto': "🧠 Carregando Faces..."})
        if os.path.exists(self.pasta_funcionarios):
            for f in os.listdir(self.pasta_funcionarios):
                if f.lower().endswith(('jpg', 'png')):
                    try:
                        im = face_recognition.load_image_file(os.path.join(self.pasta_funcionarios, f))
                        enc = face_recognition.face_encodings(im)
                        if not enc:
                            locs = face_recognition.face_locations(im, number_of_times_to_upsample=2, model="hog")
                            enc = face_recognition.face_encodings(im, known_face_locations=locs)
                        if enc: self.conhecidos_enc.append(enc[0]); self.conhecidos_nom.append(os.path.splitext(f)[0].split('_')[0].capitalize())
                    except: pass
        self.dados_temporarios = []; total = len(fila_para_reconhecer)
        self.queue.put({'acao': 'config_max', 'valor': total})
        self.queue.put({'acao': 'log', 'texto': f"🚀 Analisando {total} fotos..."})
        tol = self.slider.get(); start = time.time()
        for idx, item in enumerate(fila_para_reconhecer):
            if self.parar_execucao: break
            elapsed = time.time() - start
            est = f"Restam {divmod(int((elapsed/(idx+1))*(total-(idx+1))), 60)[0]}m"
            self.queue.put({'acao': 'progresso', 'valor': idx+1, 'max': total, 'estimativa': est, 'status': f"Processando {idx+1}/{total}"})
            try:
                p = item['caminho']; im = face_recognition.load_image_file(p)
                small = np.ascontiguousarray(im[0::2, 0::2])
                encs = face_recognition.face_encodings(small)
                if not encs: encs = face_recognition.face_encodings(im)
                if not encs:
                    locs = face_recognition.face_locations(im, number_of_times_to_upsample=2, model="hog")
                    encs = face_recognition.face_encodings(im, known_face_locations=locs)
                nomes = []
                if encs:
                    for e in encs:
                        mat = face_recognition.compare_faces(self.conhecidos_enc, e, tolerance=tol)
                        if True in mat: nomes.append(self.conhecidos_nom[mat.index(True)])
                        else: nomes.append("Desconhecido")
                else: nomes.append("Desconhecido")
                for n in nomes:
                    self.dados_temporarios.append({'nome': n, 'data': item['data'], 'hora': item['hora'], 'caminho_completo': p, 'arquivo_origem': os.path.basename(p)})
                    self.queue.put({'acao': 'log', 'texto': f"✅ {n}"})
            except: self.queue.put({'acao': 'log', 'texto': f"❌ Erro foto"})
        desconhecidos = [d for d in self.dados_temporarios if d['nome'] == "Desconhecido"]
        if desconhecidos and not self.parar_execucao: self.queue.put({'acao': 'corrigir', 'lista': desconhecidos})
        else: self.queue.put({'acao': 'salvar_final'})

    def gerar_pdf(self, filepath, data):
        try:
            c = canvas.Canvas(filepath, pagesize=A4)
            w, h = A4
            resumo_global = {}
            meta_diaria = timedelta(hours=7, minutes=30)

            if not data:
                c.drawString(50, h - 50, "Nenhum dado para gerar o relatório.")
                c.save()
                return

            is_report_data = 'Entrada' in data[0]

            if not is_report_data:
                # This is consolidated data
                for d in data:
                    nome = d['nome']
                    if nome not in resumo_global:
                        resumo_global[nome] = {'dias': {}, 'saldo': timedelta(0)}
                    if d['data'] not in resumo_global[nome]['dias']:
                        resumo_global[nome]['dias'][d['data']] = []
                    resumo_global[nome]['dias'][d['data']].append(d['hora'])
            else:
                # This is report data
                for d in data:
                    nome = d['Nome']
                    if nome not in resumo_global:
                        resumo_global[nome] = {'dias': {}, 'saldo': timedelta(0)}
                    if d['Data'] not in resumo_global[nome]['dias']:
                        resumo_global[nome]['dias'][d['Data']] = []
                    resumo_global[nome]['dias'][d['Data']].append(d['Entrada'])
                    resumo_global[nome]['dias'][d['Data']].append(d['Saída'])

            for nome, dados in resumo_global.items():
                for dt, horas in dados['dias'].items():
                    horas.sort()
                    if len(horas) >= 2:
                        ent_str = horas[0]
                        sai_str = horas[-1]
                        if ent_str and sai_str and ent_str != '--:--' and sai_str != '--:--':
                            ent = datetime.datetime.strptime(ent_str, "%H:%M")
                            sai = datetime.datetime.strptime(sai_str, "%H:%M")
                            if ent != sai:
                                trabalhado = (sai - ent) - timedelta(hours=1)
                                dados['saldo'] += (trabalhado - meta_diaria)

            y = h - 50
            c.setFont("Helvetica-Bold", 18)
            c.drawString(50, y, "Relatório Executivo de Ponto")
            y -= 40
            c.setFillColor(colors.black)
            c.rect(50, y, 495, 25, fill=True, stroke=False)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(60, y + 8, "FUNCIONÁRIO")
            c.drawString(400, y + 8, "SALDO TOTAL")
            y -= 30

            def fmt_delta(td):
                s = int(td.total_seconds())
                sign = "+" if s >= 0 else "-"
                s = abs(s)
                return f"{sign}{s//3600:02d}:{(s%3600)//60:02d}"

            for nome in sorted(resumo_global.keys()):
                saldo = resumo_global[nome]['saldo']
                cor = colors.green if saldo.total_seconds() >= 0 else colors.red
                c.setFillColor(colors.black)
                c.setFont("Helvetica", 11)
                c.drawString(60, y, nome)
                c.setFillColor(cor)
                c.setFont("Helvetica-Bold", 11)
                c.drawString(400, y, fmt_delta(saldo))
                c.setStrokeColor(colors.lightgrey)
                c.line(50, y - 5, 545, y - 5)
                y -= 25

            c.showPage()
            y = h - 50

            for nome in sorted(resumo_global.keys()):
                if y < 150:
                    c.showPage()
                    y = h - 50
                c.setFillColor(colors.darkblue)
                c.setFont("Helvetica-Bold", 14)
                c.drawString(50, y, f"Extrato: {nome}")
                y -= 25
                c.setFillColor(colors.lightgrey)
                c.rect(50, y, 495, 15, fill=True, stroke=False)
                c.setFillColor(colors.black)
                c.setFont("Helvetica-Bold", 9)
                c.drawString(55, y + 4, "DATA")
                c.drawString(130, y + 4, "ENTRADA")
                c.drawString(200, y + 4, "SAÍDA")
                c.drawString(270, y + 4, "STATUS")
                y -= 20
                dias = resumo_global[nome]['dias']

                datas_ordenadas = sorted(dias.keys(), key=lambda x: datetime.datetime.strptime(x, "%d/%m/%Y"))

                for dt in datas_ordenadas:
                    horas = dias[dt]
                    horas.sort()
                    ent = horas[0]
                    sai = horas[-1]
                    status = "OK"
                    cor_st = colors.black
                    if ent == sai:
                        status = "Ponto Incompleto"
                        cor_st = colors.orange
                        sai = "--:--"
                    c.setFillColor(colors.black)
                    c.setFont("Helvetica", 10)
                    c.drawString(55, y, dt)
                    c.drawString(130, y, ent)
                    c.drawString(200, y, sai)
                    c.setFillColor(cor_st)
                    c.drawString(270, y, status)
                    y -= 15
                    if y < 50:
                        c.showPage()
                        y = h - 50
                y -= 30
            c.save()
        except Exception as e:
            raise e

    def gerar_pdf_acao_wrapper(self):
        if not self.dados_consolidados:
            messagebox.showwarning("Aviso", "Nenhum dado consolidado para gerar PDF.")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not filepath:
            return
        try:
            self.gerar_pdf(filepath, self.dados_consolidados)
            messagebox.showinfo("Sucesso", "PDF Gerado!")
            self.select_frame_by_name("Relatórios") # Muda para a aba de relatórios
            # Tenta abrir o arquivo (pode não funcionar em todos os sistemas)
            try:
                if os.name == 'nt': # Windows
                    os.startfile(filepath)
                elif os.name == 'posix': # macOS, Linux
                    subprocess.call(('open', filepath) if sys.platform == 'darwin' else ('xdg-open', filepath))
            except:
                pass # Não faz nada se não conseguir abrir
        except Exception as e:
            messagebox.showerror("Erro PDF", str(e))

    def _get_color(self, color_name):
        return THEMES[self.current_theme][color_name]

    def _toggle_theme(self):
        self.current_theme = "Diurno" if self.current_theme == "Noturno" else "Noturno"
        self._apply_theme()

    def _apply_theme(self):
        """Aplica as cores do tema atual a todos os widgets da aplicação."""

        # Cores
        bg_color = self._get_color("COLOR_BG")
        card_color = self._get_color("COLOR_CARD")
        border_color = self._get_color("COLOR_BORDER")
        text_main_color = self._get_color("COLOR_TEXT_MAIN")
        text_dim_color = self._get_color("COLOR_TEXT_DIM")
        accent_color = self._get_color("COLOR_ACCENT")
        accent_hover_color = self._get_color("COLOR_BTN_HOVER")
        info_color = self._get_color("COLOR_INFO")
        danger_color = self._get_color("COLOR_DANGER")

        # Janela Principal
        self.configure(fg_color=bg_color)

        # Header
        self.header.configure(fg_color=bg_color)
        self.lbl_title.configure(text_color=text_main_color)

        # --- ÍCONES ---
        self.icons["Processamento"] = load_icon("rocket", self.current_theme)
        self.icons["Funcionários"] = load_icon("users", self.current_theme)
        self.icons["Relatórios"] = load_icon("clipboard-list", self.current_theme)

        for name, btn in self.nav_buttons.items():
            btn.configure(image=self.icons.get(name))

        self.theme_switch.configure(
            text="Tema Claro" if self.current_theme == "Noturno" else "Tema Escuro",
            text_color=text_dim_color,
            button_color=accent_color,
            progress_color=card_color,
            button_hover_color=accent_hover_color
        )
        self.btn_parar.configure(fg_color=danger_color, text_color=text_main_color, hover_color="#B91C1C")
        self.btn_parar.grid(row=0, column=2, sticky="e", padx=10)

        # --- NAVEGAÇÃO ---
        self.nav_frame.configure(fg_color=card_color)
        self.select_frame_by_name(self.selected_frame_name) # Re-aplica o estilo do botão ativo

        # --- FRAMES DE CONTEÚDO ---
        for frame in self.frames.values():
            if isinstance(frame, (AbaProcessamento, AbaFuncionarios, AbaRelatorios)):
                frame.configure(fg_color=bg_color)

        # --- Aba de Processamento ---
        aba_proc = self.frames["Processamento"]
        aba_proc.configure(fg_color=bg_color)
        aba_proc.frame_left.configure(fg_color=card_color, border_color=border_color)
        aba_proc.lbl_de.configure(text_color=text_dim_color)
        aba_proc.lbl_ate.configure(text_color=text_dim_color)
        aba_proc.separator1.configure(fg_color=border_color)
        aba_proc.separator2.configure(fg_color=border_color)
        self.slider.configure(progress_color=accent_color, button_color=accent_color, button_hover_color=accent_hover_color)
        self.lbl_slider_value.configure(text_color=text_dim_color)
        self.frame_file.configure(border_color=border_color, fg_color=bg_color)
        self.lbl_arquivo.configure(text_color=text_dim_color)
        aba_proc.btn_select.configure(fg_color=info_color, text_color=text_main_color)
        self.btn_iniciar.configure(fg_color=accent_color, hover_color=accent_hover_color, text_color=self._get_color("COLOR_TEXT_MAIN") if self.current_theme == 'Diurno' else '#000000')
        self.btn_pdf.configure(border_color=border_color, text_color=text_dim_color)

        aba_proc.frame_right.configure(fg_color=bg_color)
        aba_proc.card_status.configure(fg_color=card_color, border_color=border_color)
        aba_proc.lbl_status_title.configure(text_color=text_dim_color)
        self.progress_bar.configure(progress_color=accent_color)
        self.lbl_status_txt.configure(text_color=accent_color)
        self.lbl_estimativa.configure(text_color=text_dim_color)
        self.txt_log.configure(fg_color=bg_color, text_color="#22c55e" if self.current_theme=="Noturno" else "#16a34a", border_color=border_color, border_width=1)

        # --- Aba de Funcionários ---
        aba_func = self.frames["Funcionarios"]
        aba_func.configure(fg_color=bg_color)
        aba_func.btn_add.configure(fg_color=accent_color, hover_color=accent_hover_color, text_color=self._get_color("COLOR_TEXT_MAIN") if self.current_theme == 'Diurno' else '#000000')
        aba_func.btn_reload.configure(fg_color=info_color, text_color=text_main_color)
        aba_func.filter_frame.configure(fg_color=card_color, border_color=border_color)
        aba_func.lbl_filter_name.configure(text_color=text_dim_color)
        aba_func.lbl_sort_by.configure(text_color=text_dim_color)
        self.filtro_nome.configure(fg_color=bg_color, border_color=border_color, text_color=text_main_color)
        self.filtro_ordem_campo.configure(button_color=accent_color, border_color=border_color, text_color=text_main_color, dropdown_fg_color=card_color, button_hover_color=accent_hover_color)
        self.filtro_ordem_dir.configure(button_color=accent_color, border_color=border_color, text_color=text_main_color, dropdown_fg_color=card_color, button_hover_color=accent_hover_color)
        self.scroll_func.configure(fg_color=bg_color)
        self.carregar_lista_funcionarios() # Recarrega para aplicar tema nos cards

        # --- Aba de Relatórios ---
        aba_rel = self.frames["Relatorios"]
        aba_rel.configure(fg_color=bg_color)
        aba_rel.frame_controles.configure(fg_color=card_color, border_color=border_color)
        aba_rel.lbl_periodo.configure(text_color=text_dim_color)
        aba_rel.lbl_ate.configure(text_color=text_dim_color)
        aba_rel.lbl_funcionario.configure(text_color=text_dim_color)
        self.combo_funcionarios.configure(button_color=accent_color, border_color=border_color, text_color=text_main_color, dropdown_fg_color=card_color, button_hover_color=accent_hover_color)
        aba_rel.btn_gerar_relatorio.configure(fg_color=accent_color, hover_color=accent_hover_color, text_color=self._get_color("COLOR_TEXT_MAIN") if self.current_theme == 'Diurno' else '#000000')
        self.txt_relatorio.configure(fg_color=bg_color, text_color=text_main_color, border_color=border_color, border_width=1)
        self.btn_export_csv.configure(fg_color=info_color)
        self.btn_export_excel.configure(fg_color=info_color)
        self.btn_export_pdf.configure(fg_color=info_color)
        aba_rel.lbl_hist_title.configure(text_color=text_dim_color)
        self.scroll_historico.configure(fg_color=card_color)
        self.lbl_historico_vazio.configure(text_color=text_dim_color)
        self.atualizar_visualizacao_historico() # Recarrega para aplicar tema nos cards

if __name__ == "__main__":
    app = AppPonto()
    app.mainloop()