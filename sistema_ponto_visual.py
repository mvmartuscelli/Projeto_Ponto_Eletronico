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
import threading
import queue
import traceback
import logging
import zipfile
import shutil
import tempfile
import numpy as np
import json
from collections import defaultdict
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors

# --- CONFIGURAÇÃO VISUAL ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

COLOR_BG = "#0a0a0a"
COLOR_CARD = "#0f172a"
COLOR_BORDER = "#1e293b"
COLOR_TEXT_MAIN = "#f8fafc"
COLOR_TEXT_DIM = "#94a3b8"
COLOR_ACCENT = "#10b981"
COLOR_BTN_HOVER = "#059669"
COLOR_DANGER = "#ef4444"
COLOR_INFO = "#3b82f6"

# LOG
logging.basicConfig(filename='log_debug.txt', level=logging.DEBUG, format='%(asctime)s - %(message)s', filemode='w')

def log_debug(msg):
    print(msg)
    logging.info(msg)

NOME_PLANILHA_GOOGLE = "PontoFuncionarios"
ARQUIVO_CONFIG = "config.json"

# --- CLASSE PRINCIPAL COM SUPORTE A DND ---
class AppPonto(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        
        # Inicializa Drag & Drop
        self.TkdndVersion = TkinterDnD._require(self)
        
        self.title("Sistema Ponto Neural v16.2 - Estável")
        self.geometry("1200x850")
        self.configure(fg_color=COLOR_BG)
        
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
        
        self.dados_consolidados = []
        self.caminho_zip = ""
        self.conhecidos_nom = []

        # --- ABAS ---
        self.tabview = ctk.CTkTabview(self, fg_color="transparent")
        self.tabview.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.tab_process = self.tabview.add(" 🚀 Processamento ")
        self.tab_func = self.tabview.add(" 👥 Funcionários ")
        
        self.setup_tab_processamento()
        self.setup_tab_funcionarios()

    # ==========================================
    # ABA 1: PROCESSAMENTO
    # ==========================================
    def setup_tab_processamento(self):
        frame = self.tab_process
        frame.grid_columnconfigure(0, weight=4)
        frame.grid_columnconfigure(1, weight=5)
        frame.grid_rowconfigure(0, weight=1)

        # HEADER
        lbl_title = ctk.CTkLabel(frame, text="Painel de Controle", font=("Roboto Medium", 24), text_color="white")
        lbl_title.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=10)

        # ESQUERDA
        frame_left = ctk.CTkFrame(frame, fg_color=COLOR_CARD, corner_radius=15, border_color=COLOR_BORDER, border_width=1)
        frame_left.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=10)
        
        # Datas
        self.create_section_label(frame_left, "1. PERÍODO")
        frame_datas = ctk.CTkFrame(frame_left, fg_color="transparent")
        frame_datas.pack(fill="x", padx=20)
        
        style_cal = {'background': '#1e293b', 'foreground': 'white', 'borderwidth': 0, 'date_pattern': 'dd/mm/yyyy'}
        
        ctk.CTkLabel(frame_datas, text="De:", text_color=COLOR_TEXT_DIM).grid(row=0, column=0, sticky="w")
        self.cal_inicio = DateEntry(frame_datas, width=12, font=("Arial", 11), **style_cal)
        self.cal_inicio.set_date(datetime.date.today().replace(day=1))
        self.cal_inicio.grid(row=1, column=0, padx=5, pady=5)

        ctk.CTkLabel(frame_datas, text="Até:", text_color=COLOR_TEXT_DIM).grid(row=0, column=1, sticky="w")
        self.cal_fim = DateEntry(frame_datas, width=12, font=("Arial", 11), **style_cal)
        self.cal_fim.set_date(datetime.date.today())
        self.cal_fim.grid(row=1, column=1, padx=5, pady=5)

        # Slider
        ctk.CTkFrame(frame_left, height=1, fg_color=COLOR_BORDER).pack(fill="x", padx=20, pady=20)
        self.create_section_label(frame_left, "2. SENSIBILIDADE")
        self.slider = ctk.CTkSlider(frame_left, from_=0.35, to=0.60, number_of_steps=25, progress_color=COLOR_ACCENT)
        self.slider.set(self.config.get('tolerancia', 0.45))
        self.slider.pack(fill="x", padx=25, pady=5)

        # Arquivo (Drag & Drop)
        ctk.CTkFrame(frame_left, height=1, fg_color=COLOR_BORDER).pack(fill="x", padx=20, pady=20)
        self.create_section_label(frame_left, "3. ARQUIVO ZIP (Arraste aqui)")
        
        self.frame_file = ctk.CTkFrame(frame_left, fg_color="#1e293b", corner_radius=8, border_width=2, border_color="#334155")
        self.frame_file.pack(fill="x", padx=20, pady=5, ipady=10)
        
        # Registro DND
        self.frame_file.drop_target_register(DND_FILES)
        self.frame_file.dnd_bind('<<Drop>>', self.drop_file)
        
        self.lbl_arquivo = ctk.CTkLabel(self.frame_file, text="Arraste o ZIP ou clique na pasta", text_color=COLOR_TEXT_DIM, font=("Consolas", 11))
        self.lbl_arquivo.pack(side="left", padx=10, pady=10)
        
        self.btn_select = ctk.CTkButton(self.frame_file, text="📂", width=40, fg_color=COLOR_INFO, command=self.selecionar_zip)
        self.btn_select.pack(side="right", padx=10)

        # Ações
        frame_actions = ctk.CTkFrame(frame_left, fg_color="transparent")
        frame_actions.pack(side="bottom", fill="x", padx=20, pady=30)
        self.btn_iniciar = ctk.CTkButton(frame_actions, text="INICIAR", height=50, font=("Arial", 14, "bold"), fg_color=COLOR_ACCENT, hover_color=COLOR_BTN_HOVER, command=self.iniciar_thread)
        self.btn_iniciar.pack(fill="x", pady=(0, 10))
        self.btn_pdf = ctk.CTkButton(frame_actions, text="GERAR RELATÓRIO PDF", height=40, fg_color="transparent", border_width=1, border_color=COLOR_BORDER, state="disabled", command=self.gerar_pdf_acao)
        self.btn_pdf.pack(fill="x")

        # DIREITA
        frame_right = ctk.CTkFrame(frame, fg_color="transparent")
        frame_right.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=10)
        
        self.card_status = ctk.CTkFrame(frame_right, fg_color=COLOR_CARD, corner_radius=15, border_color=COLOR_BORDER, border_width=1)
        self.card_status.pack(fill="x", pady=(0, 15))
        ctk.CTkLabel(self.card_status, text="STATUS", font=("Arial", 11, "bold"), text_color=COLOR_TEXT_DIM).pack(anchor="w", padx=20, pady=(15, 5))
        self.progress_bar = ctk.CTkProgressBar(self.card_status, height=10, progress_color=COLOR_ACCENT)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=20, pady=5)
        self.lbl_status_txt = ctk.CTkLabel(self.card_status, text="Pronto.", text_color=COLOR_ACCENT)
        self.lbl_status_txt.pack(side="left", padx=20, pady=(0, 15))
        self.lbl_estimativa = ctk.CTkLabel(self.card_status, text="--:--", text_color=COLOR_TEXT_DIM)
        self.lbl_estimativa.pack(side="right", padx=20, pady=(0, 15))

        self.txt_log = ctk.CTkTextbox(frame_right, fg_color="black", text_color="#22c55e", font=("Consolas", 11), corner_radius=10)
        self.txt_log.pack(fill="both", expand=True)
        self.btn_parar = ctk.CTkButton(frame_right, text="PARAR", fg_color="#450a0a", text_color=COLOR_DANGER, width=80, state="disabled", command=self.solicitar_parada)
        self.btn_parar.place(relx=1.0, rely=0.0, anchor="ne", x=0, y=-40)

    # ==========================================
    # ABA 2: GESTÃO DE FUNCIONÁRIOS
    # ==========================================
    def setup_tab_funcionarios(self):
        frame = self.tab_func
        
        toolbar = ctk.CTkFrame(frame, fg_color="transparent", height=50)
        toolbar.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkButton(toolbar, text="+ Adicionar Novo", fg_color=COLOR_ACCENT, hover_color=COLOR_BTN_HOVER, command=self.add_funcionario).pack(side="left")
        ctk.CTkButton(toolbar, text="🔄 Atualizar", fg_color=COLOR_INFO, width=80, command=self.carregar_lista_funcionarios).pack(side="left", padx=10)
        
        self.scroll_func = ctk.CTkScrollableFrame(frame, fg_color="transparent")
        self.scroll_func.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.carregar_lista_funcionarios()

    def carregar_lista_funcionarios(self):
        for widget in self.scroll_func.winfo_children():
            widget.destroy()
            
        if not os.path.exists(self.pasta_funcionarios):
            ctk.CTkLabel(self.scroll_func, text="Nenhum funcionário cadastrado.").pack(pady=20)
            return

        arquivos = sorted([f for f in os.listdir(self.pasta_funcionarios) if f.lower().endswith(('.jpg', '.jpeg', '.png'))])
        funcionarios_dict = defaultdict(list)
        
        for arq in arquivos:
            nome = os.path.splitext(arq)[0].split('_')[0].capitalize()
            funcionarios_dict[nome].append(arq)

        if not funcionarios_dict:
            ctk.CTkLabel(self.scroll_func, text="Nenhum funcionário encontrado.").pack(pady=20)
            return

        for nome, lista_fotos in funcionarios_dict.items():
            self.criar_grupo_funcionario(nome, lista_fotos)

    def criar_grupo_funcionario(self, nome, fotos):
        card = ctk.CTkFrame(self.scroll_func, fg_color=COLOR_CARD, corner_radius=10, border_color=COLOR_BORDER, border_width=1)
        card.pack(fill="x", padx=5, pady=5)
        
        header = ctk.CTkFrame(card, fg_color="transparent", height=50)
        header.pack(fill="x", padx=10, pady=5)
        
        try:
            path = os.path.join(self.pasta_funcionarios, fotos[0])
            pil = Image.open(path)
            pil.thumbnail((40, 40))
            tk_img = ctk.CTkImage(light_image=pil, dark_image=pil, size=(40, 40))
            lbl_img = ctk.CTkLabel(header, image=tk_img, text="")
            lbl_img.pack(side="left")
        except:
            ctk.CTkLabel(header, text="👤", font=("Arial", 20)).pack(side="left", padx=10)

        ctk.CTkLabel(header, text=nome, font=("Arial", 14, "bold"), text_color="white").pack(side="left", padx=15)
        ctk.CTkLabel(header, text=f"{len(fotos)} variações", font=("Arial", 11), text_color=COLOR_TEXT_DIM).pack(side="left", padx=5)

        btn_expand = ctk.CTkButton(header, text="▼", width=30, fg_color="transparent", text_color=COLOR_INFO, hover_color="#1e293b")
        btn_expand.pack(side="right")

        body = ctk.CTkFrame(card, fg_color="#0b1120")
        
        for f in fotos:
            row = ctk.CTkFrame(body, fg_color="transparent", height=30)
            row.pack(fill="x", padx=10, pady=2)
            ctk.CTkLabel(row, text=f"📄 {f}", font=("Consolas", 11), text_color="gray").pack(side="left")
            btn_del = ctk.CTkButton(row, text="🗑️", width=30, height=20, fg_color="transparent", hover_color=COLOR_DANGER,
                                    command=lambda file=f: self.delete_funcionario(file))
            btn_del.pack(side="right")

        def toggle():
            if body.winfo_ismapped():
                body.pack_forget(); btn_expand.configure(text="▼")
            else:
                body.pack(fill="x", padx=10, pady=(0, 10)); btn_expand.configure(text="▲")
        
        btn_expand.configure(command=toggle)

    def add_funcionario(self):
        path = filedialog.askopenfilename(filetypes=[("Imagens", "*.jpg *.jpeg *.png")])
        if not path: return
        nome = ctk.CTkInputDialog(text="Nome (Primeiro nome):", title="Cadastro").get_input()
        if not nome: return
        try:
            nome_limpo = nome.strip().capitalize()
            novo_nome = f"{nome_limpo}_{int(time.time())}.jpg"
            shutil.copy(path, os.path.join(self.pasta_funcionarios, novo_nome))
            self.carregar_lista_funcionarios()
        except Exception as e: messagebox.showerror("Erro", str(e))

    def delete_funcionario(self, filename):
        if messagebox.askyesno("Confirmar", f"Excluir {filename}?"):
            try:
                os.remove(os.path.join(self.pasta_funcionarios, filename))
                self.carregar_lista_funcionarios()
            except Exception as e: messagebox.showerror("Erro", str(e))

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

    # --- UI HELPERS ---
    def create_section_label(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=("Arial", 11, "bold"), text_color=COLOR_TEXT_DIM).pack(anchor="w", padx=20, pady=(0, 5))

    def log_tela(self, msg):
        self.txt_log.insert("end", f"> {msg}\n"); self.txt_log.see("end")

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
        self.progress['value'] = 0; self.lbl_status_txt.configure(text="Pronto.")

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

    def gerar_pdf_acao(self):
        if not self.dados_consolidados: return
        arq = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile="Relatorio.pdf")
        if not arq: return
        try:
            c = canvas.Canvas(arq, pagesize=A4); w, h = A4
            resumo_global = {}
            meta_diaria = timedelta(hours=7, minutes=30)
            for d in self.dados_consolidados:
                nome = d['nome']
                if nome not in resumo_global: resumo_global[nome] = {'dias': {}, 'saldo': timedelta(0)}
                if d['data'] not in resumo_global[nome]['dias']: resumo_global[nome]['dias'][d['data']] = []
                resumo_global[nome]['dias'][d['data']].append(d['hora'])
            for nome, dados in resumo_global.items():
                for dt, horas in dados['dias'].items():
                    horas.sort()
                    if len(horas) >= 2:
                        ent = datetime.datetime.strptime(horas[0], "%H:%M")
                        sai = datetime.datetime.strptime(horas[-1], "%H:%M")
                        if ent != sai:
                            trabalhado = (sai - ent) - timedelta(hours=1)
                            dados['saldo'] += (trabalhado - meta_diaria)
            y = h - 50
            c.setFont("Helvetica-Bold", 18); c.drawString(50, y, "Relatório Executivo de Ponto"); y -= 40
            c.setFillColor(colors.black); c.rect(50, y, 495, 25, fill=True, stroke=False)
            c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 12)
            c.drawString(60, y+8, "FUNCIONÁRIO"); c.drawString(400, y+8, "SALDO TOTAL"); y -= 30
            def fmt_delta(td):
                s = int(td.total_seconds()); sign = "+" if s >= 0 else "-"; s = abs(s)
                return f"{sign}{s//3600:02d}:{(s%3600)//60:02d}"
            for nome in sorted(resumo_global.keys()):
                saldo = resumo_global[nome]['saldo']
                cor = colors.green if saldo.total_seconds() >= 0 else colors.red
                c.setFillColor(colors.black); c.setFont("Helvetica", 11); c.drawString(60, y, nome)
                c.setFillColor(cor); c.setFont("Helvetica-Bold", 11); c.drawString(400, y, fmt_delta(saldo))
                c.setStrokeColor(colors.lightgrey); c.line(50, y-5, 545, y-5); y -= 25
            c.showPage(); y = h - 50
            for nome in sorted(resumo_global.keys()):
                if y < 150: c.showPage(); y = h - 50
                c.setFillColor(colors.darkblue); c.setFont("Helvetica-Bold", 14); c.drawString(50, y, f"Extrato: {nome}"); y -= 25
                c.setFillColor(colors.lightgrey); c.rect(50, y, 495, 15, fill=True, stroke=False)
                c.setFillColor(colors.black); c.setFont("Helvetica-Bold", 9)
                c.drawString(55, y+4, "DATA"); c.drawString(130, y+4, "ENTRADA"); c.drawString(200, y+4, "SAÍDA"); c.drawString(270, y+4, "STATUS"); y -= 20
                dias = resumo_global[nome]['dias']
                datas_ordenadas = sorted(dias.keys(), key=lambda x: datetime.datetime.strptime(x, "%d/%m/%Y"))
                for dt in datas_ordenadas:
                    horas = dias[dt]; horas.sort()
                    ent = horas[0]; sai = horas[-1]
                    status = "OK"; cor_st = colors.black
                    if ent == sai: status = "Ponto Incompleto"; cor_st = colors.orange; sai = "--:--"
                    c.setFillColor(colors.black); c.setFont("Helvetica", 10)
                    c.drawString(55, y, dt); c.drawString(130, y, ent); c.drawString(200, y, sai)
                    c.setFillColor(cor_st); c.drawString(270, y, status); y -= 15
                    if y < 50: c.showPage(); y = h - 50
                y -= 30
            c.save(); messagebox.showinfo("Sucesso", "PDF Gerado!"); os.startfile(arq)
        except Exception as e: messagebox.showerror("Erro PDF", str(e))

    def restaurar_botoes(self):
        self.btn_iniciar.configure(state="normal"); self.btn_parar.configure(state="disabled")
        self.progress['value'] = 0; self.lbl_status_txt.configure(text="Pronto.")

if __name__ == "__main__":
    app = AppPonto()
    app.mainloop()