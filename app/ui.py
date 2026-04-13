import threading
import traceback
import time
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import json
import os

import pandas as pd
import customtkinter as ctk

from utils import clean_dataframe, generate_messages
from database import HistoryDB
from whatsapp import WhatsAppSender


# Aparência geral
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


GS_COLORS = {
    "orange": "#F57C00",
    "orange_hover": "#D96C00",
    "black": "#111111",
    "graphite": "#1C1C1C",
    "gray": "#2A2A2A",
    "light": "#F4F4F4",
    "white": "#FFFFFF",
    "success": "#2E8B57",
    "danger": "#C62828",
    "warning": "#F9A825",
}


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("GS Trator | Cobrança via WhatsApp")
        self.geometry("1450x1000")
        self.configure(fg_color=GS_COLORS["black"])

        self.db = HistoryDB()
        self.df_original = pd.DataFrame()
        self.df_filtered = pd.DataFrame()
        self.generated_messages = []
        self.sender = None
        self.is_sending = False
        self.stop_requested = False

        self._build_variables()
        self._build_layout()
        self.load_templates()
        self._configure_treeview_style()

    def _build_variables(self):
        self.df_original = pd.DataFrame()
        self.df_filtered = pd.DataFrame()
        self.generated_messages = []
        self.stop_requested = False

        self.templates_file = "tamplate.json"
        self.templates_data = {"tamplates": []}
        self.template_names = []
        self.selected_template_var = tk.StringVar(value="")
        self.template_name_var = tk.StringVar(value="")

        self.excel_path_var = tk.StringVar(value="")
        self.phone_column_var = tk.StringVar(value="")
        self.skip_invalid_var = tk.BooleanVar(value=True)
        self.interval_var = tk.StringVar(value="7")
        self.driver_path_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Pronto.")
        self.total_var = tk.StringVar(value="0 mensagens geradas")

        self.template_default = (
            "Olá {Historico}, tudo bem?\n\n"
            "Consta em nosso sistema o documento {Documento} no valor de {Vl.Documento}.\n\n"
            "Caso já tenha realizado o pagamento, por favor desconsidere esta mensagem.\n\n"
            "Contato enviado para {telefone}.\n\n"
            "Obrigado."
        )

    def _configure_treeview_style(self):
        style = ttk.Style()
        style.theme_use("default")

        style.configure(
            "Treeview",
            background=GS_COLORS["graphite"],
            foreground=GS_COLORS["white"],
            fieldbackground=GS_COLORS["graphite"],
            rowheight=28,
            borderwidth=0,
            font=("Segoe UI", 10)
        )

        style.configure(
            "Treeview.Heading",
            background=GS_COLORS["orange"],
            foreground=GS_COLORS["white"],
            relief="flat",
            font=("Segoe UI", 10, "bold")
        )

        style.map(
            "Treeview",
            background=[("selected", GS_COLORS["orange_hover"])],
            foreground=[("selected", GS_COLORS["white"])]
        )

        style.map(
            "Treeview.Heading",
            background=[("active", GS_COLORS["orange_hover"])]
        )

    def load_templates(self):
        try:
            if not os.path.exists(self.templates_file):
                with open(self.templates_file, "w", encoding="utf-8") as f:
                    json.dump({"tamplates": []}, f, ensure_ascii=False, indent=2)

            with open(self.templates_file, "r", encoding="utf-8") as f:
                self.templates_data = json.load(f)

            if "tamplates" not in self.templates_data or not isinstance(self.templates_data["tamplates"], list):
                self.templates_data = {"tamplates": []}

        except Exception as e:
            self.templates_data = {"tamplates": []}
            self.log(f"[ERRO] Falha ao carregar templates: {e}")

        self.refresh_template_menu()
    
    def get_template_text(self):
        return self.template_box.get("1.0", "end").strip()

    def new_template(self):
        self.selected_template_var.set("")
        self.template_name_var.set("")
        self.template_box.delete("1.0", "end")
        self.template_box.insert("1.0", self.template_default)
        self.log("[INFO] Novo template iniciado.")

    def delete_selected_template(self):
        nome = self.selected_template_var.get().strip()

        if not nome:
            messagebox.showwarning("Aviso", "Selecione um template para excluir.")
            return

        confirm = messagebox.askyesno("Confirmar", f"Deseja excluir o template '{nome}'?")
        if not confirm:
            return

        try:
            self.templates_data["tamplates"] = [
                item for item in self.templates_data.get("tamplates", [])
                if item.get("nome") != nome
            ]

            self.save_templates_file()
            self.refresh_template_menu()

            self.template_name_var.set("")
            self.template_box.delete("1.0", "end")

            self.log(f"[OK] Template excluído: {nome}")
            messagebox.showinfo("Sucesso", f"Template '{nome}' excluído com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir template:\n{e}")
            self.log(f"[ERRO] Falha ao excluir template: {e}")
            
    def save_current_template(self):
        nome = self.template_name_var.get().strip()
        mensagem = self.get_template_text()

        if not nome:
            messagebox.showwarning("Aviso", "Digite um nome para o template.")
            return

        if not mensagem:
            messagebox.showwarning("Aviso", "O template está vazio.")
            return

        found = False
        for item in self.templates_data.get("tamplates", []):
            if item.get("nome") == nome:
                item["mensagem"] = mensagem
                found = True
                break

        if not found:
            self.templates_data["tamplates"].append({
                "nome": nome,
                "mensagem": mensagem
            })

        try:
            self.save_templates_file()
            self.refresh_template_menu()
            self.selected_template_var.set(nome)
            self.log(f"[OK] Template salvo: {nome}")
            messagebox.showinfo("Sucesso", f"Template '{nome}' salvo com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar template:\n{e}")
            self.log(f"[ERRO] Falha ao salvar template: {e}")
    
    def apply_selected_template(self, selected_name=None):
        name = selected_name or self.selected_template_var.get().strip()
        if not name:
            return

        for item in self.templates_data.get("tamplates", []):
            if item.get("nome") == name:
                self.template_box.delete("1.0", "end")
                self.template_box.insert("1.0", item.get("mensagem", ""))
                self.template_name_var.set(name)
                self.log(f"[OK] Template selecionado: {name}")
                return

    def save_templates_file(self):
        with open(self.templates_file, "w", encoding="utf-8") as f:
            json.dump(self.templates_data, f, ensure_ascii=False, indent=2)
        
    def refresh_template_menu(self):
        self.template_names = [item.get("nome", "") for item in self.templates_data.get("tamplates", []) if item.get("nome")]

        values = self.template_names if self.template_names else [""]

        if hasattr(self, "template_select_menu"):
            self.template_select_menu.configure(values=values)

        current = self.selected_template_var.get().strip()
        if current not in self.template_names:
            self.selected_template_var.set(values[0] if values else "")

    def _make_button(self, parent, text, command, color=None, hover=None):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            fg_color=color or GS_COLORS["orange"],
            hover_color=hover or GS_COLORS["orange_hover"],
            text_color=GS_COLORS["white"],
            corner_radius=10,
            height=38,
            font=ctk.CTkFont(size=13, weight="bold")
        )

    def _build_layout(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Painel esquerdo
        left = ctk.CTkFrame(
            self,
            width=370,
            fg_color=GS_COLORS["graphite"],
            corner_radius=14
        )
        left.grid(row=0, column=0, sticky="nsw", padx=12, pady=12)
        left.grid_propagate(False)

        ctk.CTkLabel(
            left,
            text="GS Trator • Configurações",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=22, weight="bold")
        ).pack(anchor="w", padx=14, pady=(14, 10))

        ctk.CTkLabel(left, text="Arquivo Excel", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14)
        ctk.CTkEntry(
            left,
            textvariable=self.excel_path_var,
            width=330,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            text_color=GS_COLORS["white"]
        ).pack(padx=14, pady=5)

        self._make_button(left, "Importar planilha", self.import_excel).pack(padx=14, pady=6, fill="x")

        ctk.CTkLabel(left, text="Coluna do número", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14, pady=(10, 0))

        self.phone_column_menu = ctk.CTkOptionMenu(
            left,
            variable=self.phone_column_var,
            values=[""],
            fg_color=GS_COLORS["orange"],
            button_color=GS_COLORS["orange"],
            button_hover_color=GS_COLORS["orange_hover"],
            text_color=GS_COLORS["white"],
            dropdown_fg_color=GS_COLORS["graphite"],
            dropdown_text_color=GS_COLORS["white"]
        )
        self.phone_column_menu.pack(padx=14, pady=5, fill="x")

        ctk.CTkLabel(left, text="Intervalo entre envios (segundos)", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14, pady=(10, 0))
        ctk.CTkEntry(
            left,
            textvariable=self.interval_var,
            width=110,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            text_color=GS_COLORS["white"]
        ).pack(anchor="w", padx=14, pady=5)

        ctk.CTkLabel(left, text="Caminho do ChromeDriver (opcional)", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14, pady=(10, 0))
        ctk.CTkEntry(
            left,
            textvariable=self.driver_path_var,
            width=330,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            text_color=GS_COLORS["white"]
        ).pack(padx=14, pady=5)

        self._make_button(left, "Selecionar ChromeDriver", self.select_chromedriver).pack(padx=14, pady=6, fill="x")
        self._make_button(left, "Gerar mensagens", self.generate_messages_action).pack(padx=14, pady=(14, 6), fill="x")
        self._make_button(left, "Iniciar WhatsApp / Login", self.start_whatsapp).pack(padx=14, pady=6, fill="x")
        self._make_button(left, "Enviar mensagens", self.send_messages, color=GS_COLORS["success"], hover="#256F47").pack(
            padx=14, pady=6, fill="x"
        )
        self._make_button(left, "Parar envio", self.stop_sending, color=GS_COLORS["danger"], hover="#9E1F1F").pack(
            padx=14, pady=6, fill="x"
        )

        self.progress = ctk.CTkProgressBar(
            left,
            progress_color=GS_COLORS["orange"],
            fg_color=GS_COLORS["gray"]
        )
        self.progress.pack(padx=14, pady=(14, 6), fill="x")
        self.progress.set(0)

        ctk.CTkLabel(left, textvariable=self.total_var, text_color=GS_COLORS["light"]).pack(anchor="w", padx=14)
        ctk.CTkLabel(
            left,
            textvariable=self.status_var,
            wraplength=325,
            justify="left",
            text_color=GS_COLORS["white"]
        ).pack(anchor="w", padx=14, pady=(10, 8))

        # Painel direito
        right = ctk.CTkFrame(
            self,
            fg_color=GS_COLORS["graphite"],
            corner_radius=14
        )
        right.grid(row=0, column=1, sticky="nsew", padx=(0, 12), pady=12)
        right.grid_columnconfigure(0, weight=1)

        for i in range(13):
            right.grid_rowconfigure(i, weight=0)

        right.grid_rowconfigure(5, weight=1)   # template_box
        right.grid_rowconfigure(8, weight=1)   # preview
        right.grid_rowconfigure(10, weight=1)  # tabela
        right.grid_rowconfigure(12, weight=1)  # log

        ctk.CTkLabel(
            right,
            text="Templates salvos",
            text_color=GS_COLORS["light"]
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 0))

        self.template_select_menu = ctk.CTkOptionMenu(
            right,
            variable=self.selected_template_var,
            values=[""],
            command=self.apply_selected_template,
            fg_color=GS_COLORS["orange"],
            button_color=GS_COLORS["orange"],
            button_hover_color=GS_COLORS["orange_hover"],
            text_color=GS_COLORS["white"],
            dropdown_fg_color=GS_COLORS["graphite"],
            dropdown_text_color=GS_COLORS["white"]
        )
        self.template_select_menu.grid(row=1, column=0, sticky="ew", padx=14, pady=5)

        ctk.CTkLabel(
            right,
            text="Nome do template",
            text_color=GS_COLORS["light"]
        ).grid(row=2, column=0, sticky="w", padx=14, pady=(8, 0))

        self.template_name_entry = ctk.CTkEntry(
            right,
            textvariable=self.template_name_var,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            text_color=GS_COLORS["white"]
        )
        self.template_name_entry.grid(row=3, column=0, sticky="ew", padx=14, pady=5)

        ctk.CTkLabel(
            right,
            text="Template da mensagem",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=4, column=0, sticky="w", padx=12, pady=(8, 6))

        self.template_box = ctk.CTkTextbox(
            right,
            height=180,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            border_width=1,
            text_color=GS_COLORS["white"]
        )
        self.template_box.grid(row=5, column=0, sticky="nsew", padx=12, pady=(0, 8))
        self.template_box.insert("1.0", self.template_default)

        template_buttons = ctk.CTkFrame(right, fg_color="transparent")
        template_buttons.grid(row=6, column=0, sticky="w", padx=14, pady=(0, 8))

        ctk.CTkButton(
            template_buttons,
            text="Novo",
            command=self.new_template,
            fg_color=GS_COLORS["graphite"],
            hover_color=GS_COLORS["gray"]
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            template_buttons,
            text="Salvar",
            command=self.save_current_template,
            fg_color=GS_COLORS["orange"],
            hover_color=GS_COLORS["orange_hover"]
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            template_buttons,
            text="Excluir",
            command=self.delete_selected_template,
            fg_color="#8B0000",
            hover_color="#A40000"
        ).pack(side="left", padx=4)

        ctk.CTkLabel(
            right,
            text="Preview / mensagens geradas",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=7, column=0, sticky="w", padx=12, pady=(4, 6))

        self.preview_box = ctk.CTkTextbox(
            right,
            height=180,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            border_width=1,
            text_color=GS_COLORS["white"]
        )
        self.preview_box.grid(row=8, column=0, sticky="nsew", padx=12, pady=(0, 10))

        ctk.CTkLabel(
            right,
            text="Dados da planilha",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=9, column=0, sticky="w", padx=12, pady=(4, 6))

        table_frame = ctk.CTkFrame(right, fg_color=GS_COLORS["gray"], corner_radius=10)
        table_frame.grid(row=10, column=0, sticky="nsew", padx=12, pady=(0, 10))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_frame,
            columns=(),
            show="headings",
            height=12
        )

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")

        ctk.CTkLabel(
            right,
            text="Log",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=11, column=0, sticky="w", padx=12, pady=(4, 6))

        self.log_box = ctk.CTkTextbox(
            right,
            height=150,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            border_width=1,
            text_color=GS_COLORS["white"]
        )
        self.log_box.grid(row=12, column=0, sticky="nsew", padx=12, pady=(0, 12))
    
    def log(self, text: str):
        if hasattr(self, "log_box"):
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")
            self.update_idletasks()

    def select_chromedriver(self):
        path = filedialog.askopenfilename(
            title="Selecione o ChromeDriver",
            filetypes=[("Executável", "*.exe"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.driver_path_var.set(path)
            self.log(f"[OK] ChromeDriver selecionado: {path}")

    def import_excel(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            self.excel_path_var.set(path)
            df = pd.read_excel(path)

            if df.empty:
                messagebox.showwarning("Aviso", "A planilha está vazia.")
                return

            self.df_original = df.copy()
            self.df_filtered = pd.DataFrame()
            self.generated_messages = []

            columns = [str(col) for col in df.columns]
            self.phone_column_menu.configure(values=columns)

            # tenta selecionar automaticamente uma coluna provável
            preferred = ""
            for candidate in columns:
                candidate_lower = candidate.strip().lower()
                if candidate_lower in ["telefone", "telefone cliente", "celular", "whatsapp", "numero", "número"]:
                    preferred = candidate
                    break

            if not preferred and columns:
                preferred = columns[0]

            self.phone_column_var.set(preferred)

            self.refresh_table(self.df_original)

            self.status_var.set(f"Planilha carregada com {len(df)} registros.")
            self.log(f"[OK] Planilha importada: {path}")
            self.log(f"[INFO] Colunas encontradas: {', '.join(columns)}")
            self.log(f"[INFO] Coluna de número selecionada: {preferred}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao importar planilha:\n{e}")
            self.log(f"[ERRO] Importação falhou: {e}")

    def refresh_table(self, df: pd.DataFrame):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df.empty:
            self.tree["columns"] = ()
            return

        preview_df = df.head(300).copy()
        columns = [str(col) for col in preview_df.columns]

        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor="w")

        for _, row in preview_df.iterrows():
            values = [row.get(col, "") for col in columns]
            self.tree.insert("", "end", values=values)

    def generate_messages_action(self):
        if self.df_original.empty:
            messagebox.showwarning("Aviso", "Importe uma planilha primeiro.")
            return

        phone_column = self.phone_column_var.get().strip()
        if not phone_column:
            messagebox.showwarning("Aviso", "Selecione a coluna do número.")
            return

        template = self.get_template_text().strip()
        if not template:
            messagebox.showwarning("Aviso", "Digite um template antes de gerar as mensagens.")
            return

        try:
            self.status_var.set("Gerando mensagens...")
            self.total_var.set("0 mensagens geradas")
            self.preview_box.delete("1.0", "end")
            self.generated_messages = []
            self.df_filtered = pd.DataFrame()

            self.log("[INFO] Iniciando geração de mensagens...")

            self.df_filtered = clean_dataframe(self.df_original, phone_column)

            if self.df_filtered.empty:
                self.status_var.set("Nenhum registro disponível.")
                self.preview_box.insert("1.0", "Nenhum registro disponível após o tratamento da planilha.")
                self.log("[INFO] Nenhum registro disponível após tratamento da planilha.")
                return

            self.refresh_table(self.df_filtered)

            raw_messages = generate_messages(self.df_filtered, template)
            self.generated_messages = [item for item in raw_messages if isinstance(item, dict)]

            if not self.generated_messages:
                self.preview_box.insert("1.0", "Nenhuma mensagem gerada.")
                self.status_var.set("Nenhuma mensagem gerada.")
                self.log("[INFO] Nenhuma mensagem válida foi gerada.")
                return

            preview_texts = []
            for i, item in enumerate(self.generated_messages[:10], start=1):
                telefone = item.get("telefone", "")
                row_data = item.get("row_data") or {}
                nome_preview = row_data.get("Historico", "") or row_data.get("nome", "") or row_data.get("Nome", "")
                mensagem = item.get("mensagem", "")

                preview_texts.append(
                    f"{i}) {telefone} - {nome_preview}\n{mensagem}\n{'-' * 70}"
                )

            self.preview_box.insert("1.0", "\n".join(preview_texts))
            self.total_var.set(f"{len(self.generated_messages)} mensagens geradas")
            self.status_var.set("Mensagens geradas com sucesso.")
            self.log(f"[OK] {len(self.generated_messages)} mensagens geradas.")
            self.log(f"[INFO] Coluna de número usada: {phone_column}")

        except Exception as e:
            self.status_var.set("Erro ao gerar mensagens.")
            self.preview_box.delete("1.0", "end")
            self.preview_box.insert("1.0", f"Erro ao gerar mensagens:\n{e}")
            self.log(f"[ERRO] Geração falhou: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror("Erro", f"Falha ao gerar mensagens:\n{e}")
            
    def start_whatsapp(self):
        driver_path = self.driver_path_var.get().strip()

        def worker():
            try:
                self.status_var.set("Abrindo WhatsApp Web...")
                self.log("[INFO] Iniciando navegador...")
                self.sender = WhatsAppSender(driver_path)
                self.sender.start()
                self.sender.open_whatsapp()

                self.status_var.set("Escaneie o QR Code no WhatsApp Web...")
                self.log("[INFO] Aguardando login manual por QR Code...")
                self.sender.wait_for_login()

                self.status_var.set("WhatsApp conectado.")
                self.log("[OK] WhatsApp conectado com sucesso.")
            except Exception as e:
                error_msg = str(e)
                error_trace = traceback.format_exc()

                self.status_var.set("Erro ao iniciar WhatsApp.")
                self.log(f"[ERRO] WhatsApp: {error_msg}")
                self.log(error_trace)

                self.after(0, lambda msg=error_msg: messagebox.showerror("Erro", f"Falha ao iniciar WhatsApp:\n{msg}"))

        threading.Thread(target=worker, daemon=True).start()

    def stop_sending(self):
        if not self.is_sending:
            self.log("[INFO] Nenhum envio em andamento.")
            return

        self.stop_requested = True
        self.status_var.set("Solicitação de parada enviada...")
        self.log("[INFO] Parada solicitada. O sistema irá interromper após a mensagem atual.")

    def send_messages(self):
        if self.is_sending:
            messagebox.showinfo("Aviso", "Já existe um envio em andamento.")
            return

        if not self.generated_messages:
            messagebox.showwarning("Aviso", "Gere as mensagens primeiro.")
            return

        if not self.sender:
            messagebox.showwarning("Aviso", "Inicie o WhatsApp primeiro.")
            return

        self.is_sending = True
        self.stop_requested = False

        def worker():
            total = len(self.generated_messages)
            sent = 0

            try:
                interval = int(self.interval_var.get() or "7")
                skip_invalid = self.skip_invalid_var.get()

                self.progress.set(0)
                self.status_var.set("Enviando mensagens...")

                for idx, item in enumerate(self.generated_messages, start=1):
                    if self.stop_requested:
                        self.status_var.set("Envio interrompido pelo usuário.")
                        self.log("[PARADO] Envio encerrado manualmente.")
                        break

                    if item is None or not isinstance(item, dict):
                        self.log(f"[ERRO] Item inválido na posição {idx}: {item}")
                        self.progress.set(idx / total)
                        continue

                    nome = item.get("nome", "")
                    documento = item.get("documento", "")
                    telefone = item.get("telefone", "")
                    mensagem = item.get("mensagem", "")
                    telefone_valido = item.get("telefone_valido", False)
                    row_data = item.get("row_data") or {}

                    try:
                        if skip_invalid and not telefone_valido:
                            self.db.save_send(nome, documento, telefone, mensagem, "ignorado", "telefone inválido")
                            self.log(f"[IGNORADO] {telefone} - telefone inválido")
                            self.progress.set(idx / total)
                            continue

                        if not telefone:
                            self.db.save_send(nome, documento, telefone, mensagem, "ignorado", "telefone vazio")
                            self.log(f"[IGNORADO] Registro sem telefone: {row_data}")
                            self.progress.set(idx / total)
                            continue

                        if not mensagem:
                            self.db.save_send(nome, documento, telefone, mensagem, "ignorado", "mensagem vazia")
                            self.log(f"[IGNORADO] {telefone} - mensagem vazia")
                            self.progress.set(idx / total)
                            continue

                        self.status_var.set(f"Carregando conversa... {idx}/{total}")
                        self.log(f"[INFO] Abrindo conversa de {telefone}...")

                        # só retorna True quando realmente enviou
                        self.sender.send_message(telefone, mensagem)

                        self.db.save_send(nome, documento, telefone, mensagem, "enviado", "")
                        sent += 1
                        self.log(f"[ENVIADO] {telefone} - {nome}")

                        self.progress.set(idx / total)
                        self.status_var.set(f"Enviado {idx}/{total}. Aguardando intervalo...")

                        # intervalo começa somente após envio confirmado
                        if idx < total:
                            for sec in range(interval, 0, -1):
                                if self.stop_requested:
                                    self.status_var.set("Envio interrompido pelo usuário.")
                                    self.log("[PARADO] Envio encerrado durante intervalo.")
                                    break

                                self.status_var.set(
                                    f"Enviado {idx}/{total}. Próximo em {sec}s..."
                                )
                                time.sleep(1)

                            if self.stop_requested:
                                break

                    except Exception as inner_error:
                        error_text = str(inner_error)

                        if "não está no WhatsApp" in error_text or "not on WhatsApp" in error_text:
                            self.db.save_send(nome, documento, telefone, mensagem, "ignorado", "não tem whatsapp")
                            self.log(f"[IGNORADO] {telefone} - número não tem WhatsApp")
                        else:
                            self.db.save_send(nome, documento, telefone, mensagem, "erro", error_text)
                            self.log(f"[ERRO] {telefone} - {error_text}")
                            self.log(traceback.format_exc())

                    self.progress.set(idx / total)

                if not self.stop_requested:
                    self.status_var.set(f"Envio concluído. {sent} enviadas.")
                    self.log(f"[FIM] Processo concluído. Total enviadas: {sent}")

            except Exception as e:
                self.status_var.set("Erro no processo de envio.")
                self.log(f"[ERRO GERAL] {e}")
                self.log(traceback.format_exc())
            finally:
                self.is_sending = False
                self.stop_requested = False

        threading.Thread(target=worker, daemon=True).start()