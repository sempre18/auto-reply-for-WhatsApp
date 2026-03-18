import threading
import traceback
import time
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import pandas as pd
import customtkinter as ctk

from utils import (
    validate_columns,
    clean_dataframe,
    filter_dataframe,
    generate_messages,
)
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
        self.geometry("1450x900")
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
        self._configure_treeview_style()

    def _build_variables(self):
        self.excel_path_var = tk.StringVar(value="")
        self.filter_mode_var = tk.StringVar(value="todos")
        self.days_var = tk.StringVar(value="3")
        self.only_valid_phone_var = tk.BooleanVar(value=True)
        self.skip_invalid_var = tk.BooleanVar(value=True)
        self.skip_duplicates_var = tk.BooleanVar(value=True)
        self.interval_var = tk.StringVar(value="7")
        self.driver_path_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Pronto.")
        self.total_var = tk.StringVar(value="0 mensagens geradas")

        self.template_default = (
            "Olá {nome}, tudo bem?\n\n"
            "Consta em nosso sistema o documento {documento}, "
            "com vencimento em {vencimento}, no valor de {valor}.\n\n"
            "Caso já tenha realizado o pagamento, por favor desconsidere esta mensagem.\n\n"
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

        ctk.CTkLabel(left, text="Filtro", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14, pady=(10, 0))
        ctk.CTkOptionMenu(
            left,
            variable=self.filter_mode_var,
            values=["todos", "vencidos", "a_vencer"],
            fg_color=GS_COLORS["orange"],
            button_color=GS_COLORS["orange"],
            button_hover_color=GS_COLORS["orange_hover"],
            text_color=GS_COLORS["white"],
            dropdown_fg_color=GS_COLORS["graphite"],
            dropdown_text_color=GS_COLORS["white"]
        ).pack(padx=14, pady=5)

        ctk.CTkLabel(left, text="Dias para 'a vencer'", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14)
        ctk.CTkEntry(
            left,
            textvariable=self.days_var,
            width=110,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            text_color=GS_COLORS["white"]
        ).pack(anchor="w", padx=14, pady=5)

        ctk.CTkCheckBox(left, text="Somente telefone válido", variable=self.only_valid_phone_var).pack(anchor="w", padx=14, pady=3)
        ctk.CTkCheckBox(left, text="Pular números inválidos", variable=self.skip_invalid_var).pack(anchor="w", padx=14, pady=3)
        ctk.CTkCheckBox(left, text="Não reenviar duplicados", variable=self.skip_duplicates_var).pack(anchor="w", padx=14, pady=3)

        ctk.CTkLabel(left, text="Intervalo entre envios (segundos)", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14, pady=(10, 0))
        ctk.CTkEntry(
            left,
            textvariable=self.interval_var,
            width=110,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            text_color=GS_COLORS["white"]
        ).pack(anchor="w", padx=14, pady=5)

        ctk.CTkLabel(left, text="Caminho do ChromeDriver", text_color=GS_COLORS["light"]).pack(anchor="w", padx=14, pady=(10, 0))
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
        right.grid_rowconfigure(1, weight=1)
        right.grid_rowconfigure(3, weight=1)
        right.grid_rowconfigure(5, weight=1)
        right.grid_rowconfigure(7, weight=1)

        ctk.CTkLabel(
            right,
            text="Template da mensagem",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        self.template_box = ctk.CTkTextbox(
            right,
            height=140,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            border_width=1,
            text_color=GS_COLORS["white"]
        )
        self.template_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        self.template_box.insert("1.0", self.template_default)

        ctk.CTkLabel(
            right,
            text="Preview / mensagens geradas",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=2, column=0, sticky="w", padx=12, pady=(4, 6))

        self.preview_box = ctk.CTkTextbox(
            right,
            height=180,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            border_width=1,
            text_color=GS_COLORS["white"]
        )
        self.preview_box.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 10))

        ctk.CTkLabel(
            right,
            text="Dados da planilha",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=4, column=0, sticky="w", padx=12, pady=(4, 6))

        table_frame = ctk.CTkFrame(right, fg_color=GS_COLORS["gray"], corner_radius=10)
        table_frame.grid(row=5, column=0, sticky="nsew", padx=12, pady=(0, 10))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_frame,
            columns=("nome", "documento", "vencimento", "valor", "telefone", "valido"),
            show="headings",
            height=12
        )
        self.tree.heading("nome", text="Cliente")
        self.tree.heading("documento", text="Documento")
        self.tree.heading("vencimento", text="Vencimento")
        self.tree.heading("valor", text="Valor")
        self.tree.heading("telefone", text="Telefone")
        self.tree.heading("valido", text="Válido")

        self.tree.column("nome", width=220)
        self.tree.column("documento", width=120)
        self.tree.column("vencimento", width=110)
        self.tree.column("valor", width=110)
        self.tree.column("telefone", width=150)
        self.tree.column("valido", width=70)

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")

        ctk.CTkLabel(
            right,
            text="Log",
            text_color=GS_COLORS["white"],
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=6, column=0, sticky="w", padx=12, pady=(4, 6))

        self.log_box = ctk.CTkTextbox(
            right,
            height=150,
            fg_color=GS_COLORS["gray"],
            border_color=GS_COLORS["orange"],
            border_width=1,
            text_color=GS_COLORS["white"]
        )
        self.log_box.grid(row=7, column=0, sticky="nsew", padx=12, pady=(0, 12))

    def log(self, text: str):
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

            missing = validate_columns(df)
            if missing:
                messagebox.showerror(
                    "Colunas faltando",
                    "As seguintes colunas obrigatórias não foram encontradas:\n\n" + "\n".join(missing)
                )
                return

            df = clean_dataframe(df)

            self.df_original = df
            self.df_filtered = df.copy()
            self.refresh_table(self.df_filtered)

            self.status_var.set(f"Planilha carregada com {len(df)} registros.")
            self.log(f"[OK] Planilha importada: {path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao importar planilha:\n{e}")
            self.log(f"[ERRO] Importação falhou: {e}")

    def refresh_table(self, df: pd.DataFrame):
        for item in self.tree.get_children():
            self.tree.delete(item)

        preview_df = df.head(300)

        for _, row in preview_df.iterrows():
            self.tree.insert("", "end", values=(
                row.get("nome", ""),
                row.get("documento", ""),
                row.get("vencimento_fmt", ""),
                row.get("valor_fmt", ""),
                row.get("telefone", ""),
                "Sim" if row.get("telefone_valido", False) else "Não"
            ))

    def get_template_text(self):
        return self.template_box.get("1.0", "end").strip()

    def generate_messages_action(self):
        if self.df_original.empty:
            messagebox.showwarning("Aviso", "Importe uma planilha primeiro.")
            return

        try:
            mode = self.filter_mode_var.get()
            days = int(self.days_var.get() or "3")
            only_valid_phone = self.only_valid_phone_var.get()

            self.df_filtered = filter_dataframe(
                self.df_original,
                mode=mode,
                dias=days,
                only_valid_phone=only_valid_phone
            )

            self.refresh_table(self.df_filtered)

            template = self.get_template_text()
            self.generated_messages = generate_messages(self.df_filtered, template)

            self.preview_box.delete("1.0", "end")
            preview_texts = []
            for i, item in enumerate(self.generated_messages[:10], start=1):
                preview_texts.append(
                    f"{i}) {item['telefone']} - {item['nome']}\n{item['mensagem']}\n{'-' * 70}"
                )

            self.preview_box.insert("1.0", "\n".join(preview_texts) if preview_texts else "Nenhuma mensagem gerada.")
            self.total_var.set(f"{len(self.generated_messages)} mensagens geradas")
            self.status_var.set("Mensagens geradas com sucesso.")
            self.log(f"[OK] {len(self.generated_messages)} mensagens geradas.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar mensagens:\n{e}")
            self.log(f"[ERRO] Geração falhou: {e}")

    def start_whatsapp(self):
        driver_path = self.driver_path_var.get().strip()
        if not driver_path:
            messagebox.showwarning("Aviso", "Selecione o ChromeDriver.")
            return

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
                self.status_var.set("Erro ao iniciar WhatsApp.")
                self.log(f"[ERRO] WhatsApp: {e}")
                messagebox.showerror("Erro", f"Falha ao iniciar WhatsApp:\n{e}")

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
                skip_duplicates = self.skip_duplicates_var.get()

                self.progress.set(0)
                self.status_var.set("Enviando mensagens...")

                for idx, item in enumerate(self.generated_messages, start=1):
                    if self.stop_requested:
                        self.status_var.set("Envio interrompido pelo usuário.")
                        self.log("[PARADO] Envio encerrado manualmente.")
                        break

                    nome = item["nome"]
                    documento = item["documento"]
                    telefone = item["telefone"]
                    mensagem = item["mensagem"]
                    telefone_valido = item["telefone_valido"]

                    try:
                        if skip_invalid and not telefone_valido:
                            self.db.save_send(nome, documento, telefone, mensagem, "ignorado", "telefone inválido")
                            self.log(f"[IGNORADO] {telefone} - telefone inválido")
                            self.progress.set(idx / total)
                            continue

                        if skip_duplicates and self.db.was_already_sent(telefone, documento):
                            self.db.save_send(nome, documento, telefone, mensagem, "ignorado", "duplicado")
                            self.log(f"[IGNORADO] {telefone} - já enviado anteriormente")
                            self.progress.set(idx / total)
                            continue

                        self.sender.send_message(telefone, mensagem)
                        self.db.save_send(nome, documento, telefone, mensagem, "enviado", "")
                        sent += 1
                        self.log(f"[ENVIADO] {telefone} - {nome}")

                        self.progress.set(idx / total)
                        self.status_var.set(f"Enviando... {idx}/{total}")

                        # espera entre envios, mas respeitando parada manual
                        if idx < total:
                            for _ in range(interval):
                                if self.stop_requested:
                                    self.status_var.set("Envio interrompido pelo usuário.")
                                    self.log("[PARADO] Envio encerrado durante intervalo.")
                                    break
                                time.sleep(1)

                            if self.stop_requested:
                                break

                    except Exception as inner_error:
                        self.db.save_send(nome, documento, telefone, mensagem, "erro", str(inner_error))
                        self.log(f"[ERRO] {telefone} - {inner_error}")

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