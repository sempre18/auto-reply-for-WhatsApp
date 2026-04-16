import json
import os
import threading
import traceback
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

from database import HistoryDB
from humanizer import HumanBehaviorEngine, PROFILES
from template_manager import TemplateManager
from utils import (
    clean_dataframe_dynamic,
    extract_template_variables,
    filter_dataframe,
    generate_messages,
    validate_template_columns,
)
from whatsapp import WhatsAppError, WhatsAppSender

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

C = {
    "orange": "#F57C00",
    "orange_hover": "#D96C00",
    "black": "#0D0D0D",
    "graphite": "#1A1A1A",
    "panel": "#222222",
    "gray": "#2C2C2C",
    "input_bg": "#333333",
    "light": "#E0E0E0",
    "white": "#FFFFFF",
    "success": "#2E8B57",
    "success_h": "#236B43",
    "danger": "#C62828",
    "danger_h": "#9E1F1F",
    "warning": "#F9A825",
    "info": "#1565C0",
    "muted": "#888888",
}

SETTINGS_FILE = "app_settings.json"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GS Trator • Cobrança via WhatsApp Pro")
        self.geometry("1680x980")
        self.minsize(1280, 760)
        self.configure(fg_color=C["black"])

        self.db = HistoryDB()
        self.template_manager = TemplateManager()
        self.humanizer = HumanBehaviorEngine()
        self.sender: WhatsAppSender | None = None

        self.df_original = pd.DataFrame()
        self.df_filtered = pd.DataFrame()
        self.generated_messages: list[dict] = []
        self.current_mapping: dict[str, str] = {}
        self.current_required_vars: list[str] = []

        self.is_sending = False
        self.stop_requested = False
        self._session_start = ""
        self._session_id = ""
        self._sent_session = 0
        self._errors_session = 0
        self._skipped_session = 0
        self._simulated_session = 0

        self._build_variables()
        self._build_ui()
        self._configure_treeview_style()
        self._load_settings()
        self._refresh_template_list()
        self._on_profile_change(self.profile_var.get())
        self._update_metrics_display()

    def _build_variables(self):
        self.excel_path_var = tk.StringVar()
        self.filter_mode_var = tk.StringVar(value="todos")
        self.days_var = tk.StringVar(value="3")
        self.only_valid_phone_var = tk.BooleanVar(value=True)
        self.skip_invalid_var = tk.BooleanVar(value=True)
        self.auto_driver_var = tk.BooleanVar(value=True)
        self.driver_path_var = tk.StringVar()
        self.profile_var = tk.StringVar(value=list(PROFILES.keys())[0])
        self.selected_template_var = tk.StringVar(value="Aleatório")
        self.status_var = tk.StringVar(value="Pronto.")
        self.total_var = tk.StringVar(value="0 mensagens geradas")
        self.simulation_mode_var = tk.BooleanVar(value=True)

    def _configure_treeview_style(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background=C["panel"],
            foreground=C["white"],
            fieldbackground=C["panel"],
            rowheight=27,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=C["orange"],
            foreground=C["white"],
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", C["orange_hover"])],
            foreground=[("selected", C["white"])],
        )

    def _btn(self, parent, text, cmd, color=None, hover=None, **kw):
        return ctk.CTkButton(
            parent,
            text=text,
            command=cmd,
            fg_color=color or C["orange"],
            hover_color=hover or C["orange_hover"],
            text_color=C["white"],
            corner_radius=8,
            height=36,
            font=ctk.CTkFont(size=12, weight="bold"),
            **kw,
        )

    def _label(self, parent, text="", size=11, bold=False, color=None, **kw):
        return ctk.CTkLabel(
            parent,
            text=text,
            text_color=color or C["light"],
            font=ctk.CTkFont(size=size, weight="bold" if bold else "normal"),
            **kw,
        )

    def _entry(self, parent, var, width=None, **kw):
        params = {
            "textvariable": var,
            "fg_color": C["input_bg"],
            "border_color": C["orange"],
            "text_color": C["white"],
            **kw,
        }
        if width is not None:
            params["width"] = width
        return ctk.CTkEntry(parent, **params)

    def _frame(self, parent, **kw):
        return ctk.CTkFrame(parent, fg_color=C["panel"], corner_radius=10, **kw)

    def _section_label(self, parent, text):
        box = ctk.CTkFrame(parent, fg_color=C["orange"], corner_radius=6, height=28)
        box.pack(fill="x", pady=(10, 4))
        box.pack_propagate(False)
        self._label(box, text, size=11, bold=True, color=C["white"]).pack(side="left", padx=8)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_left_panel()
        self._build_right_panel()

    def _build_left_panel(self):
        left = ctk.CTkScrollableFrame(
            self,
            width=420,
            fg_color=C["graphite"],
            corner_radius=12,
            scrollbar_button_color=C["orange"],
            scrollbar_button_hover_color=C["orange_hover"],
        )
        left.grid(row=0, column=0, sticky="nsew", padx=(10, 6), pady=10)

        hdr = ctk.CTkFrame(left, fg_color=C["orange"], corner_radius=10, height=52)
        hdr.pack(fill="x", pady=(0, 8))
        hdr.pack_propagate(False)
        self._label(
            hdr,
            "⚙ GS Trator • Configurações",
            size=14,
            bold=True,
            color=C["white"],
        ).pack(side="left", padx=14)

        self._section_label(left, "📂 Planilha Excel")
        self._entry(left, self.excel_path_var).pack(fill="x", padx=8, pady=(2, 4))
        self._btn(left, "Importar planilha (.xlsx)", self.import_excel).pack(fill="x", padx=8, pady=2)

        self._section_label(left, "🌐 ChromeDriver")
        ctk.CTkCheckBox(
            left,
            text="Baixar automaticamente (recomendado)",
            variable=self.auto_driver_var,
            text_color=C["light"],
            fg_color=C["orange"],
            hover_color=C["orange_hover"],
            command=self._toggle_driver_entry,
        ).pack(anchor="w", padx=8, pady=(2, 4))

        self.driver_frame = ctk.CTkFrame(left, fg_color="transparent")
        self.driver_frame.pack(fill="x", padx=8)
        self._entry(self.driver_frame, self.driver_path_var).pack(fill="x", side="left", expand=True)
        self._btn(self.driver_frame, "…", self.select_chromedriver, width=36).pack(side="right", padx=(4, 0))
        self._toggle_driver_entry()

        self._section_label(left, "🔎 Filtros")
        self._label(left, "Modo de filtro").pack(anchor="w", padx=8)
        ctk.CTkOptionMenu(
            left,
            variable=self.filter_mode_var,
            values=["todos", "vencidos", "a_vencer"],
            fg_color=C["orange"],
            button_color=C["orange"],
            button_hover_color=C["orange_hover"],
            text_color=C["white"],
            dropdown_fg_color=C["graphite"],
            dropdown_text_color=C["white"],
        ).pack(fill="x", padx=8, pady=4)

        self._label(left, "Dias (para 'a vencer')").pack(anchor="w", padx=8)
        self._entry(left, self.days_var, width=100).pack(anchor="w", padx=8, pady=4)

        ctk.CTkCheckBox(
            left,
            text="Somente telefone válido",
            variable=self.only_valid_phone_var,
            text_color=C["light"],
            fg_color=C["orange"],
            hover_color=C["orange_hover"],
        ).pack(anchor="w", padx=8, pady=2)

        ctk.CTkCheckBox(
            left,
            text="Pular inválidos no envio",
            variable=self.skip_invalid_var,
            text_color=C["light"],
            fg_color=C["orange"],
            hover_color=C["orange_hover"],
        ).pack(anchor="w", padx=8, pady=2)

        ctk.CTkCheckBox(
            left,
            text="Modo simulação (não envia de verdade)",
            variable=self.simulation_mode_var,
            text_color=C["light"],
            fg_color=C["orange"],
            hover_color=C["orange_hover"],
        ).pack(anchor="w", padx=8, pady=2)

        self._section_label(left, "📝 Template ativo")
        self._label(left, "Selecionar template").pack(anchor="w", padx=8)
        self.template_combo = ctk.CTkOptionMenu(
            left,
            variable=self.selected_template_var,
            values=["Aleatório"],
            fg_color=C["orange"],
            button_color=C["orange"],
            button_hover_color=C["orange_hover"],
            text_color=C["white"],
            dropdown_fg_color=C["graphite"],
            dropdown_text_color=C["white"],
            command=self._on_template_change,
        )
        self.template_combo.pack(fill="x", padx=8, pady=4)
        self._btn(left, "Gerenciar templates", self._open_template_manager, color=C["info"], hover="#0D47A1").pack(
            fill="x", padx=8, pady=2
        )

        self._section_label(left, "🤖 Comportamento humano")
        self._label(left, "Perfil de envio").pack(anchor="w", padx=8)
        ctk.CTkOptionMenu(
            left,
            variable=self.profile_var,
            values=list(PROFILES.keys()),
            fg_color=C["orange"],
            button_color=C["orange"],
            button_hover_color=C["orange_hover"],
            text_color=C["white"],
            dropdown_fg_color=C["graphite"],
            dropdown_text_color=C["white"],
            command=self._on_profile_change,
        ).pack(fill="x", padx=8, pady=4)

        self.profile_info = self._label(left, "", size=10, color=C["muted"])
        self.profile_info.pack(anchor="w", padx=8, pady=(0, 4))

        self._section_label(left, "▶ Ações")
        self._btn(left, "Gerar mensagens", self.generate_messages_action).pack(fill="x", padx=8, pady=4)
        self._btn(left, "Iniciar WhatsApp / Login", self.start_whatsapp, color=C["info"], hover="#0D47A1").pack(
            fill="x", padx=8, pady=4
        )
        self._btn(left, "▶ Enviar mensagens", self.send_messages, color=C["success"], hover=C["success_h"]).pack(
            fill="x", padx=8, pady=4
        )
        self._btn(left, "⏹ Parar envio", self.stop_sending, color=C["danger"], hover=C["danger_h"]).pack(
            fill="x", padx=8, pady=4
        )
        self._btn(left, "📤 Exportar relatório da sessão", self.export_last_session_report).pack(
            fill="x", padx=8, pady=4
        )

        self._section_label(left, "📊 Progresso")
        self.progress = ctk.CTkProgressBar(left, progress_color=C["orange"], fg_color=C["gray"])
        self.progress.pack(fill="x", padx=8, pady=(4, 2))
        self.progress.set(0)

        self._label(left, "", textvariable=self.total_var, size=11, color=C["warning"]).pack(anchor="w", padx=8)
        ctk.CTkLabel(
            left,
            textvariable=self.status_var,
            wraplength=370,
            justify="left",
            text_color=C["white"],
            font=ctk.CTkFont(size=11),
        ).pack(anchor="w", padx=8, pady=(2, 8))

        self._section_label(left, "🧭 Mapeamento da planilha")
        self.mapping_box = ctk.CTkTextbox(
            left,
            height=150,
            fg_color=C["gray"],
            border_color=C["orange"],
            border_width=1,
            text_color=C["light"],
            font=ctk.CTkFont(family="Consolas", size=10),
        )
        self.mapping_box.pack(fill="x", padx=8, pady=(4, 8))

        self._section_label(left, "📈 Métricas gerais")
        self.metrics_box = ctk.CTkTextbox(
            left,
            height=110,
            fg_color=C["gray"],
            border_color=C["orange"],
            border_width=1,
            text_color=C["light"],
            font=ctk.CTkFont(family="Consolas", size=10),
        )
        self.metrics_box.pack(fill="x", padx=8, pady=(4, 8))

    def _build_right_panel(self):
        right = ctk.CTkFrame(self, fg_color=C["graphite"], corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=10)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=3)
        right.grid_rowconfigure(3, weight=1)
        right.grid_rowconfigure(5, weight=1)

        hdr = ctk.CTkFrame(right, fg_color="transparent", height=44)
        hdr.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
        self._label(hdr, "📋 Dados da planilha", size=14, bold=True, color=C["white"]).pack(side="left")

        table_outer = ctk.CTkFrame(right, fg_color=C["gray"], corner_radius=8)
        table_outer.grid(row=1, column=0, sticky="nsew", padx=12, pady=(4, 8))
        table_outer.grid_rowconfigure(0, weight=1)
        table_outer.grid_columnconfigure(0, weight=1)

        cols = ("nome", "documento", "vencimento", "valor", "telefone", "valido", "preparo", "template")
        self.tree = ttk.Treeview(table_outer, columns=cols, show="headings")
        headers = {
            "nome": "Cliente",
            "documento": "Documento",
            "vencimento": "Vencimento",
            "valor": "Valor",
            "telefone": "Telefone",
            "valido": "Válido",
            "preparo": "Preparo",
            "template": "Template",
        }
        widths = {
            "nome": 220,
            "documento": 120,
            "vencimento": 100,
            "valor": 110,
            "telefone": 145,
            "valido": 60,
            "preparo": 130,
            "template": 180,
        }

        for c in cols:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=widths[c], anchor="w")

        sb = ttk.Scrollbar(table_outer, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        self._label(right, "🔍 Preview da mensagem selecionada", size=13, bold=True, color=C["white"]).grid(
            row=2, column=0, sticky="w", padx=12, pady=(4, 2)
        )
        self.preview_box = ctk.CTkTextbox(
            right,
            height=180,
            fg_color=C["gray"],
            border_color=C["orange"],
            border_width=1,
            text_color=C["white"],
            font=ctk.CTkFont(family="Consolas", size=10),
        )
        self.preview_box.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 8))

        self._label(right, "🗒 Log de envios", size=13, bold=True, color=C["white"]).grid(
            row=4, column=0, sticky="w", padx=12, pady=(2, 2)
        )
        self.log_box = ctk.CTkTextbox(
            right,
            height=180,
            fg_color=C["gray"],
            border_color=C["orange"],
            border_width=1,
            text_color=C["white"],
            font=ctk.CTkFont(family="Consolas", size=10),
        )
        self.log_box.grid(row=5, column=0, sticky="nsew", padx=12, pady=(0, 12))

    def _toggle_driver_entry(self):
        state = "disabled" if self.auto_driver_var.get() else "normal"
        for w in self.driver_frame.winfo_children():
            try:
                w.configure(state=state)
            except Exception:
                pass

    def _on_profile_change(self, name: str):
        profile = PROFILES.get(name)
        if profile:
            self.humanizer = HumanBehaviorEngine(profile)
            self.humanizer.set_stop_flag(lambda: self.stop_requested)
            self.profile_info.configure(
                text=(
                    f"Delay: {profile.delay_min:.0f}-{profile.delay_max:.0f}s | "
                    f"Pausa a cada {profile.long_pause_every} envios\n"
                    f"Limite: {profile.max_per_hour}/hora | Aquecimento: x{profile.warmup_factor:.1f}"
                )
            )
        self._save_settings()

    def _refresh_template_list(self):
        names = self.template_manager.get_active_names()
        options = ["Aleatório"] + names
        self.template_combo.configure(values=options)
        if self.selected_template_var.get() not in options:
            self.selected_template_var.set("Aleatório")

    def _selected_template_id(self) -> str | None:
        sel_name = self.selected_template_var.get()
        if sel_name == "Aleatório":
            return None
        tpl = self.template_manager.get_template_by_name(sel_name)
        return tpl["id"] if tpl else None

    def _on_template_change(self, _selected_name=None):
        self._save_settings()
        if self.df_original.empty:
            return

        template_id = self._selected_template_id()
        required_vars = extract_template_variables(self.template_manager, template_id)
        if "telefone" not in required_vars:
            required_vars.append("telefone")

        missing, mapping = validate_template_columns(
            self.df_original,
            required_vars,
            template_manager=self.template_manager,
            template_id=template_id,
        )

        self.current_required_vars = required_vars
        self.current_mapping = mapping
        self._update_mapping_display(required_vars, mapping, missing)

        if missing:
            self.status_var.set("Template trocado, mas a planilha atual não atende todas as variáveis.")
            self.log(f"[ALERTA] Template exige colunas ausentes: {', '.join(missing)}")
        else:
            self.status_var.set("Template validado com a planilha atual.")
            self.log("[OK] Template validado com a planilha atual.")

    def _update_mapping_display(self, required_vars=None, mapping=None, missing=None):
        required_vars = required_vars or []
        mapping = mapping or {}
        missing = missing or []

        lines = [
            f"Variáveis exigidas: {', '.join(required_vars) if required_vars else '-'}",
            "",
            "Mapeamento encontrado:",
        ]
        for k, v in mapping.items():
            lines.append(f"  {k:<18} -> {v}")
        if not mapping:
            lines.append("  (nenhum)")

        lines.append("")
        lines.append("Faltando:")
        if missing:
            for item in missing:
                lines.append(f"  - {item}")
        else:
            lines.append("  (nada)")

        self.mapping_box.configure(state="normal")
        self.mapping_box.delete("1.0", "end")
        self.mapping_box.insert("1.0", "\n".join(lines))
        self.mapping_box.configure(state="disabled")

    def _update_metrics_display(self):
        m = self.db.get_metrics()
        total = m.get("total", 0) or 0
        enviados = m.get("enviados", 0) or 0
        erros = m.get("erros", 0) or 0
        ignorados = m.get("ignorados", 0) or 0
        simulados = m.get("simulados", 0) or 0
        avg_d = m.get("avg_delay") or 0
        taxa_ok = (enviados / total * 100) if total else 0

        txt = (
            f"Total histórico : {total}\n"
            f"Enviados         : {enviados} ({taxa_ok:.1f}%)\n"
            f"Erros            : {erros}\n"
            f"Ignorados        : {ignorados}\n"
            f"Simulados        : {simulados}\n"
            f"Delay médio      : {avg_d:.1f}s\n"
            f"Sessão atual     : {self._sent_session} enviados | {self._errors_session} erros"
        )
        self.metrics_box.configure(state="normal")
        self.metrics_box.delete("1.0", "end")
        self.metrics_box.insert("1.0", txt)
        self.metrics_box.configure(state="disabled")

    def log(self, text: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}] {text}\n")
        self.log_box.see("end")
        self.update_idletasks()

    def _save_settings(self):
        data = {
            "excel_path": self.excel_path_var.get(),
            "filter_mode": self.filter_mode_var.get(),
            "days": self.days_var.get(),
            "only_valid_phone": self.only_valid_phone_var.get(),
            "skip_invalid": self.skip_invalid_var.get(),
            "auto_driver": self.auto_driver_var.get(),
            "driver_path": self.driver_path_var.get(),
            "profile": self.profile_var.get(),
            "template_name": self.selected_template_var.get(),
            "simulation_mode": self.simulation_mode_var.get(),
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_settings(self):
        if not os.path.exists(SETTINGS_FILE):
            return
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.excel_path_var.set(data.get("excel_path", ""))
            self.filter_mode_var.set(data.get("filter_mode", "todos"))
            self.days_var.set(str(data.get("days", "3")))
            self.only_valid_phone_var.set(bool(data.get("only_valid_phone", True)))
            self.skip_invalid_var.set(bool(data.get("skip_invalid", True)))
            self.auto_driver_var.set(bool(data.get("auto_driver", True)))
            self.driver_path_var.set(data.get("driver_path", ""))
            self.profile_var.set(data.get("profile", self.profile_var.get()))
            self.selected_template_var.set(data.get("template_name", "Aleatório"))
            self.simulation_mode_var.set(bool(data.get("simulation_mode", True)))
            self._toggle_driver_entry()
        except Exception:
            pass

    def import_excel(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
        )
        if not path:
            return

        try:
            df_raw = pd.read_excel(path)
            template_id = self._selected_template_id()
            required_vars = extract_template_variables(self.template_manager, template_id)

            if "telefone" not in required_vars:
                required_vars.append("telefone")

            missing, mapping = validate_template_columns(
                df_raw,
                required_vars,
                template_manager=self.template_manager,
                template_id=template_id,
            )

            self.current_required_vars = required_vars
            self.current_mapping = mapping
            self._update_mapping_display(required_vars, mapping, missing)

            if missing:
                messagebox.showerror(
                    "Colunas faltando",
                    "A planilha não contém as colunas exigidas pelo template:\n\n" + "\n".join(missing),
                )
                return

            df = clean_dataframe_dynamic(df_raw, mapping)

            self.df_original = df
            self.df_filtered = df.copy()
            self.excel_path_var.set(path)
            self._refresh_table_preview_df(self.df_filtered)
            self.status_var.set(f"Planilha carregada: {len(df)} registros.")
            self.log(f"[OK] Planilha importada: {path}")
            self.log(f"[INFO] Variáveis exigidas: {', '.join(required_vars)}")
            self._save_settings()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao importar:\n{e}")
            self.log(f"[ERRO] Importação: {e}\n{traceback.format_exc()}")

    def select_chromedriver(self):
        path = filedialog.askopenfilename(
            title="Selecione o ChromeDriver",
            filetypes=[("Executável", "*.exe"), ("Todos", "*.*")],
        )
        if path:
            self.driver_path_var.set(path)
            self.log(f"[OK] ChromeDriver: {path}")
            self._save_settings()

    def _refresh_table_preview_df(self, df: pd.DataFrame):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in df.head(500).iterrows():
            self.tree.insert(
                "",
                "end",
                values=(
                    row.get("nome", row.get("cliente", "")),
                    row.get("documento", row.get("pedido", "")),
                    row.get("vencimento_fmt", row.get("vencimento", "")),
                    row.get("valor_fmt", row.get("valor", "")),
                    row.get("telefone", ""),
                    "✔" if row.get("telefone_valido") else "✘",
                    "—",
                    "—",
                ),
            )

    def _refresh_table_messages(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for msg in self.generated_messages[:500]:
            self.tree.insert(
                "",
                "end",
                values=(
                    msg.get("nome", ""),
                    msg.get("documento", ""),
                    msg.get("vencimento", ""),
                    msg.get("valor", ""),
                    msg.get("telefone", ""),
                    "✔" if msg.get("telefone_valido") else "✘",
                    msg.get("preparation_status", ""),
                    msg.get("template_id", ""),
                ),
            )

    def generate_messages_action(self):
        if self.df_original.empty:
            messagebox.showwarning("Aviso", "Importe uma planilha primeiro.")
            return

        try:
            mode = self.filter_mode_var.get()
            days = int(self.days_var.get() or "3")
            only_valid = self.only_valid_phone_var.get()

            self.df_filtered = filter_dataframe(self.df_original, mode, days, only_valid)

            template_id = self._selected_template_id()
            self.generated_messages = generate_messages(
                self.df_filtered,
                template_manager=self.template_manager,
                template_id=template_id,
            )

            self._refresh_table_messages()
            self._render_preview_list()

            n = len(self.generated_messages)
            ready = sum(1 for x in self.generated_messages if x["preparation_status"] == "pronto")
            self.total_var.set(f"{n} mensagens geradas")
            self.status_var.set(f"{n} mensagens prontas/avaliadas. {ready} com status 'pronto'.")
            self.log(f"[OK] {n} mensagens geradas.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar:\n{e}")
            self.log(f"[ERRO] Geração: {e}\n{traceback.format_exc()}")

    def _render_preview_list(self):
        self.preview_box.delete("1.0", "end")
        lines = []
        for i, item in enumerate(self.generated_messages[:8], 1):
            lines.append(
                f"{'─'*70}\n"
                f"#{i} | {item['telefone']} | {item['nome']}\n"
                f"template: {item['template_id']} | preparo: {item['preparation_status']}\n"
                f"faltando: {', '.join(item['missing_fields']) if item['missing_fields'] else '-'}\n"
                f"placeholders: {', '.join(item['placeholders_left']) if item['placeholders_left'] else '-'}\n\n"
                f"{item['mensagem']}\n"
            )
        self.preview_box.insert("1.0", "\n".join(lines) if lines else "Nenhuma mensagem.")

    def _on_tree_select(self, _event=None):
        selected = self.tree.selection()
        if not selected:
            return

        idx = self.tree.index(selected[0])
        if idx >= len(self.generated_messages):
            return

        item = self.generated_messages[idx]
        text = (
            f"Cliente: {item.get('nome', '')}\n"
            f"Telefone: {item.get('telefone', '')}\n"
            f"Template: {item.get('template_id', '')}\n"
            f"Preparo: {item.get('preparation_status', '')}\n"
            f"Campos faltando: {', '.join(item.get('missing_fields', [])) or '-'}\n"
            f"Placeholders restantes: {', '.join(item.get('placeholders_left', [])) or '-'}\n"
            f"{'-'*70}\n"
            f"{item.get('mensagem', '')}\n"
        )
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", text)

    def start_whatsapp(self):
        if self.simulation_mode_var.get():
            self.log("[INFO] Modo simulação ativo: não é necessário iniciar o WhatsApp.")
            self.status_var.set("Modo simulação ativo.")
            return

        def worker():
            try:
                auto = self.auto_driver_var.get()
                path = self.driver_path_var.get().strip() if not auto else None

                self.status_var.set("Iniciando navegador...")
                self.log("[INFO] Iniciando Chrome...")
                self.sender = WhatsAppSender(chromedriver_path=path)
                self.sender.start()
                self.sender.open_whatsapp()
                self.status_var.set("Aguardando QR Code...")
                self.log("[INFO] Escaneie o QR Code no WhatsApp Web.")
                self.sender.wait_for_login(log_fn=self.log)
                self.status_var.set("✔ WhatsApp conectado!")
                self.log("[OK] WhatsApp Web conectado.")
                self._update_metrics_display()
            except Exception as e:
                self.status_var.set("Erro ao iniciar WhatsApp.")
                self.log(f"[ERRO] {e}")
                messagebox.showerror("Erro", str(e))

        threading.Thread(target=worker, daemon=True).start()

    def stop_sending(self):
        if not self.is_sending:
            self.log("[INFO] Nenhum envio em andamento.")
            return
        self.stop_requested = True
        self.status_var.set("⏹ Parada solicitada...")
        self.log("[INFO] Parada solicitada.")

    def send_messages(self):
        if self.is_sending:
            messagebox.showinfo("Aviso", "Envio já em andamento.")
            return
        if not self.generated_messages:
            messagebox.showwarning("Aviso", "Gere as mensagens primeiro.")
            return
        if not self.simulation_mode_var.get() and not self.sender:
            messagebox.showwarning("Aviso", "Inicie o WhatsApp primeiro.")
            return

        self.is_sending = True
        self.stop_requested = False
        self._sent_session = 0
        self._errors_session = 0
        self._skipped_session = 0
        self._simulated_session = 0
        self._session_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._session_id = datetime.now().strftime("%Y%m%d_%H%M%S")

        self.humanizer.set_stop_flag(lambda: self.stop_requested)

        def worker():
            total = len(self.generated_messages)
            try:
                self.progress.set(0)
                self.status_var.set("Processando...")

                for idx, item in enumerate(self.generated_messages, 1):
                    if self.stop_requested:
                        self.log("[⏹] Envio interrompido.")
                        break

                    nome = item["nome"]
                    documento = item["documento"]
                    telefone = item["telefone"]
                    mensagem = item["mensagem"]
                    tid = item["template_id"]
                    valido = item["telefone_valido"]
                    prep = item["preparation_status"]
                    placeholders_left = ", ".join(item.get("placeholders_left", []))
                    row_json = item.get("row_json", "")

                    if prep != "pronto":
                        self.db.save_send(
                            nome=nome,
                            documento=documento,
                            telefone=telefone,
                            mensagem=mensagem,
                            status="ignorado",
                            error=f"preparo={prep}",
                            template_id=tid,
                            session_id=self._session_id,
                            preparation_status=prep,
                            placeholders_left=placeholders_left,
                            context_json=row_json,
                            simulation=self.simulation_mode_var.get(),
                        )
                        self._skipped_session += 1
                        self.log(f"[IGNORADO] {telefone} — preparo={prep}")
                        self.progress.set(idx / total)
                        continue

                    if self.skip_invalid_var.get() and not valido:
                        self.db.save_send(
                            nome=nome,
                            documento=documento,
                            telefone=telefone,
                            mensagem=mensagem,
                            status="ignorado",
                            error="telefone inválido",
                            template_id=tid,
                            session_id=self._session_id,
                            preparation_status=prep,
                            placeholders_left=placeholders_left,
                            context_json=row_json,
                            simulation=self.simulation_mode_var.get(),
                        )
                        self._skipped_session += 1
                        self.log(f"[IGNORADO] {telefone} — telefone inválido")
                        self.progress.set(idx / total)
                        continue

                    if self.simulation_mode_var.get():
                        self.db.save_send(
                            nome=nome,
                            documento=documento,
                            telefone=telefone,
                            mensagem=mensagem,
                            status="simulado",
                            error="modo simulação",
                            template_id=tid,
                            session_id=self._session_id,
                            preparation_status=prep,
                            placeholders_left=placeholders_left,
                            context_json=row_json,
                            simulation=True,
                        )
                        self._simulated_session += 1
                        self.log(f"[SIMULADO] ({idx}/{total}) {telefone} — {nome} [{tid}]")
                        self.progress.set(idx / total)
                        continue

                    if not self.humanizer.wait_for_hourly_reset(log_fn=self.log):
                        break

                    if idx > 1:
                        if self.humanizer.should_long_pause():
                            pause = self.humanizer.compute_long_pause()
                            self.log(f"[⏸] Pausa longa ({pause:.0f}s)...")
                            if not self.humanizer.wait(pause):
                                break

                        delay = self.humanizer.compute_delay()
                        self.humanizer.record_delay(delay)
                        self.status_var.set(f"⏳ Aguardando {delay:.0f}s antes de enviar #{idx}/{total}...")
                        if not self.humanizer.wait(delay):
                            break
                    else:
                        delay = 0.0

                    typing = self.humanizer.get_typing_delay()

                    try:
                        self.sender.send_message(telefone, mensagem, typing_delay=typing)
                        self.db.save_send(
                            nome=nome,
                            documento=documento,
                            telefone=telefone,
                            mensagem=mensagem,
                            status="enviado",
                            error="",
                            template_id=tid,
                            delay_used=delay,
                            session_id=self._session_id,
                            typing_delay=typing,
                            preparation_status=prep,
                            placeholders_left=placeholders_left,
                            context_json=row_json,
                            simulation=False,
                        )
                        self.humanizer.register_send()
                        self._sent_session += 1
                        self.log(f"[✔ ENVIADO] ({idx}/{total}) {telefone} — {nome} [{tid}]")
                    except WhatsAppError as e:
                        self.db.save_send(
                            nome=nome,
                            documento=documento,
                            telefone=telefone,
                            mensagem=mensagem,
                            status="erro",
                            error=str(e),
                            template_id=tid,
                            delay_used=delay,
                            session_id=self._session_id,
                            typing_delay=typing,
                            preparation_status=prep,
                            placeholders_left=placeholders_left,
                            context_json=row_json,
                            simulation=False,
                        )
                        self._errors_session += 1
                        self.log(f"[✘ ERRO] {telefone} — {e}")

                    self.progress.set(idx / total)
                    self.status_var.set(
                        f"Processando {idx}/{total} — "
                        f"✔{self._sent_session} ✘{self._errors_session} "
                        f"⏭{self._skipped_session} 🧪{self._simulated_session}"
                    )
                    self._update_metrics_display()

                if not self.stop_requested:
                    self.status_var.set(
                        f"✔ Concluído — enviados: {self._sent_session} | erros: {self._errors_session} | "
                        f"ignorados: {self._skipped_session} | simulados: {self._simulated_session}"
                    )
                    self.log("[FIM] Sessão concluída.")

                avg_delay = self.humanizer.avg_delay() if hasattr(self.humanizer, "avg_delay") else 0.0
                self.db.save_session(
                    session_id=self._session_id,
                    session_start=self._session_start,
                    total_sent=self._sent_session,
                    total_errors=self._errors_session,
                    total_skipped=self._skipped_session,
                    total_simulated=self._simulated_session,
                    avg_delay=avg_delay,
                )
                self._update_metrics_display()
            except Exception as e:
                self.status_var.set("Erro no envio.")
                self.log(f"[ERRO GERAL] {e}\n{traceback.format_exc()}")
            finally:
                self.is_sending = False
                self.stop_requested = False

        threading.Thread(target=worker, daemon=True).start()

    def export_last_session_report(self):
        if not self._session_id:
            messagebox.showwarning("Aviso", "Ainda não existe sessão para exportar.")
            return

        rows = self.db.get_session_rows(self._session_id)
        if not rows:
            messagebox.showwarning("Aviso", "Nenhum dado encontrado para a última sessão.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Salvar relatório",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"relatorio_sessao_{self._session_id}.xlsx",
        )
        if not save_path:
            return

        try:
            pd.DataFrame(rows).to_excel(save_path, index=False)
            self.log(f"[OK] Relatório exportado: {save_path}")
            messagebox.showinfo("Sucesso", "Relatório exportado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar relatório:\n{e}")

    def _open_template_manager(self):
        TemplateManagerWindow(self, self.template_manager, on_close=self._on_templates_updated)

    def _on_templates_updated(self):
        self.template_manager.load()
        self._refresh_template_list()
        self._on_template_change()


class TemplateManagerWindow(ctk.CTkToplevel):
    def __init__(self, parent, manager: TemplateManager, on_close=None):
        super().__init__(parent)
        self.title("Gerenciador de Templates")
        self.geometry("960x720")
        self.configure(fg_color=C["black"])
        self.grab_set()

        self.manager = manager
        self._on_close = on_close
        self._selected_id: str | None = None

        self._build_ui()
        self._load_list()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(self, fg_color=C["graphite"], corner_radius=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        left.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(left, text="Templates", text_color=C["white"], font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 4)
        )

        self.list_box = ctk.CTkScrollableFrame(left, fg_color=C["panel"], scrollbar_button_color=C["orange"])
        self.list_box.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

        btn_row = ctk.CTkFrame(left, fg_color="transparent")
        btn_row.grid(row=2, column=0, padx=8, pady=(0, 8), sticky="ew")
        ctk.CTkButton(
            btn_row,
            text="+ Novo",
            command=self._new_template,
            fg_color=C["success"],
            hover_color=C["success_h"],
            text_color=C["white"],
            height=32,
        ).pack(side="left", expand=True, fill="x", padx=(0, 4))
        ctk.CTkButton(
            btn_row,
            text="Excluir",
            command=self._delete_template,
            fg_color=C["danger"],
            hover_color=C["danger_h"],
            text_color=C["white"],
            height=32,
        ).pack(side="left", expand=True, fill="x")

        right = ctk.CTkFrame(self, fg_color=C["graphite"], corner_radius=10)
        right.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(7, weight=1)

        ctk.CTkLabel(right, text="Editor de Template", text_color=C["white"], font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 4)
        )

        self.f_id = ctk.CTkEntry(right, fg_color=C["input_bg"], border_color=C["orange"], text_color=C["white"])
        self.f_id.grid(row=1, column=0, sticky="ew", padx=10, pady=4)
        self.f_id.insert(0, "id_template")

        self.f_name = ctk.CTkEntry(right, fg_color=C["input_bg"], border_color=C["orange"], text_color=C["white"])
        self.f_name.grid(row=2, column=0, sticky="ew", padx=10, pady=4)
        self.f_name.insert(0, "Nome do template")

        self.f_cat = ctk.CTkEntry(right, fg_color=C["input_bg"], border_color=C["orange"], text_color=C["white"])
        self.f_cat.grid(row=3, column=0, sticky="ew", padx=10, pady=4)
        self.f_cat.insert(0, "cobranca")

        self.active_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            right,
            text="Ativo",
            variable=self.active_var,
            text_color=C["white"],
            fg_color=C["orange"],
        ).grid(row=4, column=0, sticky="w", padx=10, pady=4)

        ctk.CTkLabel(right, text="Aliases por variável (JSON)", text_color=C["light"]).grid(
            row=5, column=0, sticky="w", padx=10, pady=(8, 2)
        )
        self.aliases_box = ctk.CTkTextbox(
            right,
            height=120,
            fg_color=C["gray"],
            border_color=C["orange"],
            border_width=1,
            text_color=C["white"],
        )
        self.aliases_box.grid(row=6, column=0, sticky="ew", padx=10, pady=4)

        ctk.CTkLabel(right, text="Variantes (1 por linha separada por ---)", text_color=C["light"]).grid(
            row=7, column=0, sticky="nw", padx=10, pady=(8, 2)
        )
        self.variants_box = ctk.CTkTextbox(
            right,
            fg_color=C["gray"],
            border_color=C["orange"],
            border_width=1,
            text_color=C["white"],
        )
        self.variants_box.grid(row=8, column=0, sticky="nsew", padx=10, pady=4)

        actions = ctk.CTkFrame(right, fg_color="transparent")
        actions.grid(row=9, column=0, sticky="ew", padx=10, pady=8)

        ctk.CTkButton(actions, text="Salvar", command=self._save_template, fg_color=C["success"], hover_color=C["success_h"]).pack(
            side="left", fill="x", expand=True, padx=(0, 4)
        )
        ctk.CTkButton(actions, text="Alternar ativo", command=self._toggle_active, fg_color=C["info"], hover_color="#0D47A1").pack(
            side="left", fill="x", expand=True, padx=(0, 4)
        )
        ctk.CTkButton(actions, text="Fechar", command=self._close, fg_color=C["gray"], hover_color=C["input_bg"]).pack(
            side="left", fill="x", expand=True
        )

    def _load_list(self):
        for w in self.list_box.winfo_children():
            w.destroy()

        for tpl in self.manager.get_all_templates():
            txt = f"{tpl['name']} {'[ativo]' if tpl.get('active', True) else '[inativo]'}"
            btn = ctk.CTkButton(
                self.list_box,
                text=txt,
                fg_color=C["panel"],
                hover_color=C["orange_hover"],
                anchor="w",
                command=lambda tid=tpl["id"]: self._load_template(tid),
            )
            btn.pack(fill="x", padx=4, pady=2)

    def _load_template(self, tid: str):
        tpl = self.manager.get_template_by_id(tid)
        if not tpl:
            return

        self._selected_id = tid
        self.f_id.delete(0, "end")
        self.f_id.insert(0, tpl.get("id", ""))

        self.f_name.delete(0, "end")
        self.f_name.insert(0, tpl.get("name", ""))

        self.f_cat.delete(0, "end")
        self.f_cat.insert(0, tpl.get("category", ""))

        self.active_var.set(bool(tpl.get("active", True)))

        self.aliases_box.delete("1.0", "end")
        self.aliases_box.insert("1.0", json.dumps(tpl.get("aliases", {}), ensure_ascii=False, indent=2))

        self.variants_box.delete("1.0", "end")
        self.variants_box.insert("1.0", "\n---\n".join(tpl.get("variants", [])))

    def _new_template(self):
        self._selected_id = None
        self.f_id.delete(0, "end")
        self.f_name.delete(0, "end")
        self.f_cat.delete(0, "end")
        self.aliases_box.delete("1.0", "end")
        self.variants_box.delete("1.0", "end")
        self.active_var.set(True)

    def _save_template(self):
        try:
            aliases_text = self.aliases_box.get("1.0", "end").strip() or "{}"
            aliases = json.loads(aliases_text)

            variants_raw = self.variants_box.get("1.0", "end").strip()
            variants = [x.strip() for x in variants_raw.split("\n---\n") if x.strip()]

            tpl = {
                "id": self.f_id.get().strip(),
                "name": self.f_name.get().strip(),
                "category": self.f_cat.get().strip() or "cobranca",
                "tone": "informal",
                "active": self.active_var.get(),
                "weight": 1,
                "aliases": aliases,
                "variants": variants,
            }

            if self._selected_id and self.manager.get_template_by_id(self._selected_id):
                self.manager.update_template(self._selected_id, tpl)
            else:
                self.manager.add_template(tpl)

            self._load_list()
            messagebox.showinfo("Sucesso", "Template salvo com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar template:\n{e}")

    def _delete_template(self):
        if not self._selected_id:
            messagebox.showwarning("Aviso", "Selecione um template.")
            return
        if not messagebox.askyesno("Confirmar", "Deseja excluir o template selecionado?"):
            return
        if self.manager.delete_template(self._selected_id):
            self._selected_id = None
            self._load_list()
            self._new_template()

    def _toggle_active(self):
        if not self._selected_id:
            messagebox.showwarning("Aviso", "Selecione um template.")
            return
        self.manager.toggle_active(self._selected_id)
        self._load_list()

    def _close(self):
        if self._on_close:
            self._on_close()
        self.destroy()