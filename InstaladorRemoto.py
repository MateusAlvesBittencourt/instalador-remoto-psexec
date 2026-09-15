import ctypes
import csv
import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_NAME = "Instalador Remoto via PsExec"
DEFAULT_MSI_ARGS = "/qn /norestart"
KNOWN_MSI_ARGS = {
    "/qn /norestart",
    "/norestart /qn",
    "/quiet /norestart",
    "/norestart /quiet",
}
DEFAULT_PSEXEC = r"C:\Windows\System32\PsExec64.exe"
SOFTWARE_UNC_ROOT = r"\\POA01FLS05\Software$"
CREATE_NO_WINDOW = 0x08000000
STATUS_PATTERN = re.compile(r"^\[(?P<time>[^]]+)] \[(?P<computer>W\d+)] (?P<stage>[^:]+): (?P<result>[^-]+?)(?: - (?P<details>.*))?$")
VALIDATION_LABELS = {
    "Código de saída": "ExitCode",
    "Arquivo instalado": "File",
    "Serviço do Windows": "Service",
}


def normalized_args(value: str) -> str:
    """Normaliza espaços e maiúsculas para reconhecer parâmetros padrão de MSI."""
    return " ".join(value.lower().split())


def has_known_msi_args(value: str) -> bool:
    return normalized_args(value) in KNOWN_MSI_ARGS


def build_table_report(headings: list[str], rows: list[list[str]]) -> str:
    """Gera texto tabulado, pronto para colar no Excel ou em uma mensagem."""
    lines = ["\t".join(headings)]
    lines.extend("\t".join(str(value) for value in row) for row in rows)
    return "\r\n".join(lines)

COLORS = {
    "navy": "#102A43",
    "navy_light": "#1D4E78",
    "blue": "#1769E0",
    "blue_hover": "#0E56C2",
    "background": "#F2F5F9",
    "card": "#FFFFFF",
    "border": "#DCE3EC",
    "text": "#172B4D",
    "muted": "#61758A",
    "green": "#14804A",
    "green_soft": "#E7F6ED",
    "amber": "#A15C00",
    "red": "#C9372C",
    "console": "#0B1623",
    "console_text": "#D9E2EC",
}


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> bool:
    if getattr(sys, "frozen", False):
        executable = sys.executable
        parameters = subprocess.list2cmdline(sys.argv[1:])
    else:
        executable = sys.executable
        parameters = subprocess.list2cmdline([str(Path(__file__).resolve()), *sys.argv[1:]])
    result = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", executable, parameters, str(Path.cwd()), 1
    )
    return result > 32


class InstallerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.process = None
        self.running = False
        self.output_queue = queue.Queue()
        self.machine_rows = {}
        self.machine_overall = {}
        self.active_list_path = None
        self.settings_path = Path(os.environ.get("APPDATA", Path.home())) / "InstaladorRemoto" / "config.json"

        self.source = tk.StringVar()
        self.installer = tk.StringVar()
        self.silent_args = tk.StringVar()
        self.computer_list = tk.StringVar()
        self.psexec = tk.StringVar(value=DEFAULT_PSEXEC)
        self.destination = tk.StringVar(value="PacoteInstalacao")
        self.throttle = tk.IntVar(value=4)
        self.validation_type = tk.StringVar(value="Código de saída")
        self.validation_value = tk.StringVar()
        self.status = tk.StringVar(value="Pronto")

        self.configure_window()
        self.build_interface()
        self.load_settings()
        self.on_validation_changed()
        self.root.bind("<Configure>", self.on_window_resize)
        self.root.after_idle(lambda: self.apply_responsive_layout(self.root.winfo_width()))
        self.root.after(100, self.consume_output)

    def configure_window(self) -> None:
        self.root.title(f"{APP_NAME} - Python x64")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = min(1240, max(820, int(screen_width * 0.88)))
        height = min(900, max(620, int(screen_height * 0.84)))
        position_x = max(0, (screen_width - width) // 2)
        position_y = max(0, (screen_height - height) // 2)
        self.root.geometry(f"{width}x{height}+{position_x}+{position_y}")
        self.root.minsize(760, 600)
        self.root.configure(bg=COLORS["background"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=COLORS["card"])
        style.configure("TLabel", background=COLORS["card"], foreground=COLORS["text"], font=("Segoe UI", 9))
        style.configure("Field.TLabel", font=("Segoe UI", 9, "bold"), foreground=COLORS["text"])
        style.configure("Hint.TLabel", font=("Segoe UI", 8), foreground=COLORS["muted"])
        style.configure("TEntry", padding=8, fieldbackground="#FBFCFE", bordercolor=COLORS["border"], lightcolor=COLORS["border"], darkcolor=COLORS["border"])
        style.map("TEntry", bordercolor=[("focus", COLORS["blue"])], lightcolor=[("focus", COLORS["blue"])])
        style.configure("Browse.TButton", font=("Segoe UI", 9), padding=(12, 7), background="#E9EFF6", foreground=COLORS["text"], borderwidth=0)
        style.map("Browse.TButton", background=[("active", "#DDE7F1")])
        style.configure("Secondary.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 10), background="#E8F0FA", foreground=COLORS["navy_light"], borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#D8E7F8"), ("disabled", "#EEF1F5")])
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 10), background=COLORS["blue"], foreground="#FFFFFF", borderwidth=0)
        style.map("Primary.TButton", background=[("active", COLORS["blue_hover"]), ("disabled", "#9AB9E7")], foreground=[("disabled", "#F6F8FB")])
        style.configure("Report.TButton", font=("Segoe UI", 9, "bold"), padding=(14, 7), background=COLORS["blue"], foreground="#FFFFFF", borderwidth=1, relief="solid")
        style.map("Report.TButton", background=[("active", COLORS["blue_hover"]), ("pressed", COLORS["navy_light"])], foreground=[("active", "#FFFFFF"), ("pressed", "#FFFFFF")])
        style.configure("Ghost.TButton", font=("Segoe UI", 8), padding=(8, 4), background=COLORS["card"], foreground=COLORS["muted"], borderwidth=0)
        style.map("Ghost.TButton", background=[("active", "#EEF2F7")])
        style.configure("Horizontal.TProgressbar", troughcolor="#DCE5EF", background=COLORS["blue"], borderwidth=0)
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=28, background="#FFFFFF", fieldbackground="#FFFFFF", foreground=COLORS["text"], borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E8EEF5", foreground=COLORS["navy"], relief="flat", padding=(6, 7))
        style.map("Treeview", background=[("selected", "#D9E9FF")], foreground=[("selected", COLORS["navy"])])

    def build_interface(self) -> None:
        header = tk.Frame(self.root, bg=COLORS["navy"], height=104)
        header.pack(fill="x")
        header.pack_propagate(False)
        header_inner = tk.Frame(header, bg=COLORS["navy"])
        header_inner.pack(fill="both", expand=True, padx=30, pady=18)
        tk.Label(header_inner, text="INSTALADOR REMOTO", bg=COLORS["navy"], fg="#AFC7DE", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(header_inner, text="Implantação de software via PsExec", bg=COLORS["navy"], fg="#FFFFFF", font=("Segoe UI", 20, "bold")).pack(anchor="w", pady=(2, 0))
        self.header_badge = tk.Label(header_inner, text="WINDOWS  •  X64  •  ADMINISTRADOR", bg=COLORS["navy_light"], fg="#DCEBFA", font=("Segoe UI", 8, "bold"), padx=12, pady=6)
        self.header_badge.place(relx=1.0, rely=0.5, anchor="e")

        footer = tk.Frame(self.root, bg="#E7EDF4", height=44)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        self.status_label = tk.Label(footer, textvariable=self.status, bg=COLORS["green_soft"], fg=COLORS["green"], font=("Segoe UI", 9, "bold"), padx=12, pady=5)
        self.status_label.pack(side="left", padx=(24, 12), pady=8)
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=180)
        self.progress.pack(side="left", pady=12)
        self.footer_info = tk.Label(footer, text="Execução remota como SYSTEM", bg="#E7EDF4", fg=COLORS["muted"], font=("Segoe UI", 8))
        self.footer_info.pack(side="right", padx=24)

        body = tk.Frame(self.root, bg=COLORS["background"])
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.content_canvas = tk.Canvas(body, bg=COLORS["background"], highlightthickness=0)
        content_scroll = ttk.Scrollbar(body, orient="vertical", command=self.content_canvas.yview)
        self.content_canvas.configure(yscrollcommand=content_scroll.set)
        self.content_canvas.grid(row=0, column=0, sticky="nsew")
        content_scroll.grid(row=0, column=1, sticky="ns")

        outer = tk.Frame(self.content_canvas, bg=COLORS["background"])
        self.content_window = self.content_canvas.create_window((0, 0), window=outer, anchor="nw")
        outer.bind("<Configure>", self.on_content_resize)
        self.content_canvas.bind("<Configure>", self.on_canvas_resize)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        form_card = self.create_card(outer)
        form_card.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 0))
        form_card.columnconfigure(0, weight=1)
        form = ttk.Frame(form_card, padding=(20, 16, 20, 18))
        form.grid(row=0, column=0, sticky="ew")
        self.form = form
        form.columnconfigure(0, weight=1, uniform="section")
        form.columnconfigure(2, weight=1, uniform="section")

        package_form = ttk.Frame(form)
        self.package_form = package_form
        package_form.grid(row=0, column=0, sticky="nsew")
        package_form.columnconfigure(1, weight=1)
        self.add_section_title(package_form, 0, "1", "Pacote de instalação", "Origem, instalador e execução silenciosa.")
        self.add_path_row(package_form, 1, "Pasta do pacote", self.source, self.choose_source, "Inclua todos os arquivos necessários")
        self.add_path_row(package_form, 2, "Instalador", self.installer, self.choose_installer, "Arquivos .exe e .msi")
        self.add_entry_row(package_form, 3, "Parâmetros", self.silent_args, "MSI: /qn /norestart")
        self.add_validation_rows(package_form, 4)

        self.form_separator = ttk.Separator(form, orient="vertical")
        self.form_separator.grid(row=0, column=1, sticky="ns", padx=18)

        remote_form = ttk.Frame(form)
        self.remote_form = remote_form
        remote_form.grid(row=0, column=2, sticky="nsew")
        remote_form.columnconfigure(1, weight=1)
        self.add_section_title(remote_form, 0, "2", "Destino e computadores", "Lista, PsExec e pasta remota.")
        self.add_path_row(remote_form, 1, "Patrimônios", self.computer_list, self.choose_list, "TXT ou CSV; formato W1234567")
        self.add_path_row(remote_form, 2, "PsExec", self.psexec, self.choose_psexec, "C:\\Windows\\System32\\PsExec64.exe")
        self.add_entry_row(remote_form, 3, "Destino", self.destination, "Dentro de C:\\Temp")
        self.add_throttle_row(remote_form, 4)

        actions_card = self.create_card(outer)
        actions_card.grid(row=1, column=0, sticky="ew", padx=24, pady=(14, 0))
        actions = ttk.Frame(actions_card, padding=(18, 14))
        self.actions = actions
        actions.pack(fill="x")
        actions.columnconfigure(1, weight=1)
        self.action_badge = tk.Label(actions, text="3", bg=COLORS["green_soft"], fg=COLORS["green"], font=("Segoe UI", 10, "bold"), width=3, height=1)
        self.action_badge.grid(row=0, column=0, rowspan=2, padx=(0, 12))
        self.action_title = ttk.Label(actions, text="Validar e executar", style="Field.TLabel")
        self.action_title.grid(row=0, column=1, sticky="w")
        self.action_subtitle = ttk.Label(actions, text="Teste o acesso antes de iniciar a implantação.", style="Hint.TLabel")
        self.action_subtitle.grid(row=1, column=1, sticky="w")
        self.test_button = ttk.Button(actions, text="Testar acesso", style="Secondary.TButton", command=lambda: self.start("Test"))
        self.test_button.grid(row=0, column=2, rowspan=2, padx=(18, 8), sticky="ew")
        self.retry_button = ttk.Button(actions, text="Reexecutar falhas", style="Secondary.TButton", command=self.retry_failures, state="disabled")
        self.retry_button.grid(row=0, column=3, rowspan=2, padx=(0, 8), sticky="ew")
        self.run_button = ttk.Button(actions, text="Copiar e instalar  ›", style="Primary.TButton", command=lambda: self.start("Install"))
        self.run_button.grid(row=0, column=4, rowspan=2, sticky="ew")

        output_card = self.create_card(outer)
        output_card.grid(row=2, column=0, sticky="nsew", padx=24, pady=(14, 18))
        output_card.columnconfigure(0, weight=1)
        output_card.rowconfigure(2, weight=1)
        output_header = ttk.Frame(output_card, padding=(16, 11, 12, 8))
        output_header.grid(row=0, column=0, sticky="ew")
        output_header.columnconfigure(0, weight=1)
        ttk.Label(output_header, text="Andamento da operação", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(output_header, text="COPIAR RELATÓRIO", style="Report.TButton", command=self.copy_report).grid(row=0, column=1, padx=(0, 10), sticky="e")
        ttk.Button(output_header, text="Limpar log", style="Ghost.TButton", command=self.clear_output).grid(row=0, column=2, sticky="e")

        table_frame = ttk.Frame(output_card, padding=(10, 0, 10, 9))
        table_frame.grid(row=1, column=0, sticky="ew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        columns = ("computer", "access", "copy", "install", "validation", "cleanup", "duration", "details")
        self.machine_table = ttk.Treeview(table_frame, columns=columns, show="headings", height=7)
        headings = {
            "computer": "Patrimônio", "access": "Acesso", "copy": "Cópia", "install": "Instalação",
            "validation": "Validação", "cleanup": "Limpeza", "duration": "Tempo", "details": "Detalhes",
        }
        widths = {"computer": 95, "access": 80, "copy": 95, "install": 120, "validation": 110, "cleanup": 90, "duration": 70, "details": 260}
        for column in columns:
            self.machine_table.heading(column, text=headings[column])
            self.machine_table.column(column, width=widths[column], minwidth=60, anchor="w", stretch=column == "details")
        self.machine_table.tag_configure("success", background="#ECF8F1")
        self.machine_table.tag_configure("failure", background="#FFF0EF")
        self.machine_table.tag_configure("running", background="#EDF5FF")
        table_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.machine_table.yview)
        table_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.machine_table.xview)
        self.machine_table.configure(yscrollcommand=table_scroll.set, xscrollcommand=table_scroll_x.set)
        self.machine_table.grid(row=0, column=0, sticky="nsew")
        table_scroll.grid(row=0, column=1, sticky="ns")
        table_scroll_x.grid(row=1, column=0, sticky="ew")

        console = tk.Frame(output_card, bg=COLORS["console"])
        console.grid(row=2, column=0, sticky="nsew", padx=1, pady=(0, 1))
        console.columnconfigure(0, weight=1)
        console.rowconfigure(0, weight=1)
        self.output = tk.Text(console, wrap="word", height=7, state="disabled", font=("Cascadia Mono", 9), bg=COLORS["console"], fg=COLORS["console_text"], insertbackground="#FFFFFF", relief="flat", padx=14, pady=12, spacing1=2, spacing3=2)
        scroll = ttk.Scrollbar(console, orient="vertical", command=self.output.yview)
        self.output.configure(yscrollcommand=scroll.set)
        self.output.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.output.tag_configure("success", foreground="#61D095")
        self.output.tag_configure("error", foreground="#FF7B72")
        self.output.tag_configure("warning", foreground="#F7C873")
        self.output.tag_configure("normal", foreground=COLORS["console_text"])

        self.responsive_mode = None

    def on_content_resize(self, _event=None) -> None:
        self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))

    def on_canvas_resize(self, event) -> None:
        self.content_canvas.itemconfigure(self.content_window, width=event.width)
        self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))

    def on_window_resize(self, event) -> None:
        if event.widget is self.root:
            self.apply_responsive_layout(event.width)

    def apply_responsive_layout(self, width: int) -> None:
        mode = "wide" if width >= 1020 else "compact"
        if mode == self.responsive_mode:
            return
        self.responsive_mode = mode

        if mode == "wide":
            self.form.columnconfigure(0, weight=1, uniform="section")
            self.form.columnconfigure(1, weight=0, uniform="")
            self.form.columnconfigure(2, weight=1, uniform="section")
            self.package_form.grid_configure(row=0, column=0, columnspan=1, sticky="nsew")
            self.form_separator.configure(orient="vertical")
            self.form_separator.grid_configure(row=0, column=1, columnspan=1, sticky="ns", padx=18, pady=0)
            self.remote_form.grid_configure(row=0, column=2, columnspan=1, sticky="nsew")

            for column in range(5):
                self.actions.columnconfigure(column, weight=1 if column == 1 else 0)
            self.action_badge.grid_configure(row=0, column=0, columnspan=1, rowspan=2, padx=(0, 12), pady=0, sticky="")
            self.action_title.grid_configure(row=0, column=1, columnspan=1, sticky="w")
            self.action_subtitle.grid_configure(row=1, column=1, columnspan=1, sticky="w")
            self.test_button.grid_configure(row=0, column=2, rowspan=2, padx=(18, 8), pady=0, sticky="ew")
            self.retry_button.grid_configure(row=0, column=3, rowspan=2, padx=(0, 8), pady=0, sticky="ew")
            self.run_button.grid_configure(row=0, column=4, rowspan=2, padx=0, pady=0, sticky="ew")
            self.header_badge.place(relx=1.0, rely=0.5, anchor="e")
            self.footer_info.pack(side="right", padx=24)
        else:
            self.form.columnconfigure(0, weight=1, uniform="")
            self.form.columnconfigure(1, weight=0, uniform="")
            self.form.columnconfigure(2, weight=0, uniform="")
            self.package_form.grid_configure(row=0, column=0, columnspan=3, sticky="ew")
            self.form_separator.configure(orient="horizontal")
            self.form_separator.grid_configure(row=1, column=0, columnspan=3, sticky="ew", padx=0, pady=14)
            self.remote_form.grid_configure(row=2, column=0, columnspan=3, sticky="ew")

            for column in range(5):
                self.actions.columnconfigure(column, weight=1 if column < 3 else 0)
            self.action_badge.grid_configure(row=0, column=0, columnspan=1, rowspan=2, padx=(0, 12), pady=0, sticky="w")
            self.action_title.grid_configure(row=0, column=1, columnspan=2, sticky="w")
            self.action_subtitle.grid_configure(row=1, column=1, columnspan=2, sticky="w")
            self.test_button.grid_configure(row=2, column=0, rowspan=1, padx=(0, 6), pady=(12, 0), sticky="ew")
            self.retry_button.grid_configure(row=2, column=1, rowspan=1, padx=6, pady=(12, 0), sticky="ew")
            self.run_button.grid_configure(row=2, column=2, rowspan=1, padx=(6, 0), pady=(12, 0), sticky="ew")
            self.header_badge.place_forget()
            self.footer_info.pack_forget()

        self.root.after_idle(self.on_content_resize)

    def create_card(self, parent) -> tk.Frame:
        return tk.Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["border"], highlightthickness=1)

    def add_section_title(self, parent, row, number, title, subtitle) -> None:
        badge = tk.Label(parent, text=number, bg="#E9F2FF", fg=COLORS["blue"], font=("Segoe UI", 10, "bold"), width=3, height=1)
        badge.grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=(0, 9))
        block = ttk.Frame(parent)
        block.grid(row=row, column=1, columnspan=2, sticky="ew", pady=(0, 9))
        ttk.Label(block, text=title, style="Field.TLabel").pack(anchor="w")
        ttk.Label(block, text=subtitle, style="Hint.TLabel").pack(anchor="w", pady=(2, 0))

    def add_path_row(self, parent, row, label, variable, callback, hint) -> None:
        label_frame = ttk.Frame(parent)
        label_frame.grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=5)
        ttk.Label(label_frame, text=label, style="Field.TLabel").pack(anchor="w")
        ttk.Label(label_frame, text=hint, style="Hint.TLabel", wraplength=210).pack(anchor="w", pady=(2, 0))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=5)
        ttk.Button(parent, text="Selecionar", style="Browse.TButton", command=callback).grid(row=row, column=2, padx=(8, 0), pady=5)

    def add_entry_row(self, parent, row, label, variable, hint) -> None:
        label_frame = ttk.Frame(parent)
        label_frame.grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=5)
        ttk.Label(label_frame, text=label, style="Field.TLabel").pack(anchor="w")
        ttk.Label(label_frame, text=hint, style="Hint.TLabel", wraplength=210).pack(anchor="w", pady=(2, 0))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)

    def add_validation_rows(self, parent, row) -> None:
        label_frame = ttk.Frame(parent)
        label_frame.grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=5)
        ttk.Label(label_frame, text="Validação", style="Field.TLabel").pack(anchor="w")
        ttk.Label(label_frame, text="Confirma o software após instalar", style="Hint.TLabel", wraplength=210).pack(anchor="w", pady=(2, 0))
        self.validation_combo = ttk.Combobox(parent, textvariable=self.validation_type, values=list(VALIDATION_LABELS), state="readonly", width=18)
        self.validation_combo.grid(row=row, column=1, sticky="ew", pady=5)
        self.validation_combo.bind("<<ComboboxSelected>>", self.on_validation_changed)
        self.validation_entry = ttk.Entry(parent, textvariable=self.validation_value)
        self.validation_entry.grid(row=row, column=2, sticky="ew", padx=(8, 0), pady=5)

    def add_throttle_row(self, parent, row) -> None:
        label_frame = ttk.Frame(parent)
        label_frame.grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=5)
        ttk.Label(label_frame, text="Simultâneas", style="Field.TLabel").pack(anchor="w")
        ttk.Label(label_frame, text="De 1 a 5 máquinas; padrão 4", style="Hint.TLabel", wraplength=210).pack(anchor="w", pady=(2, 0))
        ttk.Spinbox(parent, from_=1, to=5, textvariable=self.throttle, width=8).grid(row=row, column=1, sticky="w", pady=5)

    def on_validation_changed(self, _event=None) -> None:
        enabled = VALIDATION_LABELS.get(self.validation_type.get()) != "ExitCode"
        self.validation_entry.configure(state="normal" if enabled else "disabled")
        if not enabled:
            self.validation_value.set("")

    def choose_source(self) -> None:
        selected = filedialog.askdirectory(title="Selecione a pasta completa do pacote")
        if selected:
            self.source.set(selected)
            self.destination.set(Path(selected).name or "PacoteInstalacao")

    def choose_installer(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione o instalador",
            initialdir=self.source.get() or None,
            filetypes=[("Instaladores", "*.exe *.msi"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            previous_extension = Path(self.installer.get().strip()).suffix.lower()
            self.installer.set(selected)
            extension = Path(selected).suffix.lower()
            current_args = self.silent_args.get().strip()
            if extension == ".msi" and (not current_args or previous_extension == ".exe"):
                self.silent_args.set(DEFAULT_MSI_ARGS)
            elif extension == ".exe" and (
                previous_extension == ".msi" or has_known_msi_args(current_args)
            ):
                self.silent_args.set("")

    def choose_list(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione a lista de patrimônios",
            filetypes=[("Listas", "*.txt *.csv"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            self.computer_list.set(selected)

    def choose_psexec(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione PsExec.exe ou PsExec64.exe",
            filetypes=[("PsExec", "PsExec*.exe"), ("Executáveis", "*.exe")],
        )
        if selected:
            self.psexec.set(selected)

    def validate(self, mode: str, list_override: str | None = None) -> bool:
        list_path = Path(list_override or self.computer_list.get().strip())
        if not list_path.is_file() or list_path.suffix.lower() not in {".txt", ".csv"}:
            messagebox.showwarning(APP_NAME, "Selecione uma lista TXT ou CSV válida.")
            return False
        try:
            if not 1 <= int(self.throttle.get()) <= 5:
                raise ValueError
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "A quantidade simultânea deve estar entre 1 e 5.")
            return False
        if mode == "Test":
            return True

        source = Path(self.source.get().strip())
        installer = Path(self.installer.get().strip())
        psexec = Path(self.psexec.get().strip())
        if not source.is_dir():
            messagebox.showwarning(APP_NAME, "Selecione uma pasta de pacote válida.")
            return False
        try:
            self.remote_source_path()
        except ValueError as error:
            messagebox.showwarning(APP_NAME, str(error))
            return False
        if not installer.is_file() or installer.suffix.lower() not in {".exe", ".msi"}:
            messagebox.showwarning(APP_NAME, "Selecione um instalador EXE ou MSI válido.")
            return False
        try:
            if os.path.commonpath([str(source.resolve()), str(installer.resolve())]).lower() != str(source.resolve()).lower():
                raise ValueError
        except (ValueError, OSError):
            messagebox.showwarning(APP_NAME, "O instalador precisa estar dentro da pasta do pacote.")
            return False
        if not psexec.is_file() or psexec.name.lower() not in {"psexec.exe", "psexec64.exe"}:
            messagebox.showwarning(APP_NAME, "Selecione PsExec.exe ou PsExec64.exe.")
            return False
        destination = self.destination.get().strip()
        if not destination or any(char in destination for char in '<>:"/\\|?*'):
            messagebox.showwarning(APP_NAME, "Informe apenas um nome de pasta válido para C:\\Temp.")
            return False
        if installer.suffix.lower() == ".exe" and not self.silent_args.get().strip():
            if not messagebox.askyesno(APP_NAME, "Nenhum parâmetro foi informado para o EXE. Deseja continuar mesmo assim?"):
                return False
        validation_code = VALIDATION_LABELS.get(self.validation_type.get(), "ExitCode")
        validation_value = self.validation_value.get().strip()
        if validation_code == "File" and (not validation_value or not re.match(r"^[Cc]:\\", validation_value)):
            messagebox.showwarning(APP_NAME, "Para validar por arquivo, informe o caminho completo iniciado por C:\\.")
            return False
        if validation_code == "Service" and not validation_value:
            messagebox.showwarning(APP_NAME, "Informe o nome interno do serviço do Windows.")
            return False
        return True

    def remote_source_path(self) -> str:
        source = self.source.get().strip().replace("/", "\\")
        if source.upper() == "J:":
            return SOFTWARE_UNC_ROOT
        if source.upper().startswith("J:\\"):
            return SOFTWARE_UNC_ROOT + source[2:]
        if source.startswith("\\\\"):
            return source.rstrip("\\")
        raise ValueError("Para a cópia direta, selecione uma pasta da unidade J: ou informe um caminho UNC.")

    def build_command(self, mode: str, list_path: str) -> list[str]:
        command = [
            "powershell.exe", "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(resource_path("RemoteInstallerParallel.ps1")),
            "-Mode", mode,
            "-ComputerList", list_path,
            "-ThrottleLimit", str(self.throttle.get()),
        ]
        if mode == "Install":
            command.extend([
                "-SourceFolder", self.source.get().strip(),
                "-RemoteSourceFolder", self.remote_source_path(),
                "-Installer", self.installer.get().strip(),
                "-SilentArgs", self.silent_args.get().strip(),
                "-PsExecPath", self.psexec.get().strip(),
                "-DestinationName", self.destination.get().strip(),
                "-ValidationType", VALIDATION_LABELS.get(self.validation_type.get(), "ExitCode"),
                "-ValidationValue", self.validation_value.get().strip(),
            ])
        return command

    def start(self, mode: str, list_override: str | None = None) -> None:
        list_path = list_override or self.computer_list.get().strip()
        if self.running or not self.validate(mode, list_path):
            return
        computers = self.read_computers(list_path)
        if not computers:
            messagebox.showwarning(APP_NAME, "A lista não possui patrimônios válidos.")
            return
        if mode == "Install" and not messagebox.askyesno(
            APP_NAME,
            f"A instalação será executada em {len(computers)} máquina(s), com até {self.throttle.get()} simultâneas. Continuar?",
        ):
            return

        self.save_settings()
        self.clear_output()
        self.prepare_machine_table(computers)
        self.active_list_path = list_path
        self.append_output("Iniciando teste de acesso…" if mode == "Test" else "Iniciando cópia e instalação…")
        self.set_running(True)
        threading.Thread(target=self.run_command, args=(self.build_command(mode, list_path),), daemon=True).start()

    def read_computers(self, path: str) -> list[str]:
        try:
            file_path = Path(path)
            if file_path.suffix.lower() == ".csv":
                text = file_path.read_text(encoding="utf-8-sig", errors="replace")
                first_line = text.splitlines()[0] if text.splitlines() else ""
                delimiter = ";" if first_line.count(";") > first_line.count(",") else ","
                rows = list(csv.DictReader(text.splitlines(), delimiter=delimiter))
                if not rows:
                    return []
                preferred = ["Patrimonio", "Patrimônio", "Computer", "Computador", "Hostname", "Host", "Name", "Nome"]
                column = next((name for name in preferred if name in rows[0]), next(iter(rows[0]), ""))
                values = [row.get(column, "") for row in rows]
            else:
                values = file_path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
        except OSError:
            return []

        result = []
        for value in values:
            name = str(value).strip().upper()
            if name.isdigit():
                name = "W" + name
            if re.fullmatch(r"W\d+", name) and name not in result:
                result.append(name)
        return result

    def prepare_machine_table(self, computers: list[str]) -> None:
        for item in self.machine_table.get_children():
            self.machine_table.delete(item)
        self.machine_rows.clear()
        self.machine_overall.clear()
        for computer in computers:
            item = self.machine_table.insert("", "end", values=(computer, "Aguardando", "—", "—", "—", "—", "—", "Na fila"))
            self.machine_rows[computer] = item
            self.machine_overall[computer] = "Pendente"
        self.retry_button.configure(state="disabled")

    def retry_failures(self) -> None:
        failures = [computer for computer, overall in self.machine_overall.items() if overall == "Falha"]
        if not failures:
            messagebox.showinfo(APP_NAME, "Não há máquinas com falha para reexecutar.")
            return
        retry_path = self.settings_path.parent / "reexecutar_falhas.txt"
        try:
            retry_path.parent.mkdir(parents=True, exist_ok=True)
            retry_path.write_text("\n".join(failures) + "\n", encoding="utf-8")
        except OSError as error:
            messagebox.showerror(APP_NAME, f"Não foi possível criar a lista de reexecução:\n{error}")
            return
        self.start("Install", str(retry_path))

    def copy_report(self) -> None:
        items = self.machine_table.get_children()
        if not items:
            messagebox.showinfo(APP_NAME, "Ainda não há resultados para copiar.")
            return

        headings = [
            "Patrimônio", "Acesso", "Cópia", "Instalação",
            "Validação", "Limpeza", "Tempo", "Detalhes",
        ]
        rows = [list(self.machine_table.item(item, "values")) for item in items]
        report = build_table_report(headings, rows)
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(report)
            self.root.update_idletasks()
        except tk.TclError as error:
            messagebox.showerror(APP_NAME, f"Não foi possível copiar o relatório:\n{error}")
            return
        self.status.set(f"Relatório copiado: {len(rows)} máquina(s)")
        messagebox.showinfo(APP_NAME, "Relatório copiado para a área de transferência.")

    def update_machine_status(self, text: str) -> None:
        match = STATUS_PATTERN.match(text.strip())
        if not match:
            return
        computer = match.group("computer")
        if computer not in self.machine_rows:
            item = self.machine_table.insert("", "end", values=(computer, "Aguardando", "—", "—", "—", "—", "—", ""))
            self.machine_rows[computer] = item
            self.machine_overall[computer] = "Pendente"
        item = self.machine_rows[computer]
        values = list(self.machine_table.item(item, "values"))
        stage = match.group("stage").strip().upper()
        result = match.group("result").strip()
        details = (match.group("details") or "").strip()
        indexes = {"ACESSO": 1, "COPIA": 2, "INSTALACAO": 3, "VALIDACAO": 4, "LIMPEZA": 5}
        if stage in indexes:
            values[indexes[stage]] = result.title() if result.upper() == "INICIANDO" else result
        elif stage == "FILA":
            values[7] = "Processando"
        elif stage == "FINAL":
            overall = "Sucesso" if result.upper() == "SUCESSO" else "Falha"
            self.machine_overall[computer] = overall
            duration_match = re.search(r"Tempo\s+([\d.,]+)s", details, re.IGNORECASE)
            if duration_match:
                values[6] = duration_match.group(1).replace(",", ".") + " s"
        if details:
            values[7] = details
        self.machine_table.item(item, values=values)
        overall = self.machine_overall.get(computer)
        tag = "success" if overall == "Sucesso" else "failure" if overall == "Falha" else "running"
        self.machine_table.item(item, tags=(tag,))
        self.machine_table.see(item)

    def run_command(self, command: list[str]) -> None:
        try:
            self.process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=CREATE_NO_WINDOW,
            )
            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.output_queue.put(("line", line.rstrip()))
            code = self.process.wait()
            self.output_queue.put(("done", code))
        except Exception as error:
            self.output_queue.put(("error", str(error)))
        finally:
            self.process = None

    def consume_output(self) -> None:
        try:
            while True:
                kind, value = self.output_queue.get_nowait()
                if kind == "line":
                    self.append_output(value)
                elif kind == "done":
                    self.append_output("Processo concluído." if value == 0 else f"Processo encerrado com erro (código {value}).")
                    self.set_running(False)
                    if value != 0:
                        for computer, overall in list(self.machine_overall.items()):
                            if overall == "Pendente":
                                self.machine_overall[computer] = "Falha"
                                item = self.machine_rows[computer]
                                values = list(self.machine_table.item(item, "values"))
                                values[7] = "Processo encerrado antes de concluir"
                                self.machine_table.item(item, values=values, tags=("failure",))
                    failures = sum(1 for overall in self.machine_overall.values() if overall == "Falha")
                    if failures:
                        self.status.set(f"Concluído com {failures} falha(s)")
                        self.status_label.configure(bg="#FFE9E7", fg=COLORS["red"])
                        self.retry_button.configure(state="normal")
                elif kind == "error":
                    self.append_output(f"Erro ao iniciar o PowerShell: {value}")
                    self.set_running(False)
        except queue.Empty:
            pass
        self.root.after(100, self.consume_output)

    def append_output(self, text: str) -> None:
        self.update_machine_status(text)
        upper = text.upper()
        if "ERRO" in upper or "INDISPONIVEL" in upper or "OFFLINE" in upper:
            tag = "error"
        elif "IGNORADO" in upper or "REINICIO" in upper or "AVISO" in upper or "MANTIDA" in upper:
            tag = "warning"
        elif "SUCESSO" in upper or "] OK" in upper or ": OK" in upper or "CONCLUIDO" in upper:
            tag = "success"
        else:
            tag = "normal"
        self.output.configure(state="normal")
        self.output.insert("end", text + "\n", tag)
        self.output.see("end")
        self.output.configure(state="disabled")

    def clear_output(self) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.configure(state="disabled")

    def set_running(self, running: bool) -> None:
        self.running = running
        state = "disabled" if running else "normal"
        self.test_button.configure(state=state)
        self.run_button.configure(state=state)
        if running:
            self.retry_button.configure(state="disabled")
        self.status.set("Executando… não feche a ferramenta" if running else "Pronto")
        if running:
            self.status_label.configure(bg="#FFF1D6", fg=COLORS["amber"])
            self.progress.start(12)
        else:
            self.status_label.configure(bg=COLORS["green_soft"], fg=COLORS["green"])
            self.progress.stop()

    def load_settings(self) -> None:
        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            self.source.set(data.get("source", ""))
            self.installer.set(data.get("installer", ""))
            self.silent_args.set(data.get("silent_args", ""))
            if (
                Path(self.installer.get().strip()).suffix.lower() == ".exe"
                and has_known_msi_args(self.silent_args.get())
            ):
                self.silent_args.set("")
            self.computer_list.set(data.get("computer_list", ""))
            self.psexec.set(data.get("psexec") or DEFAULT_PSEXEC)
            self.destination.set(data.get("destination", "PacoteInstalacao"))
            self.throttle.set(data.get("throttle", 4))
            self.validation_type.set(data.get("validation_type", "Código de saída"))
            self.validation_value.set(data.get("validation_value", ""))
        except (OSError, ValueError, TypeError):
            pass

    def save_settings(self) -> None:
        try:
            self.settings_path.parent.mkdir(parents=True, exist_ok=True)
            data = {
                "source": self.source.get(),
                "installer": self.installer.get(),
                "silent_args": self.silent_args.get(),
                "computer_list": self.computer_list.get(),
                "psexec": self.psexec.get(),
                "destination": self.destination.get(),
                "throttle": self.throttle.get(),
                "validation_type": self.validation_type.get(),
                "validation_value": self.validation_value.get(),
            }
            self.settings_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def on_close(self) -> None:
        if self.running:
            messagebox.showwarning(APP_NAME, "Aguarde a operação terminar antes de fechar. A instalação remota pode estar em andamento.")
            return
        self.save_settings()
        self.root.destroy()


def main() -> None:
    if os.name != "nt":
        raise SystemExit("Esta ferramenta deve ser executada no Windows.")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    if not is_admin():
        if not relaunch_as_admin():
            ctypes.windll.user32.MessageBoxW(None, "A execução como administrador foi cancelada.", APP_NAME, 0x30)
        return
    root = tk.Tk()
    InstallerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
