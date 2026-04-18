from __future__ import annotations

import ctypes
import importlib
import io
import math
import re
import tempfile
import textwrap
import traceback
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

APP_NAME = "DocuLoom Prism"
APP_SLUG = "doculoom_prism"
APP_VERSION = "1.0.0"
APP_TAGLINE = "Presentation-ready document analytics with graph and media export."
APP_ID = "DocuLoomPrism.Desktop.1"
ASSETS_DIR = Path(__file__).resolve().parent / "assets"
ICON_ICO_PATH = ASSETS_DIR / "prizm.ico"
ICON_PNG_PATH = ASSETS_DIR / "logo.png"
DOWNLOADS_DIR = Path(tempfile.gettempdir()) / APP_SLUG / "downloads"
DOCUMENT_SUFFIXES = {".csv", ".xlsx", ".xls", ".xlsm", ".pdf", ".docx", ".pptx", ".txt"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
CONTENT_TYPE_SUFFIXES = {
    "text/csv": ".csv",
    "text/plain": ".txt",
    "application/pdf": ".pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.ms-excel.sheet.macroenabled.12": ".xlsm",
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/webp": ".webp",
}

SUPPORTED_FILE_TYPES = (
    ("Supported files", "*.csv *.xlsx *.xls *.xlsm *.pdf *.docx *.pptx *.txt"),
    ("CSV files", "*.csv"),
    ("Excel files", "*.xlsx *.xls *.xlsm"),
    ("PDF files", "*.pdf"),
    ("Word files", "*.docx"),
    ("PowerPoint files", "*.pptx"),
    ("Text files", "*.txt"),
    ("All files", "*.*"),
)

CHART_TYPES = (
    "Bar Chart",
    "Line Chart",
    "Horizontal Bar Chart",
    "Area Chart",
    "Scatter Plot",
    "Pie Chart",
    "Histogram",
)


@dataclass
class DatasetChoice:
    name: str
    dataframe: pd.DataFrame
    category_column: str
    numeric_columns: list[str]
    source_note: str


def _optional_attr(module_name: str, attr_name: str) -> object | None:
    try:
        module = importlib.import_module(module_name)
    except ImportError:
        return None
    return getattr(module, attr_name, None)


def _missing_dependency_message(package_name: str) -> str:
    return (
        f"{package_name} is not installed in the Python environment running this app.\n\n"
        "Install all required packages with:\n"
        "pip install -r requirements.txt"
    )


def _guess_suffix_from_url_or_type(resource_url: str, content_type: str) -> str:
    parsed_path = urllib.parse.urlparse(resource_url).path
    suffix = Path(parsed_path).suffix.lower()
    if suffix:
        return suffix
    normalized_content_type = content_type.split(";")[0].strip().lower()
    return CONTENT_TYPE_SUFFIXES.get(normalized_content_type, "")


def _normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    working = df.copy()
    working.columns = [str(column).strip() or f"Column {index + 1}" for index, column in enumerate(working.columns)]
    working = working.dropna(how="all").reset_index(drop=True)
    for column in working.columns:
        if working[column].dtype == object:
            working[column] = working[column].astype(str).str.strip()
            working.loc[working[column].isin({"", "nan", "None"}), column] = pd.NA
    return working


def _coerce_numeric_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    working = df.copy()
    numeric_columns: list[str] = []
    for column in working.columns:
        converted = pd.to_numeric(working[column], errors="coerce")
        if converted.notna().sum() >= max(2, math.ceil(len(working) * 0.5)):
            working[column] = converted
            numeric_columns.append(column)
    return working, numeric_columns


def _build_dataset_choices(df: pd.DataFrame, label: str, source_note: str) -> list[DatasetChoice]:
    normalized = _normalize_dataframe(df)
    if normalized.empty:
        return []

    normalized, numeric_columns = _coerce_numeric_columns(normalized)
    if not numeric_columns:
        return []

    non_numeric_columns = [column for column in normalized.columns if column not in numeric_columns]
    category_column = non_numeric_columns[0] if non_numeric_columns else "Record Number"

    if category_column == "Record Number":
        normalized.insert(0, category_column, [f"Record {index + 1}" for index in range(len(normalized))])

    filtered = normalized[[category_column, *numeric_columns]].dropna(subset=numeric_columns, how="all")
    if filtered.empty:
        return []

    return [
        DatasetChoice(
            name=label,
            dataframe=filtered.reset_index(drop=True),
            category_column=category_column,
            numeric_columns=numeric_columns,
            source_note=source_note,
        )
    ]


def _extract_label_value_pairs(text: str, label_prefix: str) -> list[DatasetChoice]:
    pairs: list[tuple[str, float]] = []
    seen_labels: set[str] = set()

    pattern = re.compile(
        r"^\s*(?P<label>[A-Za-z][A-Za-z0-9 /&().,_-]{1,80}?)\s*(?:[:=-]|,)\s*(?P<value>[-+]?\d+(?:\.\d+)?)\s*$"
    )

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        match = pattern.match(stripped)
        if not match:
            continue
        label = re.sub(r"\s+", " ", match.group("label")).strip(" -")
        value = float(match.group("value"))
        if len(label) < 2 or label.lower() in seen_labels:
            continue
        pairs.append((label, value))
        seen_labels.add(label.lower())

    if len(pairs) >= 2:
        table = pd.DataFrame(pairs, columns=["Topic", "Value"])
        return [
            DatasetChoice(
                name=f"{label_prefix} - extracted values",
                dataframe=table,
                category_column="Topic",
                numeric_columns=["Value"],
                source_note="Numbers were inferred from text lines that looked like label-value pairs.",
            )
        ]

    number_pattern = re.compile(r"[-+]?\d+(?:\.\d+)?")
    numbers = [float(match.group(0)) for match in number_pattern.finditer(text)]
    if len(numbers) >= 2:
        table = pd.DataFrame(
            {
                "Topic": [f"Item {index + 1}" for index in range(len(numbers))],
                "Value": numbers,
            }
        )
        return [
            DatasetChoice(
                name=f"{label_prefix} - numeric sequence",
                dataframe=table,
                category_column="Topic",
                numeric_columns=["Value"],
                source_note="No table was found, so the app graphed the numeric sequence discovered in the document text.",
            )
        ]

    return []


def _read_csv(file_path: Path) -> list[DatasetChoice]:
    df = pd.read_csv(file_path)
    return _build_dataset_choices(df, f"{file_path.name} - CSV data", "Data loaded from the CSV file.")


def _read_excel(file_path: Path) -> list[DatasetChoice]:
    sheets = pd.read_excel(file_path, sheet_name=None)
    datasets: list[DatasetChoice] = []
    for sheet_name, frame in sheets.items():
        datasets.extend(
            _build_dataset_choices(
                frame,
                f"{file_path.name} - {sheet_name}",
                f"Data loaded from worksheet '{sheet_name}'.",
            )
        )
    return datasets


def _read_docx(file_path: Path) -> list[DatasetChoice]:
    document_factory = _optional_attr("docx", "Document")
    if document_factory is None:
        raise RuntimeError(_missing_dependency_message("python-docx"))

    document = document_factory(str(file_path))
    datasets: list[DatasetChoice] = []

    for index, table in enumerate(document.tables, start=1):
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        if len(rows) < 2:
            continue
        header = rows[0]
        body = rows[1:]
        if not any(header):
            continue
        frame = pd.DataFrame(body, columns=header)
        datasets.extend(
            _build_dataset_choices(
                frame,
                f"{file_path.name} - Table {index}",
                f"Data loaded from Word table {index}.",
            )
        )

    if datasets:
        return datasets

    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    return _extract_label_value_pairs(text, file_path.name)


def _read_pptx(file_path: Path) -> list[DatasetChoice]:
    presentation_factory = _optional_attr("pptx", "Presentation")
    if presentation_factory is None:
        raise RuntimeError(_missing_dependency_message("python-pptx"))

    presentation = presentation_factory(str(file_path))
    datasets: list[DatasetChoice] = []
    text_parts: list[str] = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if hasattr(shape, "has_table") and shape.has_table:
                rows = []
                for row in shape.table.rows:
                    rows.append([cell.text.strip() for cell in row.cells])
                if len(rows) >= 2:
                    header = rows[0]
                    body = rows[1:]
                    if any(header):
                        frame = pd.DataFrame(body, columns=header)
                        datasets.extend(
                            _build_dataset_choices(
                                frame,
                                f"{file_path.name} - Slide {slide_index} Table",
                                f"Data loaded from table on slide {slide_index}.",
                            )
                        )
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text = "".join(run.text for run in paragraph.runs).strip()
                    if text:
                        text_parts.append(text)

    if datasets:
        return datasets

    return _extract_label_value_pairs("\n".join(text_parts), file_path.name)


def _read_pdf(file_path: Path) -> list[DatasetChoice]:
    pdf_reader = _optional_attr("pypdf", "PdfReader")
    if pdf_reader is None:
        raise RuntimeError(_missing_dependency_message("pypdf"))

    reader = pdf_reader(str(file_path))
    text_parts = []
    for page in reader.pages:
        extracted = page.extract_text() or ""
        if extracted.strip():
            text_parts.append(extracted)

    return _extract_label_value_pairs("\n".join(text_parts), file_path.name)


def _read_txt(file_path: Path) -> list[DatasetChoice]:
    return _extract_label_value_pairs(file_path.read_text(encoding="utf-8", errors="ignore"), file_path.name)


def load_datasets(file_path: Path) -> list[DatasetChoice]:
    suffix = file_path.suffix.lower()
    readers = {
        ".csv": _read_csv,
        ".xlsx": _read_excel,
        ".xls": _read_excel,
        ".xlsm": _read_excel,
        ".docx": _read_docx,
        ".pptx": _read_pptx,
        ".pdf": _read_pdf,
        ".txt": _read_txt,
    }

    if suffix not in readers:
        raise RuntimeError(
            f"Unsupported file type '{suffix}'. Please choose CSV, Excel, PDF, DOCX, PPTX, or TXT."
        )

    datasets = readers[suffix](file_path)
    if not datasets:
        raise RuntimeError(
            "The app could not find chart-ready numeric data in this file. Try a document with a table or clear numbers."
        )
    return datasets


class DocuLoomPrismApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self._configure_windows_identity()
        self.root.title(APP_NAME)
        self.root.geometry("1320x820")
        self.root.minsize(1120, 720)
        self.root.configure(bg="#f4efe6")

        self.file_path_var = tk.StringVar()
        self.image_url_var = tk.StringVar()
        self.chart_type_var = tk.StringVar(value=CHART_TYPES[0])
        self.dataset_var = tk.StringVar()
        self.value_column_var = tk.StringVar()
        self.status_var = tk.StringVar(value=f"{APP_NAME} is ready. Load a document to start building a publishable visual.")
        self.note_var = tk.StringVar(
            value="Best with tables or documents that contain clear numeric values such as sales, scores, budgets, or counts."
        )
        self.image_status_var = tk.StringVar(
            value="Optional URL: paste a direct image or supported document link to fetch it."
        )
        self.dataset_metric_var = tk.StringVar(value="0 datasets")
        self.graph_metric_var = tk.StringVar(value="No graph yet")
        self.image_metric_var = tk.StringVar(value="No linked image")

        self.datasets: list[DatasetChoice] = []
        self.current_figure: Figure | None = None
        self.figure_counter = 0
        self.current_chart_title = ""
        self.current_chart_subtitle = ""
        self.current_chart_note = ""
        self.current_chart_frame = pd.DataFrame()
        self.current_value_column = ""
        self.current_chart_type = ""
        self.reference_image_bytes: bytes | None = None
        self.reference_image_name = ""
        self.reference_image_url = ""
        self.reference_image_preview: object | None = None
        self.window_icon_preview: object | None = None

        self._configure_window_branding()
        self._configure_styles()
        self._build_layout()
        self._build_menu()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure("App.TFrame", background="#f4efe6")
        style.configure("Panel.TFrame", background="#fffaf2")
        style.configure("Hero.TFrame", background="#143f38")
        style.configure("Stat.TFrame", background="#f0e7d8")
        style.configure(
            "Title.TLabel",
            background="#f4efe6",
            foreground="#123b35",
            font=("Segoe UI Semibold", 24),
        )
        style.configure(
            "Body.TLabel",
            background="#f4efe6",
            foreground="#3f5e58",
            font=("Segoe UI", 10),
        )
        style.configure(
            "HeroTitle.TLabel",
            background="#143f38",
            foreground="#f8f4eb",
            font=("Georgia", 24, "bold"),
        )
        style.configure(
            "HeroBody.TLabel",
            background="#143f38",
            foreground="#dcebe6",
            font=("Segoe UI", 10),
        )
        style.configure(
            "PanelTitle.TLabel",
            background="#fffaf2",
            foreground="#123b35",
            font=("Segoe UI Semibold", 11),
        )
        style.configure(
            "StatLabel.TLabel",
            background="#f0e7d8",
            foreground="#6f5a42",
            font=("Segoe UI Semibold", 9),
        )
        style.configure(
            "StatValue.TLabel",
            background="#f0e7d8",
            foreground="#123b35",
            font=("Segoe UI Semibold", 14),
        )
        style.configure(
            "Accent.TButton",
            background="#1d6f63",
            foreground="#ffffff",
            font=("Segoe UI Semibold", 10),
            padding=(16, 9),
            borderwidth=0,
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#145247"), ("disabled", "#93b0aa")],
            foreground=[("disabled", "#f7faf9")],
        )
        style.configure(
            "Secondary.TButton",
            background="#d9c7aa",
            foreground="#123b35",
            font=("Segoe UI Semibold", 10),
            padding=(14, 9),
            borderwidth=0,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#cbb18a"), ("disabled", "#ebe0cf")],
            foreground=[("disabled", "#7a7063")],
        )
        style.configure(
            "Treeview",
            font=("Segoe UI", 10),
            rowheight=28,
            background="#fffdf8",
            fieldbackground="#fffdf8",
            foreground="#243330",
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI Semibold", 10),
            background="#d8e6df",
            foreground="#123b35",
            relief="flat",
        )

    def _configure_window_branding(self) -> None:
        try:
            if ICON_ICO_PATH.exists():
                self.root.iconbitmap(default=str(ICON_ICO_PATH))
                return
        except tk.TclError:
            pass

    def _configure_windows_identity(self) -> None:
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except Exception:
            pass

        try:
            if ICON_PNG_PATH.exists():
                icon_image = tk.PhotoImage(file=str(ICON_PNG_PATH))
                self.window_icon_preview = icon_image
                self.root.iconphoto(True, icon_image)
        except tk.TclError:
            pass

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self.root)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Open Document", command=self._browse_file)
        file_menu.add_command(label="Fetch URL Resource", command=self._fetch_linked_image)
        file_menu.add_separator()
        file_menu.add_command(label="Generate Graph", command=self._generate_graph)
        file_menu.add_command(label="Export PowerPoint", command=self._export_to_ppt)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.destroy)
        menu_bar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menu_bar)

    def _show_about(self) -> None:
        messagebox.showinfo(
            f"About {APP_NAME}",
            (
                f"{APP_NAME} {APP_VERSION}\n\n"
                f"{APP_TAGLINE}\n\n"
                "Load documents, fetch supported files or images from URLs, build clear graphs, and export a presentation-ready PowerPoint."
            ),
        )

    def _build_layout(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        header = ttk.Frame(outer, style="Hero.TFrame", padding=22)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=2)
        header.columnconfigure(1, weight=1)

        title_area = ttk.Frame(header, style="Hero.TFrame")
        title_area.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        title_area.columnconfigure(0, weight=1)

        ttk.Label(title_area, text=APP_NAME, style="HeroTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            title_area,
            text=APP_TAGLINE,
            style="HeroBody.TLabel",
            wraplength=720,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 10))
        ttk.Label(
            title_area,
            text="Supports PDF, Word, Excel, PowerPoint, CSV, TXT, and direct image or document URLs for professional report output.",
            style="HeroBody.TLabel",
        ).grid(row=2, column=0, sticky="w")

        stats = ttk.Frame(header, style="Hero.TFrame")
        stats.grid(row=0, column=1, sticky="nsew")
        for column in range(3):
            stats.columnconfigure(column, weight=1)

        self._build_stat_card(stats, 0, "Datasets", self.dataset_metric_var)
        self._build_stat_card(stats, 1, "Graph", self.graph_metric_var)
        self._build_stat_card(stats, 2, "Image", self.image_metric_var)

        control_panel = ttk.Frame(outer, style="Panel.TFrame", padding=18)
        control_panel.grid(row=1, column=0, sticky="ew", pady=(16, 14))
        for column in range(4):
            control_panel.columnconfigure(column, weight=1)

        ttk.Label(control_panel, text="Document File", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(control_panel, textvariable=self.file_path_var, font=("Segoe UI", 10)).grid(
            row=1, column=0, columnspan=3, sticky="ew", padx=(0, 12), pady=(6, 14)
        )
        ttk.Button(control_panel, text="Browse File", style="Secondary.TButton", command=self._browse_file).grid(
            row=1, column=3, sticky="ew", pady=(6, 14)
        )

        ttk.Label(control_panel, text="File Or Image URL (optional)", style="PanelTitle.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(control_panel, textvariable=self.image_url_var, font=("Segoe UI", 10)).grid(
            row=3, column=0, columnspan=3, sticky="ew", padx=(0, 12), pady=(6, 14)
        )
        ttk.Button(
            control_panel,
            text="Fetch URL",
            style="Secondary.TButton",
            command=self._fetch_linked_image,
        ).grid(row=3, column=3, sticky="ew", pady=(6, 14))

        ttk.Label(control_panel, text="Dataset", style="PanelTitle.TLabel").grid(row=4, column=0, sticky="w")
        self.dataset_box = ttk.Combobox(control_panel, textvariable=self.dataset_var, state="readonly", font=("Segoe UI", 10))
        self.dataset_box.grid(row=5, column=0, sticky="ew", padx=(0, 12), pady=(6, 14))
        self.dataset_box.bind("<<ComboboxSelected>>", self._on_dataset_changed)

        ttk.Label(control_panel, text="Value Column", style="PanelTitle.TLabel").grid(row=4, column=1, sticky="w")
        self.value_box = ttk.Combobox(control_panel, textvariable=self.value_column_var, state="readonly", font=("Segoe UI", 10))
        self.value_box.grid(row=5, column=1, sticky="ew", padx=(0, 12), pady=(6, 14))
        self.value_box.bind("<<ComboboxSelected>>", self._on_value_changed)

        ttk.Label(control_panel, text="Graph Type", style="PanelTitle.TLabel").grid(row=4, column=2, sticky="w")
        self.chart_type_box = ttk.Combobox(
            control_panel,
            textvariable=self.chart_type_var,
            state="readonly",
            values=CHART_TYPES,
            font=("Segoe UI", 10),
        )
        self.chart_type_box.grid(row=5, column=2, sticky="ew", padx=(0, 12), pady=(6, 14))

        actions = ttk.Frame(control_panel, style="Panel.TFrame")
        actions.grid(row=5, column=3, sticky="ew", pady=(6, 14))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        actions.columnconfigure(2, weight=1)

        ttk.Button(actions, text="Load Data", style="Accent.TButton", command=self._load_selected_file).grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(actions, text="Generate Graph", style="Accent.TButton", command=self._generate_graph).grid(
            row=0, column=1, sticky="ew", padx=(0, 8)
        )
        ttk.Button(actions, text="Export PPT", style="Accent.TButton", command=self._export_to_ppt).grid(
            row=0, column=2, sticky="ew"
        )

        ttk.Label(control_panel, textvariable=self.status_var, style="Body.TLabel", wraplength=1160).grid(
            row=6, column=0, columnspan=4, sticky="w"
        )
        ttk.Label(control_panel, textvariable=self.image_status_var, style="Body.TLabel", wraplength=1160).grid(
            row=7, column=0, columnspan=4, sticky="w", pady=(4, 0)
        )
        ttk.Label(control_panel, textvariable=self.note_var, style="Body.TLabel", wraplength=1160).grid(
            row=8, column=0, columnspan=4, sticky="w", pady=(4, 0)
        )

        content = ttk.Frame(outer, style="App.TFrame")
        content.grid(row=2, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        preview_card = ttk.Frame(content, style="Panel.TFrame", padding=18)
        preview_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        preview_card.rowconfigure(1, weight=1)
        preview_card.columnconfigure(0, weight=1)

        ttk.Label(preview_card, text="Data Explorer", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")

        preview_columns = ("row", "category", "value")
        self.preview = ttk.Treeview(preview_card, columns=preview_columns, show="headings", height=12)
        self.preview.heading("row", text="#")
        self.preview.heading("category", text="Topic / Category")
        self.preview.heading("value", text="Selected Value")
        self.preview.column("row", width=60, anchor="center")
        self.preview.column("category", width=210, anchor="w")
        self.preview.column("value", width=120, anchor="e")
        self.preview.grid(row=1, column=0, sticky="nsew", pady=(10, 12))

        ttk.Label(preview_card, text="Linked Image Preview", style="PanelTitle.TLabel").grid(row=2, column=0, sticky="w")
        self.image_preview_label = tk.Label(
            preview_card,
            text="No linked image loaded yet.\nPaste a direct image or supported file URL above.",
            bg="#f4ecdf",
            fg="#526a66",
            font=("Segoe UI", 10),
            relief="solid",
            bd=1,
            height=10,
            justify="center",
        )
        self.image_preview_label.grid(row=3, column=0, sticky="ew", pady=(10, 0))

        graph_card = ttk.Frame(content, style="Panel.TFrame", padding=18)
        graph_card.grid(row=0, column=1, sticky="nsew")
        graph_card.rowconfigure(1, weight=1)
        graph_card.columnconfigure(0, weight=1)

        ttk.Label(graph_card, text="Insight Canvas", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.chart_host = ttk.Frame(graph_card, style="Panel.TFrame")
        self.chart_host.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.chart_host.rowconfigure(0, weight=1)
        self.chart_host.columnconfigure(0, weight=1)

        self._render_placeholder_figure()

    def _build_stat_card(self, parent: ttk.Frame, column: int, label: str, value_var: tk.StringVar) -> None:
        card = ttk.Frame(parent, style="Stat.TFrame", padding=14)
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0))
        card.columnconfigure(0, weight=1)
        ttk.Label(card, text=label, style="StatLabel.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, textvariable=value_var, style="StatValue.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))

    def _browse_file(self) -> None:
        selected = filedialog.askopenfilename(filetypes=SUPPORTED_FILE_TYPES)
        if not selected:
            return
        self.file_path_var.set(selected)
        self._load_selected_file()

    def _fetch_linked_image(self) -> None:
        resource_url = self.image_url_var.get().strip()
        if not resource_url:
            messagebox.showinfo("Add URL", "Please paste a direct image or supported document URL first.")
            return

        try:
            request = urllib.request.Request(resource_url, headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"})
            with urllib.request.urlopen(request, timeout=20) as response:
                content_type = response.headers.get("Content-Type", "")
                resource_bytes = response.read()
        except (urllib.error.URLError, ValueError) as exc:
            messagebox.showerror("Unable to fetch URL", f"Could not download the resource from the URL.\n\n{exc}")
            self.image_status_var.set("The URL could not be downloaded. Check the link and try again.")
            self.image_metric_var.set("No linked image")
            return

        suffix = _guess_suffix_from_url_or_type(resource_url, content_type)
        if suffix in DOCUMENT_SUFFIXES:
            self._handle_downloaded_document(resource_url, resource_bytes, suffix)
            return

        if suffix in IMAGE_SUFFIXES or content_type.split(";")[0].strip().lower().startswith("image/"):
            self._handle_downloaded_image(resource_url, resource_bytes, content_type)
            return

        messagebox.showerror(
            "Unsupported URL resource",
            (
                "This URL did not look like a supported document or image.\n\n"
                "Supported document types: CSV, XLSX, XLS, XLSM, PDF, DOCX, PPTX, TXT\n"
                "Supported image types: PNG, JPG, JPEG, GIF, BMP, WEBP"
            ),
        )
        self.image_status_var.set("The fetched URL was not a supported document or image type.")
        self.image_metric_var.set("No linked image")

    def _handle_downloaded_document(self, resource_url: str, resource_bytes: bytes, suffix: str) -> None:
        DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
        parsed_path = urllib.parse.urlparse(resource_url).path
        base_name = Path(parsed_path).stem or "downloaded_resource"
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", base_name).strip("._") or "downloaded_resource"
        destination = DOWNLOADS_DIR / f"{safe_name}{suffix}"
        destination.write_bytes(resource_bytes)

        self.reference_image_bytes = None
        self.reference_image_name = ""
        self.reference_image_url = ""
        self.reference_image_preview = None
        self.image_preview_label.configure(
            image="",
            text=f"Fetched document URL:\n{destination.name}\n\nThis link was loaded as data instead of an image preview.",
        )
        self.image_metric_var.set("URL file")
        self.file_path_var.set(str(destination))
        self.status_var.set(f"Fetched file from URL and saved it as {destination.name}.")
        self.image_status_var.set("The URL returned a supported document, and it has been loaded into the app.")
        self.note_var.set("URL-loaded files are stored in a temporary download folder so they can be parsed like local files.")
        self._load_selected_file()

    def _handle_downloaded_image(self, image_url: str, image_bytes: bytes, content_type: str) -> None:
        pil_image_module = self._optional_module("PIL.Image")
        pil_imagetk_module = self._optional_module("PIL.ImageTk")
        if pil_image_module is None or pil_imagetk_module is None:
            messagebox.showerror("Image preview unavailable", _missing_dependency_message("pillow"))
            self.image_metric_var.set("No linked image")
            return

        try:
            image = pil_image_module.open(io.BytesIO(image_bytes))
            resampling_holder = getattr(pil_image_module, "Resampling", pil_image_module)
            image.thumbnail((320, 220), getattr(resampling_holder, "LANCZOS", 1))
            photo = pil_imagetk_module.PhotoImage(image)
        except Exception as exc:
            messagebox.showerror("Invalid image", f"The provided link did not return a usable image.\n\n{exc}")
            self.image_status_var.set("The provided URL did not return a valid image file.")
            self.image_metric_var.set("No linked image")
            return

        parsed_path = urllib.parse.urlparse(image_url).path
        image_name = Path(parsed_path).name or "linked_image"
        if content_type.startswith("image/"):
            self.image_status_var.set(f"Fetched image successfully from link ({content_type}).")
        else:
            self.image_status_var.set("Fetched file from link and loaded it as an image preview.")

        self.reference_image_bytes = image_bytes
        self.reference_image_name = image_name
        self.reference_image_url = image_url
        self.reference_image_preview = photo
        self.image_preview_label.configure(image=photo, text="")
        self.image_metric_var.set("Image ready")
        self.note_var.set("The linked image can be previewed in the app and included in the exported PowerPoint.")

    def _load_selected_file(self) -> None:
        raw_path = self.file_path_var.get().strip()
        if not raw_path:
            messagebox.showinfo("Choose a file", "Please choose a file first.")
            return

        file_path = Path(raw_path)
        if not file_path.exists():
            messagebox.showerror("File not found", "The selected file does not exist.")
            return

        try:
            self.datasets = load_datasets(file_path)
        except Exception as exc:
            messagebox.showerror("Unable to load file", str(exc))
            self.status_var.set("The selected file could not be converted into chart-ready data.")
            self.note_var.set("Try a file that contains tables, labels with values, or a cleaner numeric layout.")
            return

        dataset_names = [dataset.name for dataset in self.datasets]
        self.dataset_box.configure(values=dataset_names)
        self.dataset_var.set(dataset_names[0])
        self._refresh_value_columns()
        self.dataset_metric_var.set(f"{len(self.datasets)} loaded")
        self.status_var.set(f"Loaded {len(self.datasets)} graph-ready dataset(s) from {file_path.name}.")
        self.note_var.set(self.datasets[0].source_note)
        self._update_preview()
        self._generate_graph()

    def _get_selected_dataset(self) -> DatasetChoice | None:
        selected_name = self.dataset_var.get()
        for dataset in self.datasets:
            if dataset.name == selected_name:
                return dataset
        return self.datasets[0] if self.datasets else None

    def _on_dataset_changed(self, _event: object | None = None) -> None:
        self._refresh_value_columns()
        self._update_preview()

    def _on_value_changed(self, _event: object | None = None) -> None:
        self._update_preview()

    def _refresh_value_columns(self) -> None:
        dataset = self._get_selected_dataset()
        if not dataset:
            self.value_box.configure(values=[])
            self.value_column_var.set("")
            return

        self.value_box.configure(values=dataset.numeric_columns)
        self.value_column_var.set(dataset.numeric_columns[0])
        self.note_var.set(dataset.source_note)
        self._update_preview()

    def _update_preview(self) -> None:
        for item in self.preview.get_children():
            self.preview.delete(item)

        dataset = self._get_selected_dataset()
        value_column = self.value_column_var.get()
        if not dataset or not value_column:
            return

        preview_frame = dataset.dataframe[[dataset.category_column, value_column]].head(18).copy()
        for index, (_, row) in enumerate(preview_frame.iterrows(), start=1):
            self.preview.insert(
                "",
                "end",
                values=(
                    index,
                    str(row[dataset.category_column]),
                    self._format_number(row[value_column]),
                ),
            )

    def _generate_graph(self) -> None:
        dataset = self._get_selected_dataset()
        if not dataset:
            messagebox.showinfo("Load data", "Please load a document first.")
            return

        value_column = self.value_column_var.get()
        if not value_column:
            messagebox.showinfo("Choose a value", "Please choose a numeric column to graph.")
            return

        chart_type = self.chart_type_var.get()
        chart_frame = dataset.dataframe[[dataset.category_column, value_column]].dropna().copy()
        if chart_frame.empty:
            messagebox.showerror("No data", "The selected dataset does not contain usable values.")
            return

        chart_frame[dataset.category_column] = chart_frame[dataset.category_column].astype(str).str.strip()
        chart_frame = chart_frame[chart_frame[dataset.category_column] != ""]
        if chart_frame.empty:
            messagebox.showerror("No categories", "The selected dataset does not contain usable topic labels.")
            return

        limited_frame, limit_note = self._limit_rows_for_chart(chart_frame, chart_type)

        figure = Figure(figsize=(8.8, 5.8), dpi=110, facecolor="#fffaf2")
        axis = figure.add_subplot(111)
        axis.set_facecolor("#fffdf8")

        categories = limited_frame[dataset.category_column].tolist()
        values = limited_frame[value_column].tolist()
        numbered_categories = [f"{index}. {self._trim_label(label)}" for index, label in enumerate(categories, start=1)]
        self.figure_counter += 1
        title = f"Figure {self.figure_counter}. {value_column} ({chart_type})"
        subtitle = f"Topic: {dataset.name}"

        if chart_type == "Bar Chart":
            bars = axis.bar(range(len(values)), values, color="#1f6f63", edgecolor="#0e3c35", linewidth=1)
            axis.set_xticks(range(len(numbered_categories)))
            axis.set_xticklabels(numbered_categories, rotation=35, ha="right")
            self._annotate_bars(axis, bars, values)
        elif chart_type == "Horizontal Bar Chart":
            bars = axis.barh(range(len(values)), values, color="#c16e28", edgecolor="#6f3d14", linewidth=1)
            axis.set_yticks(range(len(numbered_categories)))
            axis.set_yticklabels(numbered_categories)
            self._annotate_horizontal_bars(axis, bars, values)
        elif chart_type == "Line Chart":
            axis.plot(range(len(values)), values, color="#1f6f63", marker="o", linewidth=2.5)
            axis.set_xticks(range(len(numbered_categories)))
            axis.set_xticklabels(numbered_categories, rotation=35, ha="right")
            self._annotate_points(axis, values)
        elif chart_type == "Area Chart":
            axis.plot(range(len(values)), values, color="#0b7285", linewidth=2.2)
            axis.fill_between(range(len(values)), values, color="#66c2d4", alpha=0.45)
            axis.set_xticks(range(len(numbered_categories)))
            axis.set_xticklabels(numbered_categories, rotation=35, ha="right")
            self._annotate_points(axis, values)
        elif chart_type == "Scatter Plot":
            axis.scatter(range(len(values)), values, s=90, color="#8a4fff", edgecolors="#381f63", linewidths=0.9)
            axis.set_xticks(range(len(numbered_categories)))
            axis.set_xticklabels(numbered_categories, rotation=35, ha="right")
            self._annotate_points(axis, values)
        elif chart_type == "Pie Chart":
            if any(value <= 0 for value in values):
                messagebox.showerror("Pie chart unavailable", "Pie charts need values greater than zero.")
                return
            wedges, texts, autotexts = axis.pie(
                values,
                labels=numbered_categories,
                autopct="%1.1f%%",
                startangle=90,
                wedgeprops={"linewidth": 1, "edgecolor": "#fffaf2"},
            )
            for autotext in autotexts:
                autotext.set_color("#123b35")
                autotext.set_fontsize(9)
            axis.axis("equal")
        elif chart_type == "Histogram":
            counts, bins, patches = axis.hist(values, bins=min(8, max(4, len(values) // 2)), color="#4d7c0f", edgecolor="#35590d")
            for index, patch in enumerate(patches, start=1):
                height = patch.get_height()
                axis.annotate(
                    f"Bin {index}\n{int(height)}",
                    (patch.get_x() + patch.get_width() / 2, height),
                    textcoords="offset points",
                    xytext=(0, 8),
                    ha="center",
                    fontsize=8,
                    color="#123b35",
                )
            axis.set_xlabel(value_column)
            axis.set_ylabel("Frequency")
        else:
            messagebox.showerror("Unsupported graph", f"The graph type '{chart_type}' is not available.")
            return

        if chart_type != "Histogram":
            axis.set_xlabel("Numbered Topics")
            axis.set_ylabel(value_column)

        if chart_type not in {"Pie Chart", "Histogram"}:
            axis.grid(axis="y", alpha=0.22, linestyle="--")
        elif chart_type == "Histogram":
            axis.grid(axis="y", alpha=0.2, linestyle="--")

        axis.set_title(title, fontsize=15, fontweight="bold", color="#123b35", pad=16)
        figure.text(0.5, 0.93, subtitle, ha="center", fontsize=10, color="#526a66")
        if limit_note:
            figure.text(0.5, 0.02, limit_note, ha="center", fontsize=9, color="#7a5d2b")

        figure.tight_layout(rect=(0.02, 0.05, 0.98, 0.9))
        self._show_figure(figure)
        self.current_chart_title = title
        self.current_chart_subtitle = subtitle
        self.current_chart_note = limit_note
        self.current_chart_frame = limited_frame.reset_index(drop=True).copy()
        self.current_value_column = value_column
        self.current_chart_type = chart_type
        self.graph_metric_var.set(chart_type)
        self.status_var.set(f"Generated a {chart_type.lower()} for '{value_column}' from '{dataset.name}'.")
        self.note_var.set("Numbering on the graph matches the preview table so the topic order stays easy to follow.")

    def _export_to_ppt(self) -> None:
        if self.current_figure is None or self.current_chart_frame.empty:
            messagebox.showinfo("Generate a graph", "Please generate a graph before exporting to PowerPoint.")
            return

        presentation_factory = _optional_attr("pptx", "Presentation")
        pptx_util = self._optional_module("pptx.util")
        if presentation_factory is None or pptx_util is None:
            messagebox.showerror(
                "PowerPoint export unavailable",
                _missing_dependency_message("python-pptx"),
            )
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=(("PowerPoint Presentation", "*.pptx"), ("All files", "*.*")),
            title="Save PowerPoint Presentation",
        )
        if not save_path:
            return

        presentation = presentation_factory()
        Inches = getattr(pptx_util, "Inches")
        Pt = getattr(pptx_util, "Pt")

        title_slide = presentation.slides.add_slide(presentation.slide_layouts[0])
        title_slide.shapes.title.text = f"{APP_NAME} Report"
        title_slide.placeholders[1].text = f"{self.current_chart_title}\n{self.current_chart_subtitle}"

        chart_slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        chart_slide.shapes.title.text = self.current_chart_title

        summary_box = chart_slide.shapes.add_textbox(Inches(0.6), Inches(0.7), Inches(5.6), Inches(1.1))
        summary_frame = summary_box.text_frame
        summary_frame.text = self.current_chart_subtitle
        summary_frame.paragraphs[0].font.size = Pt(18)
        summary_frame.paragraphs[0].font.bold = True
        detail_paragraph = summary_frame.add_paragraph()
        detail_paragraph.text = f"Graph type: {self.current_chart_type} | Value column: {self.current_value_column}"
        detail_paragraph.font.size = Pt(11)
        if self.current_chart_note:
            note_paragraph = summary_frame.add_paragraph()
            note_paragraph.text = self.current_chart_note
            note_paragraph.font.size = Pt(10)

        image_stream = io.BytesIO()
        self.current_figure.savefig(
            image_stream,
            format="png",
            dpi=180,
            bbox_inches="tight",
            facecolor=self.current_figure.get_facecolor(),
        )
        image_stream.seek(0)
        chart_slide.shapes.add_picture(image_stream, Inches(0.6), Inches(1.8), width=Inches(8.2))

        summary_slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        summary_slide.shapes.title.text = "Top Data Points"
        bullet_box = summary_slide.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(8.4), Inches(5.2))
        bullet_frame = bullet_box.text_frame
        bullet_frame.word_wrap = True

        category_column = self.current_chart_frame.columns[0]
        preview_rows = self.current_chart_frame[[category_column, self.current_value_column]].head(10)
        for index, (_, row) in enumerate(preview_rows.iterrows(), start=1):
            paragraph = bullet_frame.paragraphs[0] if index == 1 else bullet_frame.add_paragraph()
            paragraph.text = (
                f"{index}. {self._trim_label(str(row[category_column]), width=70)}"
                f" - {self._format_number(row[self.current_value_column])}"
            )
            paragraph.font.size = Pt(20 if index == 1 else 18)

        if self.reference_image_bytes:
            image_slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            image_slide.shapes.title.text = "Linked Image Reference"
            caption_box = image_slide.shapes.add_textbox(Inches(0.6), Inches(0.75), Inches(8.4), Inches(0.8))
            caption_frame = caption_box.text_frame
            caption_frame.text = self.reference_image_name or "Linked image"
            caption_frame.paragraphs[0].font.size = Pt(16)
            if self.reference_image_url:
                source_line = caption_frame.add_paragraph()
                source_line.text = f"Source: {self.reference_image_url}"
                source_line.font.size = Pt(10)

            linked_image_stream = io.BytesIO(self.reference_image_bytes)
            image_slide.shapes.add_picture(linked_image_stream, Inches(1.0), Inches(1.6), width=Inches(7.4))

        presentation.save(save_path)
        self.status_var.set(f"Exported the current graph to {Path(save_path).name}.")
        if self.reference_image_bytes:
            self.note_var.set("The PowerPoint includes the graph, a data summary slide, and the linked image as a reference slide.")
        else:
            self.note_var.set("The PowerPoint includes a title slide, a chart slide, and a short data summary slide.")

    def _limit_rows_for_chart(self, frame: pd.DataFrame, chart_type: str) -> tuple[pd.DataFrame, str]:
        if chart_type == "Pie Chart" and len(frame) > 10:
            limited = frame.nlargest(10, frame.columns[1]).copy()
            return limited, "Showing the top 10 values so the pie chart stays readable."
        if chart_type in {"Bar Chart", "Horizontal Bar Chart", "Line Chart", "Area Chart", "Scatter Plot"} and len(frame) > 20:
            return frame.head(20).copy(), "Showing the first 20 items so the labels and numbering stay clear."
        return frame.copy(), ""

    def _render_placeholder_figure(self) -> None:
        figure = Figure(figsize=(8.8, 5.8), dpi=110, facecolor="#fffaf2")
        axis = figure.add_subplot(111)
        axis.set_facecolor("#fffdf8")
        axis.text(
            0.5,
            0.6,
            APP_NAME,
            ha="center",
            va="center",
            fontsize=18,
            fontweight="bold",
            color="#123b35",
            transform=axis.transAxes,
        )
        axis.text(
            0.5,
            0.44,
            "Load a document, build a presentation-ready graph, and export a polished PowerPoint.",
            ha="center",
            va="center",
            fontsize=10,
            color="#526a66",
            transform=axis.transAxes,
        )
        axis.axis("off")
        self._show_figure(figure)

    def _show_figure(self, figure: Figure) -> None:
        self.current_figure = figure
        for child in self.chart_host.winfo_children():
            child.destroy()
        canvas = FigureCanvasTkAgg(figure, master=self.chart_host)
        canvas.draw()
        canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")

    @staticmethod
    def _optional_module(module_name: str) -> object | None:
        try:
            return importlib.import_module(module_name)
        except ImportError:
            return None

    @staticmethod
    def _format_number(value: object) -> str:
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return str(value)
        if numeric.is_integer():
            return f"{int(numeric):,}"
        return f"{numeric:,.2f}"

    @staticmethod
    def _trim_label(label: str, width: int = 24) -> str:
        compact = re.sub(r"\s+", " ", str(label)).strip()
        return textwrap.shorten(compact, width=width, placeholder="...")

    def _annotate_bars(self, axis: plt.Axes, bars: object, values: list[float]) -> None:
        for index, (bar, value) in enumerate(zip(bars, values), start=1):
            axis.annotate(
                f"#{index}\n{self._format_number(value)}",
                (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                textcoords="offset points",
                xytext=(0, 8),
                ha="center",
                fontsize=8,
                color="#123b35",
            )

    def _annotate_horizontal_bars(self, axis: plt.Axes, bars: object, values: list[float]) -> None:
        for index, (bar, value) in enumerate(zip(bars, values), start=1):
            axis.annotate(
                f"#{index}  {self._format_number(value)}",
                (bar.get_width(), bar.get_y() + bar.get_height() / 2),
                textcoords="offset points",
                xytext=(8, 0),
                va="center",
                fontsize=8,
                color="#123b35",
            )

    def _annotate_points(self, axis: plt.Axes, values: list[float]) -> None:
        for index, value in enumerate(values, start=1):
            axis.annotate(
                f"#{index}: {self._format_number(value)}",
                (index - 1, value),
                textcoords="offset points",
                xytext=(0, 10),
                ha="center",
                fontsize=8,
                color="#123b35",
            )


def main() -> None:
    try:
        root = tk.Tk()
        app = DocuLoomPrismApp(root)
        root.mainloop()
    except Exception as exc:
        error_details = traceback.format_exc()
        try:
            messagebox.showerror(
                f"{APP_NAME} Startup Error",
                f"The app crashed while starting.\n\n{exc}\n\nFull details:\n{error_details}",
            )
        except Exception:
            print(error_details)
        raise


if __name__ == "__main__":
    main()
