from __future__ import annotations

import importlib
import math
import re
import textwrap
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

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
        raise RuntimeError("python-docx is not installed. Run: pip install python-docx")

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
        raise RuntimeError("python-pptx is not installed. Run: pip install python-pptx")

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
        raise RuntimeError("pypdf is not installed. Run: pip install pypdf")

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


class DocumentGraphStudio:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Document Graph Studio")
        self.root.geometry("1320x820")
        self.root.minsize(1120, 720)
        self.root.configure(bg="#f4efe6")

        self.file_path_var = tk.StringVar()
        self.chart_type_var = tk.StringVar(value=CHART_TYPES[0])
        self.dataset_var = tk.StringVar()
        self.value_column_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Choose a file to extract numbers and generate a graph.")
        self.note_var = tk.StringVar(
            value="Best with tables or documents that contain clear numeric values such as sales, scores, budgets, or counts."
        )

        self.datasets: list[DatasetChoice] = []
        self.current_figure: Figure | None = None
        self.figure_counter = 0

        self._configure_styles()
        self._build_layout()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure("App.TFrame", background="#f4efe6")
        style.configure("Panel.TFrame", background="#fffaf2")
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
            "PanelTitle.TLabel",
            background="#fffaf2",
            foreground="#123b35",
            font=("Segoe UI Semibold", 11),
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

    def _build_layout(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        header = ttk.Frame(outer, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Document Graph Studio", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Upload a PDF, Word file, Excel sheet, PowerPoint, CSV, or text file, then choose a chart from the dropdown."
            ),
            style="Body.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        control_panel = ttk.Frame(outer, style="Panel.TFrame", padding=18)
        control_panel.grid(row=1, column=0, sticky="ew", pady=(16, 14))
        for column in range(4):
            control_panel.columnconfigure(column, weight=1)

        ttk.Label(control_panel, text="Document File", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(control_panel, textvariable=self.file_path_var, font=("Segoe UI", 10)).grid(
            row=1, column=0, columnspan=3, sticky="ew", padx=(0, 12), pady=(6, 14)
        )
        ttk.Button(control_panel, text="Browse", command=self._browse_file).grid(row=1, column=3, sticky="ew", pady=(6, 14))

        ttk.Label(control_panel, text="Dataset", style="PanelTitle.TLabel").grid(row=2, column=0, sticky="w")
        self.dataset_box = ttk.Combobox(control_panel, textvariable=self.dataset_var, state="readonly", font=("Segoe UI", 10))
        self.dataset_box.grid(row=3, column=0, sticky="ew", padx=(0, 12), pady=(6, 14))
        self.dataset_box.bind("<<ComboboxSelected>>", self._on_dataset_changed)

        ttk.Label(control_panel, text="Value Column", style="PanelTitle.TLabel").grid(row=2, column=1, sticky="w")
        self.value_box = ttk.Combobox(control_panel, textvariable=self.value_column_var, state="readonly", font=("Segoe UI", 10))
        self.value_box.grid(row=3, column=1, sticky="ew", padx=(0, 12), pady=(6, 14))
        self.value_box.bind("<<ComboboxSelected>>", self._on_value_changed)

        ttk.Label(control_panel, text="Graph Type", style="PanelTitle.TLabel").grid(row=2, column=2, sticky="w")
        self.chart_type_box = ttk.Combobox(
            control_panel,
            textvariable=self.chart_type_var,
            state="readonly",
            values=CHART_TYPES,
            font=("Segoe UI", 10),
        )
        self.chart_type_box.grid(row=3, column=2, sticky="ew", padx=(0, 12), pady=(6, 14))

        actions = ttk.Frame(control_panel, style="Panel.TFrame")
        actions.grid(row=3, column=3, sticky="ew", pady=(6, 14))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)

        ttk.Button(actions, text="Load Data", style="Accent.TButton", command=self._load_selected_file).grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(actions, text="Generate Graph", style="Accent.TButton", command=self._generate_graph).grid(
            row=0, column=1, sticky="ew"
        )

        ttk.Label(control_panel, textvariable=self.status_var, style="Body.TLabel", wraplength=1160).grid(
            row=4, column=0, columnspan=4, sticky="w"
        )
        ttk.Label(control_panel, textvariable=self.note_var, style="Body.TLabel", wraplength=1160).grid(
            row=5, column=0, columnspan=4, sticky="w", pady=(6, 0)
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

        ttk.Label(preview_card, text="Detected Data Preview", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")

        preview_columns = ("row", "category", "value")
        self.preview = ttk.Treeview(preview_card, columns=preview_columns, show="headings", height=20)
        self.preview.heading("row", text="#")
        self.preview.heading("category", text="Topic / Category")
        self.preview.heading("value", text="Selected Value")
        self.preview.column("row", width=60, anchor="center")
        self.preview.column("category", width=210, anchor="w")
        self.preview.column("value", width=120, anchor="e")
        self.preview.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

        graph_card = ttk.Frame(content, style="Panel.TFrame", padding=18)
        graph_card.grid(row=0, column=1, sticky="nsew")
        graph_card.rowconfigure(1, weight=1)
        graph_card.columnconfigure(0, weight=1)

        ttk.Label(graph_card, text="Graph Output", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.chart_host = ttk.Frame(graph_card, style="Panel.TFrame")
        self.chart_host.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.chart_host.rowconfigure(0, weight=1)
        self.chart_host.columnconfigure(0, weight=1)

        self._render_placeholder_figure()

    def _browse_file(self) -> None:
        selected = filedialog.askopenfilename(filetypes=SUPPORTED_FILE_TYPES)
        if not selected:
            return
        self.file_path_var.set(selected)
        self._load_selected_file()

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
        self.status_var.set(f"Generated a {chart_type.lower()} for '{value_column}' from '{dataset.name}'.")
        self.note_var.set("Numbering on the graph matches the preview table so the topic order stays easy to follow.")

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
            "Load a file and generate a graph",
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
            "Supported: PDF, DOCX, XLSX/XLS, PPTX, CSV, TXT",
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
    root = tk.Tk()
    app = DocumentGraphStudio(root)
    root.mainloop()


if __name__ == "__main__":
    main()
