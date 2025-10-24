import json
import os
import re
import sys
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import fitz  # PyMuPDF
from PIL import Image, ImageTk
from PyPDF2 import PdfReader, PdfWriter
import webbrowser


COLUMNS = ["Financial", "Income", "Shares"]
DEFAULT_PATTERNS = {
    "Financial": ["statement of financial position"],
    "Income": ["statement of profit or loss"],
    "Shares": ["Movements in issued capital"],
}
YEAR_DEFAULT_PATTERNS = [r"(\d{4})\s+Annual\s+Report"]


@dataclass
class Match:
    page_index: int
    source: str
    pattern: Optional[str] = None


@dataclass
class PDFEntry:
    path: Path
    doc: fitz.Document
    matches: Dict[str, List[Match]] = field(default_factory=dict)
    current_index: Dict[str, Optional[int]] = field(default_factory=dict)
    year: str = ""

    def __post_init__(self) -> None:
        for column in COLUMNS:
            self.matches.setdefault(column, [])
            self.current_index.setdefault(column, 0 if self.matches[column] else None)

    @property
    def stem(self) -> str:
        return self.path.stem


class PDFCell:
    def __init__(self, parent: tk.Widget, app: "ReportApp", entry: PDFEntry, category: str):
        self.app = app
        self.entry = entry
        self.category = category
        self.frame = ttk.Frame(parent, padding=4)
        self.image_label = ttk.Label(self.frame)
        self.info_label = ttk.Label(self.frame)
        self.image_label.pack(expand=True, fill=tk.BOTH)
        self.info_label.pack(fill=tk.X)
        self.photo: Optional[ImageTk.PhotoImage] = None
        self._right_click_job: Optional[str] = None
        for widget in (self.image_label, self.info_label):
            widget.bind("<Button-1>", self._next_match)
            widget.bind("<Shift-Button-1>", self._previous_match)
            widget.bind("<Button-3>", self._on_right_click)
            widget.bind("<Double-Button-3>", self._manual_select)

    def _next_match(self, event: tk.Event) -> None:  # type: ignore[override]
        self.app.cycle_match(self.entry, self.category, forward=True)

    def _previous_match(self, event: tk.Event) -> None:  # type: ignore[override]
        self.app.cycle_match(self.entry, self.category, forward=False)

    def _on_right_click(self, event: tk.Event) -> None:  # type: ignore[override]
        if self._right_click_job is not None:
            self.frame.after_cancel(self._right_click_job)
            self._right_click_job = None

        def _open() -> None:
            self._right_click_job = None
            self.app.open_current_match(self.entry, self.category)

        self._right_click_job = self.frame.after(200, _open)

    def _manual_select(self, event: tk.Event) -> None:  # type: ignore[override]
        if self._right_click_job is not None:
            self.frame.after_cancel(self._right_click_job)
            self._right_click_job = None
        self.app.manual_select(self.entry, self.category)

    def update_display(self, photo: Optional[ImageTk.PhotoImage], info_text: str) -> None:
        self.photo = photo
        if photo is not None:
            self.image_label.configure(image=photo)
        else:
            self.image_label.configure(image="")
        self.info_label.configure(text=info_text)


class ReportApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Annual Report Analyst")
        self.folder_path = tk.StringVar()
        self.company_var = tk.StringVar()
        self.pattern_texts: Dict[str, tk.Text] = {}
        self.case_insensitive_vars: Dict[str, tk.BooleanVar] = {}
        self.whitespace_as_space_vars: Dict[str, tk.BooleanVar] = {}
        self.pdf_entries: List[PDFEntry] = []
        self.cells: Dict[tuple[Path, str], PDFCell] = {}
        self.year_vars: Dict[Path, tk.StringVar] = {}
        self.year_pattern_text: Optional[tk.Text] = None
        self.year_case_insensitive_var = tk.BooleanVar(master=self.root, value=True)
        self.year_whitespace_as_space_var = tk.BooleanVar(master=self.root, value=False)
        self.companies_dir = Path(__file__).resolve().parent / "companies"
        self.pattern_config_path = Path(__file__).resolve().parent / "pattern_config.json"
        self.config_data: Dict[str, Any] = {}
        self.last_company_preference: str = ""
        self._config_loaded = False

        self._build_ui()
        self._load_pattern_config()
        self._maximize_window()
        self.root.after(0, self._load_pdfs_on_start)

    def _build_ui(self) -> None:
        top_frame = ttk.Frame(self.root, padding=8)
        top_frame.pack(fill=tk.X)

        company_label = ttk.Label(top_frame, text="Company:")
        company_label.pack(side=tk.LEFT)

        self.company_combo = ttk.Combobox(top_frame, textvariable=self.company_var, state="readonly", width=30)
        self.company_combo.pack(side=tk.LEFT, padx=4)
        self.company_combo.bind("<<ComboboxSelected>>", self._on_company_selected)

        folder_entry = ttk.Entry(top_frame, textvariable=self.folder_path, width=60, state="readonly")
        folder_entry.pack(side=tk.LEFT, padx=4)

        load_button = ttk.Button(top_frame, text="Load PDFs", command=self.load_pdfs)
        load_button.pack(side=tk.LEFT, padx=4)

        self._refresh_company_options()

        pattern_frame = ttk.LabelFrame(self.root, text="Regex patterns (one per line)", padding=8)
        pattern_frame.pack(fill=tk.X, padx=8, pady=4)

        for idx, column in enumerate(COLUMNS):
            column_frame = ttk.Frame(pattern_frame)
            column_frame.grid(row=0, column=idx, padx=4, sticky="nsew")
            pattern_frame.columnconfigure(idx, weight=1)
            ttk.Label(column_frame, text=column).pack(anchor="w")
            text_widget = tk.Text(column_frame, height=4, width=30)
            text_widget.pack(fill=tk.BOTH, expand=True)
            text_widget.insert("1.0", "\n".join(DEFAULT_PATTERNS[column]))
            self.pattern_texts[column] = text_widget
            var = tk.BooleanVar(master=self.root, value=True)
            self.case_insensitive_vars[column] = var
            ttk.Checkbutton(column_frame, text="Case-insensitive", variable=var).pack(anchor="w", pady=(4, 0))
            whitespace_var = tk.BooleanVar(master=self.root, value=False)
            self.whitespace_as_space_vars[column] = whitespace_var
            ttk.Checkbutton(
                column_frame,
                text="Treat spaces as any whitespace",
                variable=whitespace_var,
            ).pack(anchor="w")

        update_button = ttk.Button(pattern_frame, text="Apply Patterns", command=self.apply_patterns)
        update_button.grid(row=1, column=0, columnspan=len(COLUMNS), pady=4)

        year_frame = ttk.Frame(pattern_frame)
        year_frame.grid(row=2, column=0, columnspan=len(COLUMNS), sticky="ew", pady=(8, 0))
        ttk.Label(year_frame, text="Year pattern").pack(anchor="w")
        year_text = tk.Text(year_frame, height=2, width=30)
        year_text.pack(fill=tk.BOTH, expand=True)
        year_text.insert("1.0", "\n".join(YEAR_DEFAULT_PATTERNS))
        self.year_pattern_text = year_text
        ttk.Checkbutton(year_frame, text="Case-insensitive", variable=self.year_case_insensitive_var).pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(
            year_frame,
            text="Treat spaces as any whitespace",
            variable=self.year_whitespace_as_space_var,
        ).pack(anchor="w")

        grid_container = ttk.Frame(self.root)
        grid_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        self.canvas = tk.Canvas(grid_container)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(grid_container, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        # Linux scroll events
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)

        self.inner_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.confirm_button = ttk.Button(self.root, text="Confirm", command=self.confirm_selections)
        self.confirm_button.pack(pady=8)

    def _on_frame_configure(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:  # type: ignore[override]
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _bind_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event: tk.Event) -> None:  # type: ignore[override]
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

    def _refresh_company_options(self) -> None:
        if not self.companies_dir.exists():
            self.company_combo.configure(values=[])
            return
        companies = sorted([d.name for d in self.companies_dir.iterdir() if d.is_dir()])
        self.company_combo.configure(values=companies)
        preferred = self.last_company_preference or self.company_var.get()
        if preferred and preferred in companies:
            self.company_combo.set(preferred)
            self.company_var.set(preferred)
            self._set_folder_from_company(preferred)
        elif companies and not self.company_var.get():
            self.company_combo.current(0)
            self._set_folder_from_company(companies[0])

    def _on_company_selected(self, _: tk.Event) -> None:  # type: ignore[override]
        company = self.company_var.get()
        self._set_folder_from_company(company)

    def _set_folder_from_company(self, company: str) -> None:
        if not company:
            self.folder_path.set("")
            return
        folder = self.companies_dir / company / "raw"
        self.folder_path.set(str(folder))
        if self._config_loaded:
            self._update_last_company(company)

    def _load_pdfs_on_start(self) -> None:
        if self.folder_path.get():
            self.load_pdfs()

    def apply_patterns(self) -> None:
        if not self.folder_path.get():
            messagebox.showinfo("Select Folder", "Please select a folder before applying patterns.")
            return
        if not self.pdf_entries:
            self.load_pdfs()
            return

        prev_patterns = {key: list(value) for key, value in self.config_data.get("patterns", {}).items()}
        prev_case = {key: bool(value) for key, value in self.config_data.get("case_insensitive", {}).items()}
        prev_whitespace = {key: bool(value) for key, value in self.config_data.get("space_as_whitespace", {}).items()}
        prev_year_patterns = list(self.config_data.get("year_patterns", YEAR_DEFAULT_PATTERNS))
        prev_year_case = bool(self.config_data.get("year_case_insensitive", True))
        prev_year_whitespace = bool(self.config_data.get("year_space_as_whitespace", False))

        pattern_map, year_patterns = self._gather_patterns()

        new_patterns = self.config_data.get("patterns", {})
        new_case = self.config_data.get("case_insensitive", {})
        new_whitespace = self.config_data.get("space_as_whitespace", {})

        changed_columns = set()
        for column in COLUMNS:
            old_patterns = prev_patterns.get(column, [])
            new_column_patterns = new_patterns.get(column, [])
            if old_patterns != new_column_patterns:
                changed_columns.add(column)
                continue
            old_case = prev_case.get(column, True)
            new_case_flag = bool(new_case.get(column, True))
            old_whitespace = prev_whitespace.get(column, False)
            new_whitespace_flag = bool(new_whitespace.get(column, False))
            if old_case != new_case_flag or old_whitespace != new_whitespace_flag:
                changed_columns.add(column)

        new_year_patterns = list(self.config_data.get("year_patterns", YEAR_DEFAULT_PATTERNS))
        new_year_case = bool(self.config_data.get("year_case_insensitive", True))
        new_year_whitespace = bool(self.config_data.get("year_space_as_whitespace", False))

        year_changed = (
            prev_year_patterns != new_year_patterns
            or prev_year_case != new_year_case
            or prev_year_whitespace != new_year_whitespace
        )

        if not changed_columns and not year_changed:
            return

        self._rescan_entries(pattern_map, year_patterns, changed_columns, year_changed)

    def load_pdfs(self) -> None:
        folder = self.folder_path.get()
        if not folder:
            messagebox.showinfo("Select Folder", "Please select a folder containing PDFs.")
            return

        folder_path = Path(folder)
        if not folder_path.exists():
            messagebox.showerror("Folder Not Found", f"The folder '{folder}' does not exist.")
            return

        self._clear_entries()
        pattern_map, year_patterns = self._gather_patterns()
        pdf_paths = sorted(folder_path.rglob("*.pdf"))
        if not pdf_paths:
            messagebox.showinfo("No PDFs", "No PDF files were found in the selected folder.")
            self._rebuild_grid()
            return

        for pdf_path in pdf_paths:
            try:
                doc = fitz.open(pdf_path)
            except Exception as exc:  # pragma: no cover - guard for invalid PDFs
                messagebox.showwarning("PDF Error", f"Could not open '{pdf_path}': {exc}")
                continue

            matches: Dict[str, List[Match]] = {column: [] for column in COLUMNS}
            year_value = ""
            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                page_text = page.get_text("text")
                for column, patterns in pattern_map.items():
                    for pattern in patterns:
                        if pattern.search(page_text):
                            matches[column].append(Match(page_index=page_index, source="regex", pattern=pattern.pattern))
                            break
                if not year_value:
                    for pattern in year_patterns:
                        year_match = pattern.search(page_text)
                        if year_match:
                            if year_match.groups():
                                year_value = year_match.group(1)
                            else:
                                year_value = year_match.group(0)
                            break

            entry = PDFEntry(path=pdf_path, doc=doc, matches=matches, year=year_value)
            # Reset current indices based on available matches
            for column in COLUMNS:
                entry.current_index[column] = 0 if entry.matches[column] else None
            self.pdf_entries.append(entry)

        self._rebuild_grid()

    def _clear_entries(self) -> None:
        for entry in self.pdf_entries:
            try:
                entry.doc.close()
            except Exception:
                pass
        self.pdf_entries.clear()
        self.cells.clear()
        self.year_vars.clear()
        for child in self.inner_frame.winfo_children():
            child.destroy()

    def _rebuild_grid(self) -> None:
        for child in self.inner_frame.winfo_children():
            child.destroy()
        self.cells.clear()

        header = ttk.Frame(self.inner_frame)
        header.grid(row=0, column=0, columnspan=4, sticky="ew")
        ttk.Label(header, text="PDF", width=30, anchor="w").grid(row=0, column=0, padx=4)
        for idx, column in enumerate(COLUMNS, start=1):
            ttk.Label(header, text=column, anchor="center").grid(row=0, column=idx, padx=4)

        for row_index, entry in enumerate(self.pdf_entries, start=1):
            relative_path = entry.path.relative_to(Path(self.folder_path.get())) if self.folder_path.get() else entry.path.name
            info_frame = ttk.Frame(self.inner_frame)
            info_frame.grid(row=row_index, column=0, sticky="nw", padx=4, pady=4)
            ttk.Label(info_frame, text=relative_path, anchor="w", width=30, wraplength=200).pack(anchor="w")
            year_var = tk.StringVar(value=entry.year)
            self.year_vars[entry.path] = year_var
            year_var.trace_add("write", lambda *_args, e=entry, v=year_var: setattr(e, "year", v.get()))
            year_entry = ttk.Entry(info_frame, textvariable=year_var, width=30)
            year_entry.pack(fill=tk.X, pady=(4, 0))
            for col_index, column in enumerate(COLUMNS, start=1):
                cell = PDFCell(self.inner_frame, self, entry, column)
                cell.frame.grid(row=row_index, column=col_index, padx=4, pady=4, sticky="nsew")
                self.cells[(entry.path, column)] = cell
                self._update_cell(entry, column)

        for idx in range(len(COLUMNS) + 1):
            self.inner_frame.columnconfigure(idx, weight=1)

    def _update_cell(self, entry: PDFEntry, category: str) -> None:
        cell = self.cells.get((entry.path, category))
        if cell is None:
            return

        matches = entry.matches[category]
        current_index = entry.current_index.get(category)
        if current_index is None or not matches:
            cell.update_display(None, "No matches found")
            return

        current_index = max(0, min(current_index, len(matches) - 1))
        entry.current_index[category] = current_index
        match = matches[current_index]
        total_width = self.canvas.winfo_width()
        if total_width <= 1:
            self.root.update_idletasks()
            total_width = self.canvas.winfo_width()
        if total_width <= 1:
            total_width = self.root.winfo_width()
        if total_width <= 1:
            total_width = 900
        target_width = max(50, total_width // len(COLUMNS))
        photo = self._render_page(entry.doc, match.page_index, target_width)
        info_parts = [f"Match {current_index + 1}/{len(matches)}", f"Page {match.page_index + 1}"]
        if match.source == "manual":
            info_parts.append("(manual)")
        cell.update_display(photo, " | ".join(info_parts))

    def _rescan_entries(
        self,
        pattern_map: Dict[str, List[re.Pattern[str]]],
        year_patterns: List[re.Pattern[str]],
        columns: set[str],
        year_changed: bool,
    ) -> None:
        for entry in self.pdf_entries:
            previous_pages: Dict[str, Optional[int]] = {}
            manual_matches: Dict[str, List[Match]] = {}
            new_matches: Dict[str, List[Match]] = {column: [] for column in columns}

            for column in columns:
                existing_matches = entry.matches.get(column, [])
                manual_matches[column] = [m for m in existing_matches if m.source == "manual"]
                current_index = entry.current_index.get(column)
                if current_index is not None and 0 <= current_index < len(existing_matches):
                    previous_pages[column] = existing_matches[current_index].page_index
                else:
                    previous_pages[column] = None

            detected_year: Optional[str] = None

            for page_index in range(len(entry.doc)):
                page = entry.doc.load_page(page_index)
                page_text = page.get_text("text")

                if year_changed and detected_year is None:
                    for pattern in year_patterns:
                        year_match = pattern.search(page_text)
                        if year_match:
                            if year_match.groups():
                                detected_year = year_match.group(1)
                            else:
                                detected_year = year_match.group(0)
                            break

                for column in columns:
                    compiled_patterns = pattern_map.get(column, [])
                    for pattern in compiled_patterns:
                        if pattern.search(page_text):
                            new_matches[column].append(
                                Match(page_index=page_index, source="regex", pattern=pattern.pattern)
                            )
                            break

            for column in columns:
                matches = new_matches[column]
                manual = manual_matches[column]
                manual_pages = {match.page_index for match in matches}
                for manual_match in manual:
                    if manual_match.page_index not in manual_pages:
                        matches.append(manual_match)
                entry.matches[column] = matches

                if matches:
                    target_page = previous_pages[column]
                    if target_page is not None:
                        for idx, match in enumerate(matches):
                            if match.page_index == target_page:
                                entry.current_index[column] = idx
                                break
                        else:
                            entry.current_index[column] = 0
                    else:
                        entry.current_index[column] = 0
                else:
                    entry.current_index[column] = None

                self._update_cell(entry, column)

            if year_changed:
                if detected_year is not None:
                    entry.year = detected_year
                    year_var = self.year_vars.get(entry.path)
                    if year_var is not None and year_var.get() != detected_year:
                        year_var.set(detected_year)
                else:
                    year_var = self.year_vars.get(entry.path)
                    if year_var is not None and year_var.get() != entry.year:
                        year_var.set(entry.year)

    def _render_page(self, doc: fitz.Document, page_index: int, target_width: int) -> Optional[ImageTk.PhotoImage]:
        try:
            page = doc.load_page(page_index)
            zoom_matrix = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom_matrix)
            mode = "RGBA" if pix.alpha else "RGB"
            image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            if target_width > 0 and image.width != target_width:
                ratio = target_width / image.width
                new_size = (int(image.width * ratio), int(image.height * ratio))
                image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(image)
        except Exception as exc:  # pragma: no cover - guard for rendering issues
            messagebox.showwarning("Render Error", f"Could not render page {page_index + 1}: {exc}")
            return None

    def cycle_match(self, entry: PDFEntry, category: str, *, forward: bool) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            messagebox.showinfo("No Matches", f"No matches available for {category} in {entry.path.name}.")
            return

        current_index = entry.current_index.get(category) or 0
        if forward:
            if current_index + 1 >= len(matches):
                messagebox.showinfo("End of Matches", "Reached the last matched page.")
                return
            entry.current_index[category] = current_index + 1
        else:
            if current_index - 1 < 0:
                messagebox.showinfo("Start of Matches", "Reached the first matched page.")
                return
            entry.current_index[category] = current_index - 1
        self._update_cell(entry, category)

    def manual_select(self, entry: PDFEntry, category: str) -> None:
        pdf_path = entry.path
        self._open_pdf(pdf_path)
        page_number = simpledialog.askinteger(
            "Manual Page Selection",
            f"Enter page number for {category} in {pdf_path.name}:",
            parent=self.root,
            minvalue=1,
            maxvalue=len(entry.doc),
        )
        if page_number is None:
            return

        page_index = page_number - 1
        matches = entry.matches[category]
        for idx, match in enumerate(matches):
            if match.page_index == page_index:
                entry.current_index[category] = idx
                self._update_cell(entry, category)
                return

        matches.append(Match(page_index=page_index, source="manual"))
        entry.current_index[category] = len(matches) - 1
        self._update_cell(entry, category)

    def open_current_match(self, entry: PDFEntry, category: str) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            messagebox.showinfo("No Matches", f"No matches available for {category} in {entry.path.name}.")
            return
        index = entry.current_index.get(category) or 0
        index = max(0, min(index, len(matches) - 1))
        page_index = matches[index].page_index
        self._open_pdf(entry.path, page_index)

    def _open_pdf(self, pdf_path: Path, page_index: Optional[int] = None) -> None:
        try:
            if page_index is not None:
                url = pdf_path.resolve().as_uri() + f"#page={page_index + 1}"
                opened = webbrowser.open(url)
                if opened:
                    return
            if sys.platform.startswith("win"):
                os.startfile(pdf_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f"open '{pdf_path}'")
            else:
                os.system(f"xdg-open '{pdf_path}' >/dev/null 2>&1 &")
        except Exception as exc:
            messagebox.showwarning("Open PDF", f"Could not open PDF: {exc}")

    def _gather_patterns(self) -> Tuple[Dict[str, List[re.Pattern[str]]], List[re.Pattern[str]]]:
        pattern_map: Dict[str, List[re.Pattern[str]]] = {}
        raw_map: Dict[str, List[str]] = {}
        case_flags: Dict[str, bool] = {}
        whitespace_flags: Dict[str, bool] = {}
        for column in COLUMNS:
            text_widget = self.pattern_texts[column]
            raw_text = text_widget.get("1.0", tk.END)
            patterns = [line.strip() for line in raw_text.splitlines() if line.strip()]
            if not patterns:
                patterns = DEFAULT_PATTERNS[column]
                text_widget.delete("1.0", tk.END)
                text_widget.insert("1.0", "\n".join(patterns))
            compiled = []
            flags = re.IGNORECASE if self.case_insensitive_vars[column].get() else 0
            for pattern in patterns:
                try:
                    compiled_pattern = self._apply_whitespace_option(pattern, self.whitespace_as_space_vars[column].get())
                    compiled.append(re.compile(compiled_pattern, flags))
                except re.error as exc:
                    messagebox.showerror("Invalid Pattern", f"Invalid regex '{pattern}' for {column}: {exc}")
            pattern_map[column] = compiled
            raw_map[column] = patterns
            case_flags[column] = self.case_insensitive_vars[column].get()
            whitespace_flags[column] = self.whitespace_as_space_vars[column].get()

        year_patterns: List[str] = []
        year_compiled: List[re.Pattern[str]] = []
        if self.year_pattern_text is not None:
            raw_text = self.year_pattern_text.get("1.0", tk.END)
            year_patterns = [line.strip() for line in raw_text.splitlines() if line.strip()]
        if not year_patterns:
            year_patterns = YEAR_DEFAULT_PATTERNS
            if self.year_pattern_text is not None:
                self.year_pattern_text.delete("1.0", tk.END)
                self.year_pattern_text.insert("1.0", "\n".join(year_patterns))
        year_flags = re.IGNORECASE if self.year_case_insensitive_var.get() else 0
        for pattern in year_patterns:
            try:
                compiled_pattern = self._apply_whitespace_option(pattern, self.year_whitespace_as_space_var.get())
                year_compiled.append(re.compile(compiled_pattern, year_flags))
            except re.error as exc:
                messagebox.showerror("Invalid Pattern", f"Invalid year regex '{pattern}': {exc}")
        self._save_pattern_config(
            raw_map,
            year_patterns,
            case_flags,
            self.year_case_insensitive_var.get(),
            whitespace_flags,
            self.year_whitespace_as_space_var.get(),
        )
        return pattern_map, year_compiled

    def _save_pattern_config(
        self,
        patterns: Dict[str, List[str]],
        year_patterns: List[str],
        case_flags: Dict[str, bool],
        year_case_flag: bool,
        whitespace_flags: Dict[str, bool],
        year_whitespace_flag: bool,
    ) -> None:
        self.config_data.update(
            {
                "patterns": patterns,
                "case_insensitive": case_flags,
                "year_patterns": year_patterns,
                "year_case_insensitive": year_case_flag,
                "space_as_whitespace": whitespace_flags,
                "year_space_as_whitespace": year_whitespace_flag,
            }
        )
        current_company = self.company_var.get()
        if current_company:
            self.config_data["last_company"] = current_company
        self._write_config()

    def _load_pattern_config(self) -> None:
        if not self.pattern_config_path.exists():
            self._config_loaded = True
            return
        try:
            with self.pattern_config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
                self.config_data = data
        except Exception as exc:  # pragma: no cover - guard for IO issues
            messagebox.showwarning("Load Patterns", f"Could not load pattern configuration: {exc}")
            self._config_loaded = True
            return

        patterns = data.get("patterns", {})
        case_flags = data.get("case_insensitive", {})
        whitespace_flags = data.get("space_as_whitespace", {})
        for column, text_widget in self.pattern_texts.items():
            column_patterns = patterns.get(column)
            if column_patterns:
                text_widget.delete("1.0", tk.END)
                text_widget.insert("1.0", "\n".join(column_patterns))
            if column in case_flags and column in self.case_insensitive_vars:
                self.case_insensitive_vars[column].set(bool(case_flags[column]))
            if column in whitespace_flags and column in self.whitespace_as_space_vars:
                self.whitespace_as_space_vars[column].set(bool(whitespace_flags[column]))

        year_patterns = data.get("year_patterns")
        if year_patterns and self.year_pattern_text is not None:
            self.year_pattern_text.delete("1.0", tk.END)
            self.year_pattern_text.insert("1.0", "\n".join(year_patterns))
        if "year_case_insensitive" in data:
            self.year_case_insensitive_var.set(bool(data["year_case_insensitive"]))
        if "year_space_as_whitespace" in data:
            self.year_whitespace_as_space_var.set(bool(data["year_space_as_whitespace"]))
        self.last_company_preference = data.get("last_company", "")
        self._apply_last_company_selection()
        self._config_loaded = True

    def _apply_whitespace_option(self, pattern: str, enabled: bool) -> str:
        if not enabled:
            return pattern
        return pattern.replace(" ", r"\s+")

    def _maximize_window(self) -> None:
        try:
            self.root.state("zoomed")
            return
        except tk.TclError:
            pass
        try:
            self.root.attributes("-zoomed", True)
            return
        except tk.TclError:
            pass
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width}x{screen_height}+0+0")

    def _apply_last_company_selection(self) -> None:
        if not self.last_company_preference:
            return
        companies = list(self.company_combo["values"])
        if self.last_company_preference in companies:
            self.company_combo.set(self.last_company_preference)
            self.company_var.set(self.last_company_preference)
            self._set_folder_from_company(self.last_company_preference)

    def _update_last_company(self, company: str) -> None:
        if not company:
            return
        if self.config_data.get("last_company") == company:
            return
        self.config_data["last_company"] = company
        self.last_company_preference = company
        self._write_config()

    def _write_config(self) -> None:
        try:
            with self.pattern_config_path.open("w", encoding="utf-8") as fh:
                json.dump(self.config_data, fh, indent=2)
        except Exception as exc:  # pragma: no cover - guard for IO issues
            messagebox.showwarning("Save Patterns", f"Could not save pattern configuration: {exc}")

    def confirm_selections(self) -> None:
        if not self.pdf_entries:
            messagebox.showinfo("No PDFs", "Load PDFs before confirming selections.")
            return

        folder = Path(self.folder_path.get())
        output_dir = folder / "cut"
        output_dir.mkdir(parents=True, exist_ok=True)

        for entry in self.pdf_entries:
            reader = PdfReader(str(entry.path))
            for category in COLUMNS:
                index = entry.current_index.get(category)
                matches = entry.matches.get(category, [])
                if index is None or not matches:
                    continue
                index = max(0, min(index, len(matches) - 1))
                match = matches[index]
                if match.page_index >= len(reader.pages):
                    continue
                writer = PdfWriter()
                writer.add_page(reader.pages[match.page_index])
                output_path = output_dir / f"{entry.stem}_{category}.pdf"
                with output_path.open("wb") as fh:
                    writer.write(fh)

        messagebox.showinfo("Pages Saved", f"Selected pages have been saved to '{output_dir}'.")


def main() -> None:
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
