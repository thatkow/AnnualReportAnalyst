import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import fitz  # PyMuPDF
from PIL import Image, ImageTk
from PyPDF2 import PdfReader, PdfWriter


COLUMNS = ["Financial", "Income", "Shares"]
DEFAULT_PATTERNS = {
    "Financial": ["statement of financial position"],
    "Income": ["statement of profit or loss"],
    "Shares": ["Movements in issued capital"],
}


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
        for widget in (self.image_label, self.info_label):
            widget.bind("<Button-1>", self._next_match)
            widget.bind("<Shift-Button-1>", self._previous_match)
            widget.bind("<Button-3>", self._manual_select)

    def _next_match(self, event: tk.Event) -> None:  # type: ignore[override]
        self.app.cycle_match(self.entry, self.category, forward=True)

    def _previous_match(self, event: tk.Event) -> None:  # type: ignore[override]
        self.app.cycle_match(self.entry, self.category, forward=False)

    def _manual_select(self, event: tk.Event) -> None:  # type: ignore[override]
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
        self.pattern_texts: Dict[str, tk.Text] = {}
        self.pdf_entries: List[PDFEntry] = []
        self.cells: Dict[tuple[PDFEntry, str], PDFCell] = {}

        self._build_ui()

    def _build_ui(self) -> None:
        top_frame = ttk.Frame(self.root, padding=8)
        top_frame.pack(fill=tk.X)

        folder_label = ttk.Label(top_frame, text="Folder:")
        folder_label.pack(side=tk.LEFT)
        folder_entry = ttk.Entry(top_frame, textvariable=self.folder_path, width=60)
        folder_entry.pack(side=tk.LEFT, padx=4)

        browse_button = ttk.Button(top_frame, text="Browse", command=self.select_folder)
        browse_button.pack(side=tk.LEFT)

        load_button = ttk.Button(top_frame, text="Load PDFs", command=self.load_pdfs)
        load_button.pack(side=tk.LEFT, padx=4)

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

        update_button = ttk.Button(pattern_frame, text="Apply Patterns", command=self.apply_patterns)
        update_button.grid(row=1, column=0, columnspan=len(COLUMNS), pady=4)

        grid_container = ttk.Frame(self.root)
        grid_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        self.canvas = tk.Canvas(grid_container)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(grid_container, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)

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

    def select_folder(self) -> None:
        selected = filedialog.askdirectory()
        if selected:
            self.folder_path.set(selected)

    def apply_patterns(self) -> None:
        if not self.folder_path.get():
            messagebox.showinfo("Select Folder", "Please select a folder before applying patterns.")
            return
        self.load_pdfs()

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
        pattern_map = self._gather_patterns()
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
            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                page_text = page.get_text("text")
                for column, patterns in pattern_map.items():
                    for pattern in patterns:
                        if pattern.search(page_text):
                            matches[column].append(Match(page_index=page_index, source="regex", pattern=pattern.pattern))
                            break

            entry = PDFEntry(path=pdf_path, doc=doc, matches=matches)
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
        for child in self.inner_frame.winfo_children():
            child.destroy()

    def _gather_patterns(self) -> Dict[str, List[re.Pattern[str]]]:
        pattern_map: Dict[str, List[re.Pattern[str]]] = {}
        for column in COLUMNS:
            text_widget = self.pattern_texts[column]
            raw_text = text_widget.get("1.0", tk.END)
            patterns = [line.strip() for line in raw_text.splitlines() if line.strip()]
            if not patterns:
                patterns = DEFAULT_PATTERNS[column]
                text_widget.delete("1.0", tk.END)
                text_widget.insert("1.0", "\n".join(patterns))
            compiled = []
            for pattern in patterns:
                try:
                    compiled.append(re.compile(pattern, re.IGNORECASE))
                except re.error as exc:
                    messagebox.showerror("Invalid Pattern", f"Invalid regex '{pattern}' for {column}: {exc}")
            pattern_map[column] = compiled
        return pattern_map

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
            ttk.Label(self.inner_frame, text=entry.path.relative_to(Path(self.folder_path.get())), anchor="w", width=30, wraplength=200).grid(row=row_index, column=0, sticky="nw", padx=4, pady=4)
            for col_index, column in enumerate(COLUMNS, start=1):
                cell = PDFCell(self.inner_frame, self, entry, column)
                cell.frame.grid(row=row_index, column=col_index, padx=4, pady=4, sticky="nsew")
                self.cells[(entry, column)] = cell
                self._update_cell(entry, column)

        for idx in range(len(COLUMNS) + 1):
            self.inner_frame.columnconfigure(idx, weight=1)

    def _update_cell(self, entry: PDFEntry, category: str) -> None:
        cell = self.cells.get((entry, category))
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
        photo = self._render_page(entry.doc, match.page_index)
        info_parts = [f"Match {current_index + 1}/{len(matches)}", f"Page {match.page_index + 1}"]
        if match.source == "manual":
            info_parts.append("(manual)")
        cell.update_display(photo, " | ".join(info_parts))

    def _render_page(self, doc: fitz.Document, page_index: int) -> Optional[ImageTk.PhotoImage]:
        try:
            page = doc.load_page(page_index)
            zoom_matrix = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom_matrix)
            mode = "RGBA" if pix.alpha else "RGB"
            image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            max_width = 300
            if image.width > max_width:
                ratio = max_width / image.width
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

    def _open_pdf(self, pdf_path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(pdf_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f"open '{pdf_path}'")
            else:
                os.system(f"xdg-open '{pdf_path}' >/dev/null 2>&1 &")
        except Exception as exc:
            messagebox.showwarning("Open PDF", f"Could not open PDF: {exc}")

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
