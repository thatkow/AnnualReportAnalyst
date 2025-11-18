from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from constants import (
    COLUMNS,
    CONTROL_MASK,
    DEFAULT_OPENAI_MODEL,
    DEFAULT_PATTERNS,
    YEAR_DEFAULT_PATTERNS,
)
from pdf_utils import Match, PDFEntry
from ui_widgets import CategoryRow, CollapsibleFrame


class ReviewUIMixin:
    root: tk.Misc
    review_tab: ttk.Frame
    review_canvas: tk.Canvas
    inner_frame: ttk.Frame
    category_rows: Dict[Tuple[Path, str], CategoryRow]
    thumbnail_width_var: tk.IntVar
    scrape_panels: Dict[Tuple[Path, str], Any]
    active_scrape_key: Optional[Tuple[Path, str]]
    scrape_preview_pages: List[int]
    scrape_preview_entry: Optional[PDFEntry]
    scrape_preview_category: Optional[str]
    scrape_preview_cycle_index: int
    scrape_preview_last_width: int
    scrape_preview_render_width: int
    scrape_preview_render_page: Optional[int]
    scrape_preview_photo: Any
    scrape_preview_canvas: tk.Canvas
    scrape_preview_label: tk.Label
    scrape_preview_title_var: tk.StringVar
    scrape_preview_page_var: tk.StringVar
    pdf_entries: List[PDFEntry]
    fullscreen_preview_window: Optional[tk.Toplevel]
    fullscreen_preview_image: Any
    fullscreen_preview_entry: Optional[Path]
    fullscreen_preview_page: Optional[int]
    companies_dir: Path
    company_var: tk.StringVar
    pattern_texts: Dict[str, tk.Text]
    openai_model_vars: Dict[str, tk.StringVar]
    case_insensitive_vars: Dict[str, tk.BooleanVar]
    whitespace_as_space_vars: Dict[str, tk.BooleanVar]
    year_pattern_text: Optional[tk.Text]
    year_case_insensitive_var: tk.BooleanVar
    year_whitespace_as_space_var: tk.BooleanVar
    thumbnail_scale: ttk.Scale
    commit_button: ttk.Button
    canvas_window: int

    def build_review_tab(self, notebook: ttk.Notebook) -> None:
        review_tab = ttk.Frame(notebook)
        notebook.add(review_tab, text="Review")
        self.review_tab = review_tab

        options_section = CollapsibleFrame(review_tab, "Patterns & Review Options", initially_open=False)
        options_section.pack(fill=tk.X, padx=8, pady=(4, 0))

        options_inner = ttk.Frame(options_section.content, padding=8)
        options_inner.pack(fill=tk.BOTH, expand=True)

        patterns_frame = ttk.LabelFrame(options_inner, text="Regex patterns (one per line)", padding=8)
        patterns_frame.pack(fill=tk.BOTH, expand=True)

        columns_frame = ttk.Frame(patterns_frame)
        columns_frame.pack(fill=tk.X)

        for idx, column in enumerate(COLUMNS):
            column_frame = ttk.Frame(columns_frame)
            column_frame.grid(row=0, column=idx, padx=4, sticky="nsew")
            columns_frame.columnconfigure(idx, weight=1)

            ttk.Label(column_frame, text=column).pack(anchor="w")
            text_widget = tk.Text(column_frame, height=4, width=30)
            text_widget.pack(fill=tk.BOTH, expand=True)
            defaults = DEFAULT_PATTERNS.get(column, [])
            text_widget.insert("1.0", "\n".join(defaults))
            self.pattern_texts[column] = text_widget

            model_var = tk.StringVar(master=self.root, value=DEFAULT_OPENAI_MODEL)
            self.openai_model_vars[column] = model_var

            case_var = tk.BooleanVar(master=self.root, value=True)
            self.case_insensitive_vars[column] = case_var
            ttk.Checkbutton(column_frame, text="Case-insensitive", variable=case_var).pack(anchor="w", pady=(4, 0))

            whitespace_var = tk.BooleanVar(master=self.root, value=True)
            self.whitespace_as_space_vars[column] = whitespace_var
            ttk.Checkbutton(
                column_frame,
                text="Treat spaces as any whitespace",
                variable=whitespace_var,
            ).pack(anchor="w")

        apply_button = ttk.Button(patterns_frame, text="Apply Patterns", command=self.load_pdfs)
        apply_button.pack(anchor="e", pady=(8, 0))

        year_frame = ttk.LabelFrame(options_inner, text="Year pattern", padding=8)
        year_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
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

        review_controls = ttk.Frame(options_inner, padding=(0, 8, 0, 0))
        review_controls.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(review_controls, text="Thumbnail width:").pack(side=tk.LEFT)
        self.thumbnail_scale = ttk.Scale(
            review_controls,
            from_=160,
            to=420,
            orient=tk.HORIZONTAL,
            command=self._on_thumbnail_scale,
        )
        self.thumbnail_scale.set(self.thumbnail_width_var.get())
        self.thumbnail_scale.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ttk.Label(review_controls, textvariable=self.thumbnail_width_var).pack(side=tk.LEFT)

        review_container = ttk.Frame(review_tab)
        review_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.review_canvas = tk.Canvas(review_container)
        self.review_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        review_scrollbar = ttk.Scrollbar(review_container, orient=tk.VERTICAL, command=self.review_canvas.yview)
        review_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.review_canvas.configure(yscrollcommand=review_scrollbar.set)
        self.review_canvas.bind("<Enter>", self._bind_review_mousewheel)
        self.review_canvas.bind("<Leave>", self._unbind_review_mousewheel)

        self.inner_frame = ttk.Frame(self.review_canvas)
        self.canvas_window = self.review_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind(
            "<Configure>", lambda _e: self.review_canvas.configure(scrollregion=self.review_canvas.bbox("all"))
        )
        self.review_canvas.bind(
            "<Configure>",
            lambda event: self.review_canvas.itemconfigure(self.canvas_window, width=event.width),
        )

        actions_frame = ttk.Frame(review_tab, padding=8)
        actions_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.commit_button = ttk.Button(actions_frame, text="Commit", command=self.commit_assignments)
        self.commit_button.pack(side=tk.RIGHT)

    def _rebuild_review_grid(self) -> None:
        for child in self.inner_frame.winfo_children():
            child.destroy()
        self.category_rows.clear()

        if not self.pdf_entries:
            ttk.Label(self.inner_frame, text="Load PDFs to begin reviewing.").grid(
                row=0, column=0, padx=16, pady=16, sticky="nw"
            )
            return

        for row_index, entry in enumerate(self.pdf_entries):
            container = ttk.Frame(self.inner_frame, padding=8)
            container.grid(row=row_index, column=0, sticky="ew", padx=4, pady=4)
            container.columnconfigure(1, weight=1)

            info_frame = ttk.Frame(container)
            info_frame.grid(row=0, column=0, sticky="nw", padx=(0, 12))
            ttk.Label(info_frame, text=str(entry.path.name), anchor="w", width=30, wraplength=200).pack(anchor="w")
            if entry.year:
                ttk.Label(info_frame, text=f"Year: {entry.year}", foreground="#555555").pack(anchor="w", pady=(4, 0))

            types_frame = ttk.Frame(container)
            types_frame.grid(row=0, column=1, sticky="ew")
            types_frame.columnconfigure(0, weight=1)

            for idx, column in enumerate(COLUMNS):
                row = CategoryRow(types_frame, self, entry, column)
                row.frame.grid(row=idx, column=0, sticky="ew")
                if idx:
                    row.frame.grid_configure(pady=(8, 0))
                self.category_rows[(entry.path, column)] = row
                row.refresh()

        self.inner_frame.columnconfigure(0, weight=1)
        self._refresh_scrape_results()

    def _bind_review_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.review_canvas.bind_all("<MouseWheel>", self._on_review_mousewheel)
        self.review_canvas.bind_all("<Button-4>", self._on_review_mousewheel)
        self.review_canvas.bind_all("<Button-5>", self._on_review_mousewheel)

    def _unbind_review_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.review_canvas.unbind_all("<MouseWheel>")
        self.review_canvas.unbind_all("<Button-4>")
        self.review_canvas.unbind_all("<Button-5>")

    def _on_review_mousewheel(self, event: tk.Event) -> None:  # type: ignore[override]
        if getattr(event, "delta", 0):
            step = -1 if event.delta > 0 else 1
        else:
            step = -1 if event.num == 4 else 1
        self.review_canvas.yview_scroll(step, "units")

    def _on_thumbnail_scale(self, value: str) -> None:
        try:
            width = int(float(value))
        except (TypeError, ValueError):
            return
        self.thumbnail_width_var.set(width)
        for row in self.category_rows.values():
            row.set_thumbnail_width(width)

    def select_match(
        self,
        entry: PDFEntry,
        category: str,
        index: int,
        *,
        extend_selection: bool = False,
    ) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            return
        index = max(0, min(index, len(matches) - 1))
        entry.current_index[category] = index
        match = matches[index]
        page_index = match.page_index
        if extend_selection:
            pages = entry.selected_pages.setdefault(category, [])
            if page_index not in pages:
                pages.append(page_index)
                pages.sort()
        else:
            entry.selected_pages[category] = [page_index]
        row = self.category_rows.get((entry.path, category))
        if row is not None:
            row.update_selection()
        self._refresh_scrape_results()

    def manual_select(self, entry: PDFEntry, category: str) -> None:
        self.open_pdf(entry.path)
        max_pages = len(entry.doc)
        value = simpledialog.askinteger(
            "Manual Selection",
            f"Enter the page number (1-{max_pages}) for {category}:",
            parent=self.root,
            minvalue=1,
            maxvalue=max_pages,
        )
        if value is None:
            return
        page_index = value - 1
        match = Match(page_index=page_index, source="manual")
        entry.matches.setdefault(category, []).append(match)
        entry.current_index[category] = len(entry.matches[category]) - 1
        entry.matches[category].sort(key=lambda m: m.page_index)
        entry.current_index[category] = next(
            (idx for idx, m in enumerate(entry.matches[category]) if m.page_index == page_index),
            entry.current_index[category],
        )
        entry.selected_pages[category] = [page_index]
        row = self.category_rows.get((entry.path, category))
        if row is not None:
            row.refresh()
        self._refresh_scrape_results()

    def open_pdf(self, path: Path) -> None:
        self._open_with_default_app(path, "Open PDF")

    def open_file_path(self, path: Path) -> None:
        self._open_with_default_app(path, "Open File")

    def open_scrape_folder(self) -> None:
        company = self.company_var.get().strip()
        target: Optional[Path] = None
        if company:
            target = self.companies_dir / company / "openapiscrape"
        elif self.pdf_entries:
            target = self.pdf_entries[0].path.parent / "openapiscrape"
        else:
            messagebox.showinfo(
                "Open Scrape Folder",
                "Load PDFs or select a company before opening the scrape folder.",
            )
            return

        try:
            target.mkdir(parents=True, exist_ok=True)
        except OSError:
            messagebox.showwarning(
                "Open Scrape Folder", "Unable to prepare the OpenAI scrape folder for viewing."
            )
            return

        self._open_with_default_app(target, "Open Folder")

    def _open_with_default_app(self, path: Path, failure_title: str) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                if not os.environ.get("DISPLAY"):
                    raise RuntimeError("No graphical display available")
                subprocess.Popen(
                    ["xdg-open", str(path)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
        except Exception as exc:
            message = f"Unable to open '{path.name}' with the default application."
            if isinstance(exc, RuntimeError):
                message = f"{message}\n{exc}"
            messagebox.showwarning(failure_title, message)

    def toggle_fullscreen_preview(self, entry: PDFEntry, page_index: int) -> None:
        if getattr(self, "fullscreen_preview_window", None) is not None and self.fullscreen_preview_window.winfo_exists():
            if (
                self.fullscreen_preview_entry == entry.path
                and self.fullscreen_preview_page == page_index
            ):
                self._close_fullscreen_preview()
                return
            self._close_fullscreen_preview()
        self._open_fullscreen_preview(entry, page_index)

    def _open_fullscreen_preview(self, entry: PDFEntry, page_index: int) -> None:
        window = tk.Toplevel(self.root)
        window.title(f"{entry.path.name} - Page {page_index + 1}")
        try:
            window.attributes("-fullscreen", True)
        except tk.TclError:
            try:
                window.state("zoomed")
            except tk.TclError:
                screen_width = window.winfo_screenwidth()
                screen_height = window.winfo_screenheight()
                try:
                    window.geometry(f"{screen_width}x{screen_height}")
                except tk.TclError:
                    pass

        window.bind("<Escape>", lambda _e: self._close_fullscreen_preview())

        window.rowconfigure(0, weight=1)
        window.columnconfigure(0, weight=1)

        container = ttk.Frame(window)
        container.grid(row=0, column=0, sticky="nsew")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=0)

        canvas = tk.Canvas(container, background="#111111", highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=canvas.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(1, weight=0)
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        inner = ttk.Frame(canvas)
        window_item = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _update_scroll_region(_: tk.Event) -> None:
            bbox = canvas.bbox(window_item)
            if bbox:
                canvas.configure(scrollregion=bbox)

        inner.bind("<Configure>", _update_scroll_region)

        def _on_mousewheel(event: tk.Event) -> None:
            if getattr(event, "delta", 0):
                delta = -event.delta
                step = int(delta / 120) or (1 if delta > 0 else -1)
                canvas.yview_scroll(step * 4, "units")
            else:
                step = -1 if getattr(event, "num", 0) == 4 else 1
                canvas.yview_scroll(step * 4, "units")

        def _on_shift_mousewheel(event: tk.Event) -> None:
            if getattr(event, "delta", 0):
                delta = -event.delta
                step = int(delta / 120) or (1 if delta > 0 else -1)
                canvas.xview_scroll(step * 4, "units")
            else:
                step = -1 if getattr(event, "num", 0) == 4 else 1
                canvas.xview_scroll(step * 4, "units")

        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", _on_mousewheel)
        canvas.bind("<Button-5>", _on_mousewheel)
        canvas.bind("<Shift-MouseWheel>", _on_shift_mousewheel)
        canvas.bind("<Shift-Button-4>", _on_shift_mousewheel)
        canvas.bind("<Shift-Button-5>", _on_shift_mousewheel)

        screen_height = window.winfo_screenheight()
        target_height = max(screen_height - 160, 400)
        photo = self.render_page(entry.doc, page_index, target_height=target_height)
        if photo is None:
            label = ttk.Label(inner, text="Preview unavailable", padding=24)
            label.pack(expand=True, fill=tk.BOTH)
            self.fullscreen_preview_image = None
        else:
            label = tk.Label(inner, image=photo, background="#111111")
            label.pack()
            label.bind("<Button-1>", lambda _e: self._close_fullscreen_preview())
            self.fullscreen_preview_image = photo

        self.fullscreen_preview_window = window
        self.fullscreen_preview_entry = entry.path
        self.fullscreen_preview_page = page_index

    def _close_fullscreen_preview(self) -> None:
        if getattr(self, "fullscreen_preview_window", None) is not None and self.fullscreen_preview_window.winfo_exists():
            self.fullscreen_preview_window.destroy()
        self.fullscreen_preview_window = None
        self.fullscreen_preview_image = None
        self.fullscreen_preview_entry = None
        self.fullscreen_preview_page = None

    def _on_scrape_preview_mousewheel(self, event: tk.Event) -> None:  # type: ignore[override]
        if getattr(event, "delta", 0):
            delta = -event.delta
            step = int(delta / 120) or (1 if delta > 0 else -1)
            self.scrape_preview_canvas.yview_scroll(step, "units")
        else:
            step = -1 if getattr(event, "num", 0) == 4 else 1
            self.scrape_preview_canvas.yview_scroll(step, "units")

    def _on_scrape_preview_resize(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            width = int(getattr(event, "width", 0))
        except Exception:
            return
        if width <= 1:
            return
        if abs(width - self.scrape_preview_last_width) <= 2:
            return
        self.scrape_preview_last_width = width
        try:
            self.scrape_preview_canvas.itemconfigure(self.scrape_preview_window, width=width)
        except Exception:
            pass
        if self.scrape_preview_pages:
            self._display_scrape_preview_page(force=True)

    def _on_scrape_preview_click(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            state = int(event.state)
        except Exception:
            state = 0
        if state & CONTROL_MASK:
            self._cycle_scrape_preview()


    def _bind_number_keys_to_scrape_preview(self) -> None:
        """Bind numeric keys 1–9 to directly select a page index in the Scrape preview."""
        for i in range(1, 10):
            self.root.bind(str(i), self._on_scrape_number_key)


    def _on_scrape_number_key(self, event: tk.Event) -> None:
        """Jump directly to the selected Scrape preview cycle index based on numeric key."""
        try:
            idx = int(event.char) - 1
        except (ValueError, TypeError):
            return

        pages = self.scrape_preview_pages
        if not pages:
            return

        if idx < 0 or idx >= len(pages):
            return

        # Set index directly and refresh preview
        self.scrape_preview_cycle_index = idx
        try:
            self._display_scrape_preview_page(force=True)
        except Exception as e:
            print(f"Error jumping to preview page {idx + 1}: {e}")

    # Ensure bindings are initialized when Scrape tab is built
    def build_scrape_tab(self, notebook):
        super().build_scrape_tab(notebook)
        self._bind_number_keys_to_scrape_preview()

    def _cycle_scrape_preview(self) -> None:
        if len(self.scrape_preview_pages) < 2:
            return
        self.scrape_preview_cycle_index = (self.scrape_preview_cycle_index + 1) % len(self.scrape_preview_pages)
        self._display_scrape_preview_page(force=True)

    def _on_scrape_preview_label_configure(self, _: tk.Event) -> None:  # type: ignore[override]
        self._reset_scrape_preview_scroll()

    def _reset_scrape_preview_scroll(self) -> None:
        try:
            bbox = self.scrape_preview_canvas.bbox("all")
        except Exception:
            bbox = None
        if bbox:
            self.scrape_preview_canvas.configure(scrollregion=bbox)
        try:
            self.scrape_preview_canvas.yview_moveto(0)
        except Exception:
            pass

    def _show_scrape_preview(self, entry: PDFEntry, category: str) -> None:
        self.scrape_preview_entry = entry
        self.scrape_preview_category = category
        self.scrape_preview_pages = self.get_selected_pages(entry, category)
        self.scrape_preview_cycle_index = 0
        title = f"{entry.path.name} – {category}"
        self.scrape_preview_title_var.set(title)
        if not self.scrape_preview_pages:
            self.scrape_preview_label.configure(
                image="",
                text="No pages selected for this category.",
                background="#f0f0f0",
            )
            self.scrape_preview_canvas.configure(background="#f0f0f0")
            self.scrape_preview_page_var.set("")
            self.scrape_preview_photo = None
            self._reset_scrape_preview_scroll()
            return
        self.scrape_preview_render_page = None
        self.scrape_preview_render_width = 0
        self._display_scrape_preview_page(force=True)

    def _display_scrape_preview_page(self, force: bool = False) -> None:
        if not self.scrape_preview_entry or not self.scrape_preview_pages:
            self._clear_scrape_preview()
            return
        page_count = len(self.scrape_preview_pages)
        self.scrape_preview_cycle_index %= max(page_count, 1)
        page_index = self.scrape_preview_pages[self.scrape_preview_cycle_index]
        available_width = self.scrape_preview_last_width
        if available_width <= 1:
            available_width = self.scrape_preview_canvas.winfo_width()
        if available_width <= 1:
            available_width = self.scrape_preview_canvas.winfo_reqwidth()
        if available_width <= 1:
            available_width = max(self.thumbnail_width_var.get(), 360)
        display_width = max(int(available_width) - 16, 200)
        if (
            not force
            and self.scrape_preview_render_page == page_index
            and self.scrape_preview_render_width == display_width
        ):
            photo = self.scrape_preview_photo
        else:
            photo = self.render_page(
                self.scrape_preview_entry.doc,
                page_index,
                target_width=display_width,
            )
        if photo is None:
            self.scrape_preview_label.configure(image="", text="Preview unavailable", background="#f0f0f0")
            self.scrape_preview_canvas.configure(background="#f0f0f0")
            self.scrape_preview_photo = None
            self.scrape_preview_render_page = None
        else:
            self.scrape_preview_photo = photo
            self.scrape_preview_label.configure(image=photo, text="", background="#000000")
            self.scrape_preview_canvas.configure(background="#000000")
            self.scrape_preview_render_page = page_index
            self.scrape_preview_render_width = display_width
        self._reset_scrape_preview_scroll()
        self.scrape_preview_page_var.set(
            f"Page {page_index + 1} ({self.scrape_preview_cycle_index + 1}/{page_count})"
        )
