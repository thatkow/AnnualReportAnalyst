from __future__ import annotations

import json
import re
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import os
import tkinter as tk
from tkinter import messagebox, ttk

try:
    import fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - handled at runtime
    fitz = None  # type: ignore[assignment]

from PIL import Image, ImageTk

from app_logging import get_logger
from constants import COLUMNS
from pdf_utils import Match, PDFEntry


logger = get_logger()


class PDFManagerMixin:
    """Mixin providing PDF-related behaviours for the report application."""

    root: tk.Misc
    pattern_texts: Dict[str, tk.Text]
    case_insensitive_vars: Dict[str, tk.BooleanVar]
    whitespace_as_space_vars: Dict[str, tk.BooleanVar]
    year_pattern_text: Optional[tk.Text]
    year_case_insensitive_var: tk.BooleanVar
    year_whitespace_as_space_var: tk.BooleanVar
    openai_model_vars: Dict[str, tk.StringVar]
    scrape_upload_mode_vars: Dict[str, tk.StringVar]
    pdf_entries: List[PDFEntry]
    category_rows: Dict[Tuple[Path, str], Any]
    inner_frame: tk.Frame
    scrape_panels: Dict[Tuple[Path, str], Any]
    auto_scale_tables_var: tk.BooleanVar
    scrape_type_notebook: ttk.Notebook  # type: ignore[name-defined]
    app_root: Path
    companies_dir: Path
    assigned_pages: Dict[str, Dict[str, Any]]
    assigned_pages_path: Optional[Path]
    folder_path: tk.StringVar
    company_var: tk.StringVar
    thumbnail_width_var: tk.IntVar

    def _compile_patterns(self) -> Tuple[Dict[str, List[re.Pattern[str]]], List[re.Pattern[str]]]:
        pattern_map: Dict[str, List[re.Pattern[str]]] = {}
        for column, widget in self.pattern_texts.items():
            lines = self._read_text_lines(widget)
            compiled: List[re.Pattern[str]] = []
            flags = re.IGNORECASE if self.case_insensitive_vars[column].get() else 0
            whitespace = self.whitespace_as_space_vars[column].get()
            for line in lines:
                pattern_text = line.replace(" ", r"\s+") if whitespace else line
                try:
                    compiled.append(re.compile(pattern_text, flags))
                except re.error as exc:
                    messagebox.showerror(
                        "Invalid Pattern",
                        f"Could not compile pattern '{line}' for {column}: {exc}",
                    )
                    compiled.clear()
                    break
            pattern_map[column] = compiled

        year_patterns: List[re.Pattern[str]] = []
        if self.year_pattern_text is not None:
            lines = self._read_text_lines(self.year_pattern_text)
            flags = re.IGNORECASE if self.year_case_insensitive_var.get() else 0
            whitespace = self.year_whitespace_as_space_var.get()
            for line in lines:
                pattern_text = line.replace(" ", r"\s+") if whitespace else line
                try:
                    year_patterns.append(re.compile(pattern_text, flags))
                except re.error as exc:
                    messagebox.showerror("Invalid Year Pattern", f"Could not compile '{line}': {exc}")
                    year_patterns.clear()
                    break
        self._save_pattern_config()
        return pattern_map, year_patterns

    def clear_entries(self) -> None:
        for entry in self.pdf_entries:
            try:
                entry.doc.close()
            except Exception:
                pass
        self.pdf_entries.clear()
        self.category_rows.clear()
        for child in self.inner_frame.winfo_children():
            child.destroy()
        self.active_scrape_key = None
        self._refresh_scrape_results()
        self._clear_scrape_preview()
        self.refresh_combined_tab()

    def unlock_pdfs_and_reload(self) -> None:
        """Rewrite PDFs to remove encryption and reload them into the app."""
        folder = self.folder_path.get()
        if not folder:
            messagebox.showinfo("Select Folder", "Choose a company before unlocking PDFs.")
            return

        if fitz is None:
            messagebox.showerror(
                "PDF Library Missing",
                "PyMuPDF is required to unlock PDFs. Please install dependencies and try again.",
            )
            return

        folder_path = Path(folder)
        if not folder_path.exists():
            messagebox.showerror("Folder Not Found", f"The folder '{folder}' does not exist.")
            return

        pdf_paths = sorted(folder_path.rglob("*.pdf"))
        if not pdf_paths:
            messagebox.showinfo("No PDFs", "No PDF files were found in the selected folder.")
            return

        # Close any open PDF documents before attempting to rewrite them so file handles
        # do not block replacement on Windows.
        self.clear_entries()

        progress_win = tk.Toplevel(self.root)
        progress_win.title("Unlocking PDFs")
        progress_win.geometry("360x120")
        progress_win.transient(self.root)
        progress_win.grab_set()

        ttk.Label(progress_win, text="Unlocking and cleaning PDF files…", font=("Arial", 11)).pack(pady=(20, 8))
        status_label = ttk.Label(progress_win, text="Preparing…")
        status_label.pack(pady=(0, 12))

        failures: List[str] = []
        for index, pdf_path in enumerate(pdf_paths, start=1):
            status_label.config(text=f"{index}/{len(pdf_paths)}: {pdf_path.name}")
            progress_win.update_idletasks()
            temp_path: Optional[Path] = None
            try:
                with tempfile.NamedTemporaryFile(
                    suffix=".unlocked.pdf", dir=pdf_path.parent, delete=False
                ) as tmp:
                    temp_path = Path(tmp.name)

                with fitz.open(pdf_path) as doc:  # type: ignore[arg-type]
                    if doc.needs_pass and not doc.authenticate(""):
                        raise RuntimeError("PDF requires a password to open.")

                    doc.save(
                        temp_path,
                        garbage=4,
                        deflate=True,
                        encryption=fitz.PDF_ENCRYPT_NONE,
                    )

                replaced = False
                try:
                    os.replace(temp_path, pdf_path)
                    replaced = True
                except PermissionError:
                    try:
                        os.chmod(pdf_path, 0o666)
                        os.replace(temp_path, pdf_path)
                        replaced = True
                    except Exception:
                        pass

                if not replaced:
                    raise RuntimeError(
                        "Access denied while replacing the unlocked PDF. "
                        "Ensure the file is not open in another program and try again."
                    )
            except Exception as exc:
                logger.exception("Failed to unlock %s", pdf_path)
                failures.append(f"{pdf_path.name}: {exc}")
            finally:
                if temp_path and temp_path.exists():
                    try:
                        temp_path.unlink(missing_ok=True)
                    except Exception:
                        pass

        try:
            progress_win.destroy()
        except Exception:
            pass

        if failures:
            details = "\n".join(failures[:5])
            more = "" if len(failures) <= 5 else "\n…and more"
            messagebox.showwarning("Unlock Issues", f"Some PDFs could not be unlocked:\n{details}{more}")

        self.load_pdfs()

    def load_pdfs(self) -> None:
        folder = self.folder_path.get()
        if not folder:
            messagebox.showinfo("Select Folder", "Choose a company before loading PDFs.")
            return

        # --- Create progress window ---
        progress_win = tk.Toplevel(self.root)
        progress_win.title("Loading PDFs")
        progress_win.geometry("320x100")
        progress_win.transient(self.root)
        progress_win.grab_set()

        # Simplified loading indicator (no progress bar)
        ttk.Label(progress_win, text="Loading PDF files…", font=("Arial", 11)).pack(pady=(25, 10))
        status_label = ttk.Label(progress_win, text="Please wait…")
        status_label.pack(pady=(5, 10))

        # Force the window to appear before heavy work starts
        progress_win.update_idletasks()
        progress_win.update()

        def update_progress(current: int, total: int, name: str) -> None:
            percent = (current / total) * 100 if total else 0
            progress_var.set(percent)
            status_label.config(text=f"{current}/{total}: {Path(name).name}")
            self.root.update_idletasks()

        def close_progress() -> None:
            try:
                progress_win.destroy()
            except Exception:
                pass

        folder_path = Path(folder)
        if not folder_path.exists():
            messagebox.showerror("Folder Not Found", f"The folder '{folder}' does not exist.")
            return

        pattern_map, year_patterns = self._compile_patterns()
        if any(not patterns for patterns in pattern_map.values()):
            return

        self.clear_entries()

        pdf_paths = sorted(folder_path.rglob("*.pdf"))
        if not pdf_paths:
            messagebox.showinfo("No PDFs", "No PDF files were found in the selected folder.")
            close_progress()
            return

        total = len(pdf_paths)
        current = 0

        import concurrent.futures

        def process_pdf(pdf_path: Path) -> Optional[PDFEntry]:
            try:
                doc = fitz.open(pdf_path)  # type: ignore[arg-type]
            except Exception as exc:
                messagebox.showwarning("PDF Error", f"Could not open '{pdf_path}': {exc}")
                return None

            matches: Dict[str, List[Match]] = {column: [] for column in COLUMNS}
            year_value = ""
            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                page_text = page.get_text("text")
                for column, patterns in pattern_map.items():
                    for pattern in patterns:
                        match_obj = pattern.search(page_text)
                        if match_obj:
                            matches[column].append(
                                Match(page_index=page_index, source="regex",
                                      pattern=pattern.pattern,
                                      matched_text=match_obj.group(0).strip())
                            )
                            break
                if not year_value:
                    for pattern in year_patterns:
                        year_match = pattern.search(page_text)
                        if year_match:
                            year_value = year_match.group(1) if year_match.groups() else year_match.group(0)
                            break

            entry = PDFEntry(path=pdf_path, doc=doc, matches=matches, year=year_value)
            self._apply_existing_assignments(entry)
            return entry

        from concurrent.futures import ThreadPoolExecutor, as_completed
        self.pdf_entries.clear()
        with ThreadPoolExecutor(max_workers=min(8, os.cpu_count() or 4)) as executor:
            futures = {executor.submit(process_pdf, p): p for p in pdf_paths}
            for i, future in enumerate(as_completed(futures), start=1):
                entry = future.result()
                if entry is not None:
                    self.pdf_entries.append(entry)
                self.root.after(0, lambda n=i, t=len(pdf_paths): self.root.title(f"Loading PDFs {n}/{t}"))

        self.root.title("Loading complete")

        self._rebuild_review_grid()
        self._save_config()
        self.refresh_combined_tab()

        status_label.config(text="✅ All PDFs loaded")
        self.root.update_idletasks()

        # --- Close progress window after done ---
        close_progress()

    def _apply_existing_assignments(self, entry: PDFEntry) -> None:
        record = self.assigned_pages.get(entry.path.name)
        if not isinstance(record, dict):
            return

        stored_year = record.get("year")
        if isinstance(stored_year, str) and stored_year:
            entry.year = stored_year
        elif isinstance(stored_year, (int, float)):
            entry.year = str(int(stored_year))

        selections = record.get("selections")
        if not isinstance(selections, dict):
            return

        total_pages = len(entry.doc)
        for category, raw_page in selections.items():
            try:
                page_index = int(raw_page)
            except (TypeError, ValueError):
                continue
            if page_index < 0 or page_index >= total_pages:
                continue

            matches = entry.matches.setdefault(category, [])
            selected_index: Optional[int] = None
            for idx, match in enumerate(matches):
                if match.page_index == page_index:
                    selected_index = idx
                    break
            if selected_index is None:
                manual_match = Match(page_index=page_index, source="manual")
                matches.append(manual_match)
                matches.sort(key=lambda m: m.page_index)
                try:
                    selected_index = matches.index(manual_match)
                except ValueError:
                    selected_index = None
            if selected_index is not None:
                entry.current_index[category] = selected_index
                entry.selected_pages[category] = [matches[selected_index].page_index]

        multi_map = record.get("multi_selections")
        if isinstance(multi_map, dict):
            for category, values in multi_map.items():
                if not isinstance(values, list):
                    continue
                valid_pages: List[int] = []
                matches = entry.matches.setdefault(category, [])
                for value in values:
                    try:
                        page_index = int(value)
                    except (TypeError, ValueError):
                        continue
                    if page_index < 0 or page_index >= total_pages:
                        continue
                    if all(match.page_index != page_index for match in matches):
                        matches.append(Match(page_index=page_index, source="manual"))
                    valid_pages.append(page_index)
                if matches:
                    matches.sort(key=lambda m: m.page_index)
                if valid_pages:
                    unique_sorted = sorted(dict.fromkeys(valid_pages))
                    entry.selected_pages[category] = unique_sorted
                    if unique_sorted:
                        first_page = unique_sorted[0]
                        try:
                            first_index = next(
                                idx for idx, match in enumerate(matches) if match.page_index == first_page
                            )
                        except StopIteration:
                            first_index = None
                        if first_index is not None:
                            entry.current_index[category] = first_index

    def render_page(
        self,
        doc: fitz.Document,  # type: ignore[type-arg]
        page_index: int,
        *,
        target_width: Optional[int] = None,
        target_height: Optional[int] = None,
    ) -> Optional[ImageTk.PhotoImage]:
        try:
            page = doc.load_page(page_index)
            zoom_matrix = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom_matrix)
            mode = "RGBA" if pix.alpha else "RGB"
            image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            if image.mode == "RGBA":
                image = image.convert("RGB")
            scale: Optional[float] = None
            if target_width and target_width > 0:
                scale = target_width / image.width
            if target_height and target_height > 0:
                height_scale = target_height / image.height
                scale = min(scale, height_scale) if scale else height_scale
            if scale and abs(scale - 1.0) > 0.01:
                new_size = (
                    max(1, int(image.width * scale)),
                    max(1, int(image.height * scale)),
                )
                image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(image, master=self.root)
        except Exception:
            return None

    def export_pages_to_pdf(self, doc: fitz.Document, pages: List[int]) -> Optional[Path]:  # type: ignore[type-arg]
        if not pages:
            return None
        temp_path: Optional[Path] = None
        try:
            unique_pages = sorted(dict.fromkeys(int(page) for page in pages))
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                temp_path = Path(tmp.name)
            new_doc = fitz.open()
            try:
                for page_index in unique_pages:
                    new_doc.insert_pdf(doc, from_page=page_index, to_page=page_index)
                new_doc.save(temp_path)
            finally:
                new_doc.close()
            return temp_path
        except Exception:
            try:
                if temp_path is not None:
                    temp_path.unlink()
            except Exception:
                pass
            return None

    def extract_pages_text(self, doc: fitz.Document, pages: List[int]) -> Optional[str]:  # type: ignore[type-arg]
        if not pages:
            return None
        snippets: List[str] = []
        seen: List[int] = sorted(dict.fromkeys(int(page) for page in pages))
        for page_index in seen:
            try:
                page = doc.load_page(page_index)
                text = page.get_text("text")
            except Exception:
                logger.exception(
                    "Failed to extract text for page %s in %s", page_index + 1, getattr(doc, "name", "document")
                )
                continue
            cleaned = text.strip()
            if not cleaned:
                continue
            snippets.append(f"--- Page {page_index + 1} ---\n{cleaned}")
        combined = "\n\n".join(snippets).strip()
        return combined or None

    def _get_selected_page_index(self, entry: PDFEntry, category: str) -> Optional[int]:
        matches = entry.matches.get(category, [])
        index = entry.current_index.get(category)
        if index is None or index < 0 or index >= len(matches):
            return None
        return matches[index].page_index

    def get_selected_pages(self, entry: PDFEntry, category: str) -> List[int]:
        pages = entry.selected_pages.get(category, [])
        if pages:
            return sorted(dict.fromkeys(int(page) for page in pages))
        page_index = self._get_selected_page_index(entry, category)
        if page_index is None:
            return []
        return [int(page_index)]

    def get_multi_page_indexes(self, entry: PDFEntry, category: str) -> List[int]:
        return self.get_selected_pages(entry, category)

    def _write_assigned_pages(self) -> bool:
        if self.assigned_pages_path is None:
            company = self.company_var.get().strip()
            if not company:
                return False
            self.assigned_pages_path = self.companies_dir / company / "assigned.json"
        try:
            self.assigned_pages_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError:
            messagebox.showwarning(
                "Commit Assignments",
                "Could not create the folder for saving assignments.",
            )
            return False
        try:
            with self.assigned_pages_path.open("w", encoding="utf-8") as fh:
                json.dump(self.assigned_pages, fh, indent=2)
        except OSError as exc:
            messagebox.showwarning("Commit Assignments", f"Could not save assignments: {exc}")
            return False
        return True

    def commit_assignments(self) -> None:
        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo("Select Company", "Please choose a company before committing assignments.")
            return
        if self.pdf_entries:
            for entry in self.pdf_entries:
                record: Dict[str, Any] = self.assigned_pages.get(entry.path.name, {})
                if not isinstance(record, dict):
                    record = {}
                if entry.year:
                    record["year"] = entry.year
                else:
                    record.pop("year", None)

                selections = record.get("selections")
                if not isinstance(selections, dict):
                    selections = {}

                multi = record.get("multi_selections")
                if not isinstance(multi, dict):
                    multi = {}

                for category in COLUMNS:
                    page_index = self._get_selected_page_index(entry, category)
                    if page_index is None:
                        selections.pop(category, None)
                    else:
                        selections[category] = int(page_index)

                    multi_pages = self.get_multi_page_indexes(entry, category)
                    if multi_pages:
                        multi[category] = [int(idx) for idx in multi_pages]
                    else:
                        multi.pop(category, None)

                if selections:
                    record["selections"] = selections
                else:
                    record.pop("selections", None)

                if multi:
                    record["multi_selections"] = multi
                else:
                    record.pop("multi_selections", None)

                if record:
                    self.assigned_pages[entry.path.name] = record
                else:
                    self.assigned_pages.pop(entry.path.name, None)

        if self._write_assigned_pages():
            messagebox.showinfo("Commit Assignments", "Assignments saved.")
