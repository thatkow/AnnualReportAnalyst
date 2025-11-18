from __future__ import annotations

import csv
import io
import os
import shutil
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional

from tkinter import messagebox

import requests

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover - handled at runtime
    OpenAI = None  # type: ignore[assignment]

from app_logging import get_logger
from constants import COLUMNS, DEFAULT_OPENAI_MODEL, SCRAPE_EXPECTED_COLUMNS
from models import ScrapeJob
from pdf_utils import normalize_header_row


OPENAI_API_BASE_URL = os.environ.get("OPENAI_API_BASE", "https://api.openai.com/v1")



OPENAI_HTTP_TIMEOUT = 120


class ScrapeManagerMixin:
    logger = get_logger()

    scrape_button: Any
    scrape_progress: Any
    pdf_entries: List[Any]
    company_var: Any
    companies_dir: Path
    prompts_dir: Path
    openai_model_vars: Dict[str, Any]
    scrape_upload_mode_vars: Dict[str, Any]
    scrape_panels: Dict[Any, Any]
    active_scrape_key: Optional[Any]
    root: Any

    def _create_openai_sdk_client(self, api_key: str) -> Optional[Any]:
        if OpenAI is None:
            return None
        try:
            return OpenAI(api_key=api_key)
        except TypeError as exc:
            if "proxies" in str(exc).lower():
                if not getattr(self, "_openai_proxy_warning_logged", False):
                    self.logger.warning(
                        "OpenAI SDK initialization failed due to incompatible proxy arguments; "
                        "falling back to direct HTTP requests."
                    )
                    setattr(self, "_openai_proxy_warning_logged", True)
                return None
            raise

    def _openai_http_request(
        self,
        api_key: str,
        endpoint: str,
        *,
        json_payload: Optional[Dict[str, Any]] = None,
        files: Optional[Dict[str, Any]] = None,
        data: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        base_url = OPENAI_API_BASE_URL.rstrip("/")
        url = f"{base_url}/{endpoint.lstrip('/')}"
        headers = {"Authorization": f"Bearer {api_key}"}
        if json_payload is not None:
            headers["Content-Type"] = "application/json"
        response = requests.post(
            url,
            headers=headers,
            json=json_payload,
            files=files,
            data=data,
            timeout=OPENAI_HTTP_TIMEOUT,
        )
        if response.status_code >= 400:
            try:
                error_detail: Any = response.json()
            except ValueError:
                error_detail = response.text
            raise RuntimeError(
                f"OpenAI request to '{endpoint}' failed with status {response.status_code}: {error_detail}"
            )
        try:
            return response.json()
        except ValueError as exc:  # pragma: no cover - network response issue
            raise RuntimeError(f"OpenAI response for '{endpoint}' was not valid JSON") from exc

    def _upload_pdfs_via_http(self, api_key: str, pdf_paths: List[Path]) -> List[str]:
        file_ids: List[str] = []
        for pdf_path in pdf_paths:
            self.logger.info("AIScrape uploading %s via HTTP fallback", pdf_path)
            with pdf_path.open("rb") as pdf_file:
                response_json = self._openai_http_request(
                    api_key,
                    "files",
                    files={"file": (pdf_path.name, pdf_file, "application/pdf")},
                    data={"purpose": "assistants"},
                )
            file_id = response_json.get("id")
            if not file_id:
                raise ValueError(f"Failed to upload {pdf_path.name} to OpenAI via HTTP fallback")
            file_ids.append(str(file_id))
            self.logger.info(
                "AIScrape uploaded %s via HTTP fallback as file id %s", pdf_path.name, file_id
            )
        return file_ids

    def _submit_openai_response_http(
        self, api_key: str, model: str, user_entries: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        payload = {
            "model": model,
            "input": [
                {"role": "system", "content": "You are a financial statement parser."},
                {"role": "user", "content": user_entries},
            ],
        }
        self.logger.info("AIScrape submitting HTTP response request (model=%s)", model)
        return self._openai_http_request(api_key, "responses", json_payload=payload)

    def _get_prompt_text(self, company: str, category: str) -> Optional[str]:
        candidate_paths: List[Path] = []
        if company:
            company_dir = self.companies_dir / company
            candidate_paths.extend(
                [
                    company_dir / "prompts" / f"{category}.txt",
                    company_dir / "prompt" / f"{category}.txt",
                ]
            )
        candidate_paths.append(self.prompts_dir / f"{category}.txt")
        for path in candidate_paths:
            if path.exists():
                try:
                    return path.read_text(encoding="utf-8")
                except Exception:
                    continue
        return None

    def _strip_code_fence(self, text: str) -> str:
        import re

        fence = re.search(r"```(?:[^`\n]*)\n([\s\S]*?)```", text)
        if fence:
            return fence.group(1)
        return text

    def _parse_multiplier_response(
        self, response: str
    ) -> (Optional[str], Optional[List[str]], List[List[str]]):
        cleaned = self._strip_code_fence(response)
        raw_lines = [line for line in cleaned.splitlines() if line.strip()]
        multiplier: Optional[str] = None
        data_lines: List[str] = []
        for line in raw_lines:
            stripped = line.strip()
            if multiplier is None and stripped.lower().startswith("multiplier"):
                import re

                match_obj = re.search(r"([-+]?\d[\d,]*\.?\d*)", stripped)
                if match_obj:
                    multiplier = match_obj.group(1)
                continue
            data_lines.append(stripped)

        rows: List[List[str]] = []
        if data_lines:
            reader = csv.reader(io.StringIO("\n".join(data_lines)))
            try:
                for parsed in reader:
                    rows.append([cell.strip() for cell in parsed])
            except csv.Error:
                rows.extend([line.split(",") for line in data_lines])
                rows = [[cell.strip() for cell in row] for row in rows]
        header: Optional[List[str]] = None
        data_rows: List[List[str]] = []
        if rows:
            candidate = normalize_header_row(rows[0])
            if candidate is not None:
                header = candidate
                data_rows = rows[1:]
            else:
                data_rows = rows
        column_count = len(header) if header is not None else len(SCRAPE_EXPECTED_COLUMNS)
        normalized_rows: List[List[str]] = []
        for row in data_rows:
            values = list(row[:column_count])
            if len(values) < column_count:
                values.extend([""] * (column_count - len(values)))
            normalized_rows.append(values)
        if header is not None:
            if len(header) > column_count:
                header = header[:column_count]
            elif len(header) < column_count:
                header = header + [""] * (column_count - len(header))
        return multiplier, header, normalized_rows

    def _csv_has_data(self, path: Path) -> bool:
        try:
            with path.open("r", encoding="utf-8", newline="") as fh:
                reader = csv.reader(fh)
                header_seen = False
                for raw_row in reader:
                    if not any(cell.strip() for cell in raw_row):
                        continue
                    normalized = [cell.strip() for cell in raw_row]
                    if not header_seen:
                        header_seen = True
                        if normalize_header_row(normalized) is None:
                            return True
                        continue
                    return True
        except OSError:
            return False
        return False

    def _call_openai_with_pdfs(
        self, api_key: str, prompt: str, pdf_paths: List[Path], model_name: str
    ) -> str:
        sanitized_key = api_key.strip()
        if not sanitized_key:
            raise ValueError("API key is required")
        if not pdf_paths:
            raise ValueError("No PDF pages available for OpenAI request")

        selected_model = model_name.strip() or DEFAULT_OPENAI_MODEL

        client = self._create_openai_sdk_client(sanitized_key)

        file_ids: List[str] = []
        if client is not None:
            for pdf_path in pdf_paths:
                self.logger.info("AIScrape uploading %s", pdf_path)
                with pdf_path.open("rb") as pdf_file:
                    uploaded = client.files.create(file=pdf_file, purpose="assistants")
                    file_id = getattr(uploaded, "id", None)
                    if not file_id:
                        raise ValueError(f"Failed to upload {pdf_path.name} to OpenAI")
                    file_ids.append(str(file_id))
                    self.logger.info("AIScrape uploaded %s as file id %s", pdf_path.name, file_id)
        else:
            file_ids = self._upload_pdfs_via_http(sanitized_key, pdf_paths)

        user_entries: List[Dict[str, Any]] = [
            {"type": "input_text", "text": prompt},
            {
                "type": "input_text",
                "text": "Parse the attached PDFs and return the multiplier value and CSV rows.",
            },
        ]
        user_entries.extend({"type": "input_file", "file_id": fid} for fid in file_ids)

        self.logger.info(
            "AIScrape submitting request (model=%s, files=%s)",
            selected_model,
            file_ids,
        )
        if client is not None:
            response = client.responses.create(
                model=selected_model,
                input=
                [
                    {
                        "role": "system",
                        "content": "You are a financial statement parser.",
                    },
                    {"role": "user", "content": user_entries},
                ],
            )
            self.logger.info("AIScrape response received (model=%s)", selected_model)
        else:
            response = self._submit_openai_response_http(
                sanitized_key, selected_model, user_entries
            )
            self.logger.info("AIScrape HTTP response received (model=%s)", selected_model)
        return self._extract_openai_response_text(response)

    def _call_openai_with_text(
        self, api_key: str, prompt: str, text_payload: str, model_name: str
    ) -> str:
        sanitized_key = api_key.strip()
        if not sanitized_key:
            raise ValueError("API key is required")
        cleaned_text = text_payload.strip()
        if not cleaned_text:
            raise ValueError("Extracted text is empty")

        selected_model = model_name.strip() or DEFAULT_OPENAI_MODEL

        client = self._create_openai_sdk_client(sanitized_key)

        user_entries: List[Dict[str, Any]] = [
            {"type": "input_text", "text": prompt},
            {
                "type": "input_text",
                "text": "Parse the provided text excerpt and return the multiplier value and CSV rows.",
            },
            {"type": "input_text", "text": cleaned_text},
        ]

        self.logger.info(
            "AIScrape submitting text request (model=%s, characters=%s)",
            selected_model,
            len(cleaned_text),
        )
        if client is not None:
            response = client.responses.create(
                model=selected_model,
                input=[
                    {
                        "role": "system",
                        "content": "You are a financial statement parser.",
                    },
                    {"role": "user", "content": user_entries},
                ],
            )
            self.logger.info("AIScrape text response received (model=%s)", selected_model)
        else:
            response = self._submit_openai_response_http(
                sanitized_key, selected_model, user_entries
            )
            self.logger.info("AIScrape HTTP text response received (model=%s)", selected_model)
        return self._extract_openai_response_text(response)

    def _call_openai_for_job(self, job: ScrapeJob, api_key: str) -> str:
        if job.upload_mode == "text":
            if not job.text_payload:
                raise ValueError("No extracted text available for OpenAI request")
            return self._call_openai_with_text(
                api_key,
                job.prompt_text,
                job.text_payload,
                job.model_name,
            )
        if job.temp_pdf is None:
            raise ValueError("No PDF prepared for OpenAI request")
        return self._call_openai_with_pdfs(
            api_key,
            job.prompt_text,
            [job.temp_pdf],
            job.model_name,
        )

    def _extract_openai_response_text(self, response: Any) -> str:
        def _get(obj: Any, key: str, default: Any = None) -> Any:
            if isinstance(obj, dict):
                return obj.get(key, default)
            return getattr(obj, key, default)

        text_output = _get(response, "output_text")
        if text_output:
            if isinstance(text_output, (list, tuple)):
                combined = "\n".join(str(part).strip() for part in text_output if part).strip()
            else:
                combined = str(text_output).strip()
            if combined:
                return combined

        output_items = _get(response, "output")
        if output_items:
            collected: List[str] = []
            for item in output_items:
                contents = _get(item, "content")
                if not contents:
                    continue
                for content in contents:
                    if _get(content, "type") == "output_text":
                        collected.append(str(_get(content, "text", "")))
            combined = "\n".join(part.strip() for part in collected if part).strip()
            if combined:
                return combined

        for choice in _get(response, "choices", []) or []:
            message = _get(choice, "message")
            content = _get(message, "content")
            if isinstance(content, str) and content.strip():
                return content.strip()

        raise ValueError("OpenAI response did not contain any text output")

    def scrape_selected_pages(self) -> None:
        if not self.pdf_entries:
            messagebox.showinfo("AIScrape", "Load PDFs before running AIScrape.")
            return

        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo("AIScrape", "Select a company before running AIScrape.")
            return

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("AIScrape", "Enter an OpenAI API key before running AIScrape.")
            self.api_key_entry.focus_set()
            return
        self._persist_api_key(api_key)

        prompts: Dict[str, str] = {}
        missing: List[str] = []
        for category in COLUMNS:
            prompt_text = self._get_prompt_text(company, category)
            if prompt_text is None:
                missing.append(category)
            else:
                prompts[category] = prompt_text

        if missing:
            messagebox.showerror(
                "AIScrape", f"Prompt files not found for: {', '.join(missing)}."
            )
            return

        self._save_pattern_config()

        scrape_root = self.companies_dir / company / "openapiscrape"
        scrape_root.mkdir(parents=True, exist_ok=True)

        jobs: List[ScrapeJob] = []
        prep_errors: List[str] = []

        for entry in self.pdf_entries:
            for category in COLUMNS:
                pages = self.get_selected_pages(entry, category)
                if not pages:
                    continue
                prompt_text = prompts.get(category)
                if not prompt_text:
                    continue
                model_var = self.openai_model_vars.get(category)
                model_name = model_var.get() if model_var is not None else DEFAULT_OPENAI_MODEL
                mode_var = self.scrape_upload_mode_vars.get(category)
                upload_mode = mode_var.get() if mode_var is not None else "pdf"
                target_dir = scrape_root / entry.path.stem
                csv_path = target_dir / f"{category}.csv"
                already_processed = False

                # Primary check: file exists on disk and has data
                if csv_path.exists() and self._csv_has_data(csv_path):
                    already_processed = True
                else:
                    # Fallback: check panel if present (NO csv_path attribute—only load_from_files)
                    panel = self.scrape_panels.get((entry.path, category))
                    if panel is not None:
                        # Ask panel to refresh its data and rely on disk only
                        try:
                            panel.load_from_files()
                        except Exception:
                            pass
                        if csv_path.exists() and self._csv_has_data(csv_path):
                            already_processed = True

                if already_processed:
                    self.logger.info(
                        "AIScrape skipping %s | %s (existing file on disk)",
                        entry.path.name,
                        category,
                    )
                    continue
                temp_pdf: Optional[Path] = None
                text_payload: Optional[str] = None
                if upload_mode == "text":
                    text_payload = self.extract_pages_text(entry.doc, pages)
                    if not text_payload:
                        prep_errors.append(
                            f"{entry.path.name} - {category}: Unable to extract text from selected pages"
                        )
                        continue
                else:
                    temp_pdf = self.export_pages_to_pdf(entry.doc, pages)
                    if temp_pdf is None:
                        prep_errors.append(
                            f"{entry.path.name} - {category}: Unable to prepare selected pages"
                        )
                        continue
                panel = self.scrape_panels.get((entry.path, category))
                jobs.append(
                    ScrapeJob(
                        entry=entry,
                        category=category,
                        pages=pages,
                        prompt_text=prompt_text,
                        model_name=model_name,
                        upload_mode=upload_mode,
                        target_dir=target_dir,
                        temp_pdf=temp_pdf,
                        text_payload=text_payload,
                    )
                )
                if panel is not None:
                    panel.mark_loading()

        if prep_errors and not jobs:
            messagebox.showerror("AIScrape", "\n".join(prep_errors))
            return
        if not jobs:
            messagebox.showinfo("AIScrape", "Select pages before running AIScrape.")
            return

        self.scrape_button.configure(state="disabled")
        self.scrape_progress.configure(value=0, maximum=len(jobs))

        thread = threading.Thread(
            target=self._run_scrape_jobs,
            args=(jobs, api_key, prep_errors),
            daemon=True,
        )
        self._scrape_thread = thread
        thread.start()

    def _run_scrape_jobs(
        self,
        jobs: List[ScrapeJob],
        api_key: str,
        prep_errors: List[str],
    ) -> None:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import threading, time

        errors: List[str] = list(prep_errors)
        total = len(jobs)
        start_all = time.time()

        def process_job(index: int, job: ScrapeJob) -> tuple[ScrapeJob, bool, Optional[str]]:
            thread_name = threading.current_thread().name
            start_time = time.time()
            self.logger.info(
                "[THREAD-START] %s → %s | category=%s | pages=%s | model=%s | start=%.2fs",
                thread_name,
                job.entry.path.name,
                job.category,
                job.pages,
                job.model_name,
                start_time - start_all,
            )

            multiplier: Optional[str] = None
            success = False
            try:
                job.target_dir.mkdir(parents=True, exist_ok=True)
                pdf_folder = job.target_dir / "PDF_FOLDER"
                pdf_folder.mkdir(parents=True, exist_ok=True)
                out_pdf = pdf_folder / f"{job.category}.pdf"
                if job.temp_pdf is not None and job.temp_pdf.exists():
                    shutil.copyfile(job.temp_pdf, out_pdf)
                else:
                    tmp_cut = self.export_pages_to_pdf(job.entry.doc, job.pages)
                    if tmp_cut is not None and tmp_cut.exists():
                        try:
                            shutil.copyfile(tmp_cut, out_pdf)
                        finally:
                            try:
                                tmp_cut.unlink()
                            except Exception:
                                pass

                response_text = self._call_openai_for_job(job, api_key)
                multiplier, header, rows = self._parse_multiplier_response(response_text)
                self.logger.info(
                    "[THREAD] %s finished OpenAI call for %s | %s | rows=%d",
                    thread_name,
                    job.entry.path.name,
                    job.category,
                    len(rows),
                )

                job.target_dir.mkdir(parents=True, exist_ok=True)
                raw_path = job.target_dir / f"{job.category}_raw.txt"
                raw_path.write_text(response_text, encoding="utf-8")
                if multiplier is not None:
                    multiplier_path = job.target_dir / f"{job.category}_multiplier.txt"
                    multiplier_path.write_text(str(multiplier).strip(), encoding="utf-8")

                csv_path = job.target_dir / f"{job.category}.csv"
                header_row = header or SCRAPE_EXPECTED_COLUMNS
                with csv_path.open("w", encoding="utf-8", newline="") as fh:
                    writer = csv.writer(fh, quoting=csv.QUOTE_MINIMAL)
                    writer.writerow(header_row)
                    if rows:
                        writer.writerows(rows)
                success = True
            except Exception as exc:
                self.logger.exception("[THREAD-ERROR] %s failed for %s | %s", thread_name, job.entry.path.name, job.category)
                errors.append(f"{job.entry.path.name} - {job.category}: {exc}")
            finally:
                try:
                    if job.temp_pdf is not None and job.temp_pdf.exists():
                        job.temp_pdf.unlink()
                except Exception:
                    pass

            end_time = time.time()
            elapsed = end_time - start_time
            self.logger.info(
                "[THREAD-END] %s completed → %s | category=%s | success=%s | elapsed=%.2fs",
                thread_name,
                job.entry.path.name,
                job.category,
                success,
                elapsed,
            )
            return job, success, multiplier
        # Use the thread_count from ReportAppV2 if available
        try:
            max_workers = getattr(self, "thread_count", None)
            if not isinstance(max_workers, int) or max_workers <= 0:
                self.logger.warning("⚠️ Invalid thread_count on self, defaulting to 3")
                max_workers = 3
            else:
                self.logger.info(
                    "Using configured thread_count from ReportAppV2: %d", max_workers
                )

            if total < max_workers:
                max_workers = total
        except Exception as e:
            self.logger.warning("⚠️ Could not get thread_count from ReportAppV2: %s", e)
            max_workers = 3

        self.logger.info(
            "Starting parallel AIScrape: %d jobs with %d workers (configured=%d)",
            total,
            max_workers,
            getattr(self, "thread_count", 3),
        )

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(process_job, idx, jb): jb for idx, jb in enumerate(jobs, start=1)}
            for future in as_completed(futures):
                job, success, multiplier = future.result()
                self.root.after(0, self._on_scrape_job_progress, job, 1, success, multiplier)

        total_time = time.time() - start_all
        self.logger.info("✅ All threads finished for %d AIScrape jobs | total elapsed = %.2fs", total, total_time)
        self.root.after(0, self._on_scrape_jobs_finished, total, errors)

    def _on_scrape_job_progress(
        self,
        job: ScrapeJob,
        completed: int,
        _success: bool,
        _multiplier: Optional[str],
    ) -> None:
        self.scrape_progress.configure(value=completed)
        panel = self.scrape_panels.get((job.entry.path, job.category))
        if panel is not None:
            panel.load_from_files()
            panel.set_active(self.active_scrape_key == (job.entry.path, job.category))
        if self.active_scrape_key == (job.entry.path, job.category):
            if self.scrape_preview_pages:
                self._display_scrape_preview_page(force=True)
            else:
                self._show_scrape_preview(job.entry, job.category)
        self.refresh_combined_tab()

    def _on_scrape_jobs_finished(self, total: int, errors: List[str]) -> None:
        self.scrape_button.configure(state="normal")
        self.scrape_progress.configure(value=0)
        self._scrape_thread = None
        if errors:
            messagebox.showerror("AIScrape", "\n".join(errors))
        else:
            messagebox.showinfo(
                "AIScrape",
                f"Saved {total} OpenAI response(s) to 'openapiscrape'.",
            )
