"""Configuration management helpers for ReportAppV2."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Optional

import tkinter as tk
from tkinter import messagebox

from constants import COLUMNS, DEFAULT_OPENAI_MODEL


class ConfigManagerMixin:
    pattern_config_path: Path
    local_config_path: Path
    config_path: Path
    pattern_texts: Dict[str, tk.Text]
    case_insensitive_vars: Dict[str, tk.BooleanVar]
    whitespace_as_space_vars: Dict[str, tk.BooleanVar]
    year_pattern_text: Optional[tk.Text]
    year_case_insensitive_var: tk.BooleanVar
    year_whitespace_as_space_var: tk.BooleanVar
    recent_download_minutes: tk.StringVar
    openai_model_vars: Dict[str, tk.StringVar]
    scrape_upload_mode_vars: Dict[str, tk.StringVar]
    scrape_column_widths: Dict[str, int]
    note_color_scheme: Dict[str, str]
    local_config_data: Dict[str, Any]
    api_key_var: tk.StringVar
    downloads_dir: tk.StringVar
    company_var: tk.StringVar
    auto_load_last_company_var: tk.BooleanVar
    _api_key_save_after: Optional[str]
    _suspend_api_key_save: bool
    root: tk.Misc

    def _save_pattern_config(self) -> None:
        payload = self._collect_pattern_config_payload()
        try:
            with self.pattern_config_path.open("w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
        except OSError:
            messagebox.showwarning(
                "Save Patterns", "Unable to save pattern configuration to disk."
            )

    def _load_pattern_config(self) -> None:
        if not self.pattern_config_path.exists():
            return
        try:
            with self.pattern_config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            messagebox.showwarning(
                "Load Patterns",
                "Unable to read pattern configuration; using defaults.",
            )
            return

        if not isinstance(data, dict):
            messagebox.showwarning(
                "Load Patterns", "Pattern configuration format is invalid; using defaults."
            )
            return

        patterns = data.get("patterns", {})
        if isinstance(patterns, dict):
            for column, widget in self.pattern_texts.items():
                values = patterns.get(column)
                if isinstance(values, list):
                    widget.delete("1.0", tk.END)
                    widget.insert(
                        "1.0",
                        "\n".join(str(item) for item in values if isinstance(item, str)),
                    )

        case_flags = data.get("case_insensitive", {})
        if isinstance(case_flags, dict):
            for column, var in self.case_insensitive_vars.items():
                if column in case_flags:
                    var.set(bool(case_flags[column]))

        whitespace_flags = data.get("space_as_whitespace", {})
        if isinstance(whitespace_flags, dict):
            for column, var in self.whitespace_as_space_vars.items():
                if column in whitespace_flags:
                    var.set(bool(whitespace_flags[column]))

        if self.year_pattern_text is not None:
            year_patterns = data.get("year_patterns")
            if isinstance(year_patterns, list):
                self.year_pattern_text.delete("1.0", tk.END)
                self.year_pattern_text.insert(
                    "1.0",
                    "\n".join(
                        str(item)
                        for item in year_patterns
                        if isinstance(item, str)
                    ),
                )

        year_case = data.get("year_case_insensitive")
        if isinstance(year_case, bool):
            self.year_case_insensitive_var.set(year_case)

        year_whitespace = data.get("year_space_as_whitespace")
        if isinstance(year_whitespace, bool):
            self.year_whitespace_as_space_var.set(year_whitespace)

        downloads_minutes = data.get("downloads_minutes")
        if isinstance(downloads_minutes, int):
            self.recent_download_minutes.set(str(downloads_minutes))

        openai_models = data.get("openai_models")
        if isinstance(openai_models, dict):
            for column in COLUMNS:
                if column in openai_models:
                    value = str(openai_models[column]).strip()
                    var = self.openai_model_vars.get(column)
                    if var is not None:
                        var.set(value or DEFAULT_OPENAI_MODEL)

        upload_modes = data.get("upload_modes")
        if isinstance(upload_modes, dict):
            for column in COLUMNS:
                if column in upload_modes:
                    value = str(upload_modes[column]).strip() or "pdf"
                    mode_var = self.scrape_upload_mode_vars.get(column)
                    if mode_var is not None:
                        mode_var.set(value)

        widths = data.get("scrape_column_widths")
        if isinstance(widths, dict):
            for key in ("category", "subcategory", "item", "note", "dates"):
                value = widths.get(key)
                try:
                    if value is not None:
                        self.scrape_column_widths[key] = max(40, int(value))
                except (TypeError, ValueError):
                    continue

        note_colors = data.get("note_colors")
        if isinstance(note_colors, dict):
            scheme: Dict[str, str] = {}
            for k, v in note_colors.items():
                if isinstance(k, str) and isinstance(v, str) and v.strip():
                    scheme[k.strip()] = v.strip()
            if scheme:
                self.note_color_scheme = scheme

    def _load_local_config(self) -> None:
        self.local_config_data = {}
        if not self.local_config_path.exists():
            return
        try:
            with self.local_config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return

        if not isinstance(data, dict):
            return

        self.local_config_data = data

        api_key_value = data.get("api_key")
        if isinstance(api_key_value, str):
            trimmed = api_key_value.strip()
            if trimmed:
                self.api_key_var.set(trimmed)

        downloads_dir_value = data.get("downloads_dir")
        if isinstance(downloads_dir_value, str) and not self.downloads_dir.get().strip():
            trimmed_downloads = downloads_dir_value.strip()
            if trimmed_downloads:
                self.downloads_dir.set(trimmed_downloads)

    def _write_local_config(self) -> None:
        data = dict(self.local_config_data)
        api_key_value = self.api_key_var.get().strip()
        if api_key_value:
            data["api_key"] = api_key_value
        else:
            data.pop("api_key", None)

        if not data:
            if self.local_config_path.exists():
                try:
                    self.local_config_path.unlink()
                except OSError:
                    messagebox.showwarning(
                        "Local Config", "Unable to remove local configuration file."
                    )
                    return
        else:
            try:
                with self.local_config_path.open("w", encoding="utf-8") as fh:
                    json.dump(data, fh, indent=2)
            except OSError:
                messagebox.showwarning(
                    "Local Config", "Unable to save local configuration file."
                )
                return

        self.local_config_data = data

    def _persist_api_key(self, value: str) -> None:
        trimmed = value.strip()
        if trimmed:
            if self.local_config_data.get("api_key") == trimmed:
                return
            self.local_config_data["api_key"] = trimmed
        elif "api_key" in self.local_config_data:
            self.local_config_data.pop("api_key", None)
        else:
            return
        self._write_local_config()

    def _on_api_key_var_changed(self, *_: Any) -> None:
        if self._suspend_api_key_save:
            return
        if self._api_key_save_after is not None:
            try:
                self.root.after_cancel(self._api_key_save_after)
            except Exception:
                pass
        self._api_key_save_after = self.root.after(600, self._flush_api_key_save)

    def _flush_api_key_save(self) -> None:
        self._api_key_save_after = None
        self._persist_api_key(self.api_key_var.get())

    def _load_config(self) -> None:
        if not self.config_path.exists():
            return
        try:
            with self.config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return

        downloads = data.get("downloads_dir")
        if isinstance(downloads, str):
            self.downloads_dir.set(downloads)

        last_company = data.get("last_company")
        if isinstance(last_company, str):
            self.company_var.set(last_company)

        auto_load = data.get("auto_load_last_company")
        if isinstance(auto_load, bool):
            self.auto_load_last_company_var.set(auto_load)

    def _save_config(self) -> None:
        data = {
            "downloads_dir": self.downloads_dir.get().strip(),
            "last_company": self.company_var.get().strip(),
            "auto_load_last_company": bool(self.auto_load_last_company_var.get()),
        }
        try:
            with self.config_path.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
        except OSError:
            messagebox.showwarning(
                "Save Config", "Unable to save configuration to disk."
            )
