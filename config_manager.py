"""Unified configuration management for AnnualReportAnalyst."""

from __future__ import annotations

import json
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping

from platformdirs import user_config_path

from constants import (
    COLUMNS,
    DEFAULT_NOTE_COLOR_SCHEME,
    DEFAULT_OPENAI_MODEL,
    DEFAULT_PATTERNS,
    YEAR_DEFAULT_PATTERNS,
)


CONFIG_FILENAME = "AnnualReportAnalyst.config"


def _default_patterns() -> Dict[str, List[str]]:
    return {column: list(DEFAULT_PATTERNS.get(column, [])) for column in COLUMNS}


def _default_case_flags(value: bool) -> Dict[str, bool]:
    return {column: value for column in COLUMNS}


def _default_openai_models() -> Dict[str, str]:
    return {column: DEFAULT_OPENAI_MODEL for column in COLUMNS}


def _default_upload_modes() -> Dict[str, str]:
    return {column: "pdf" for column in COLUMNS}


def _default_scrape_column_widths() -> Dict[str, int]:
    return {
        "category": 140,
        "subcategory": 140,
        "item": 140,
        "note": 140,
        "dates": 140,
    }


@dataclass
class ConfigManager:
    """Dataclass-backed configuration container."""

    api_key: str = ""
    downloads_dir: str = ""
    thread_count: int = 3
    auto_load_last_company: bool = False
    last_company: str = ""
    downloads_minutes: int = 5
    note_colors: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_NOTE_COLOR_SCHEME))
    scrape_column_widths: Dict[str, int] = field(default_factory=_default_scrape_column_widths)
    scrape_row_height: int = 22
    patterns: Dict[str, List[str]] = field(default_factory=_default_patterns)
    case_insensitive: Dict[str, bool] = field(default_factory=lambda: _default_case_flags(True))
    space_as_whitespace: Dict[str, bool] = field(default_factory=lambda: _default_case_flags(True))
    year_patterns: List[str] = field(default_factory=lambda: list(YEAR_DEFAULT_PATTERNS))
    year_case_insensitive: bool = True
    year_space_as_whitespace: bool = True
    openai_models: Dict[str, str] = field(default_factory=_default_openai_models)
    upload_modes: Dict[str, str] = field(default_factory=_default_upload_modes)

    @classmethod
    def config_path(cls) -> Path:
        """Return the path to the unified configuration file."""

        base = Path(user_config_path("AnnualReportAnalyst"))
        base.mkdir(parents=True, exist_ok=True)
        return base / CONFIG_FILENAME

    @classmethod
    def load(cls) -> "ConfigManager":
        """Load configuration from disk, falling back to defaults."""

        path = cls.config_path()
        data: Dict[str, Any] = {}
        if path.exists():
            try:
                with path.open("r", encoding="utf-8") as fh:
                    loaded = json.load(fh)
                if isinstance(loaded, Mapping):
                    data = dict(loaded)
            except (OSError, json.JSONDecodeError):
                data = {}
        else:
            data = cls._load_legacy_configs()

        instance = cls()
        if data:
            instance.update_from_dict(data)

        if not path.exists():
            # Persist defaults (and legacy migration results) for future runs.
            instance.save()

        return instance

    @classmethod
    def _load_legacy_configs(cls) -> Dict[str, Any]:
        """Attempt to merge legacy configuration files into the new format."""

        legacy_root = Path(__file__).resolve().parent
        merged: Dict[str, Any] = {}

        # Legacy general config
        general_path = legacy_root / "data2_config.json"
        merged.update(cls._safe_read_json(general_path))

        # Legacy patterns config
        patterns_path = legacy_root / "pattern_config.json"
        patterns_data = cls._safe_read_json(patterns_path)
        if patterns_data:
            merged.setdefault("patterns", patterns_data.get("patterns", {}))
            merged.setdefault("case_insensitive", patterns_data.get("case_insensitive", {}))
            merged.setdefault("space_as_whitespace", patterns_data.get("space_as_whitespace", {}))
            merged.setdefault("year_patterns", patterns_data.get("year_patterns", []))
            merged.setdefault("year_case_insensitive", patterns_data.get("year_case_insensitive"))
            merged.setdefault("year_space_as_whitespace", patterns_data.get("year_space_as_whitespace"))
            merged.setdefault("downloads_minutes", patterns_data.get("downloads_minutes"))
            merged.setdefault("openai_models", patterns_data.get("openai_models", {}))
            merged.setdefault("upload_modes", patterns_data.get("upload_modes", {}))
            merged.setdefault("scrape_column_widths", patterns_data.get("scrape_column_widths", {}))
            merged.setdefault("note_colors", patterns_data.get("note_colors", {}))
            merged.setdefault("scrape_row_height", patterns_data.get("scrape_row_height"))

        # Legacy local config (API key, downloads dir overrides)
        local_path = legacy_root / "local_config.json"
        local_data = cls._safe_read_json(local_path)
        if isinstance(local_data.get("api_key"), str):
            merged["api_key"] = local_data["api_key"]
        if isinstance(local_data.get("downloads_dir"), str):
            merged.setdefault("downloads_dir", local_data["downloads_dir"])

        return {k: v for k, v in merged.items() if v not in (None, {})}

    @staticmethod
    def _safe_read_json(path: Path) -> Dict[str, Any]:
        if not path.exists():
            return {}
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return {}
        return data if isinstance(data, dict) else {}

    def update_from_dict(self, data: Mapping[str, Any]) -> None:
        """Merge external data into this instance, preserving defaults."""

        for field_info in fields(self):
            name = field_info.name
            if name not in data:
                continue
            value = data[name]
            if name in {"api_key", "downloads_dir", "last_company"}:
                if isinstance(value, str):
                    setattr(self, name, value.strip())
            elif name in {"thread_count", "downloads_minutes", "scrape_row_height"}:
                coerced = self._coerce_int(value, getattr(self, name))
                setattr(self, name, coerced)
            elif name == "auto_load_last_company":
                setattr(self, name, bool(value))
            elif name in {"note_colors", "scrape_column_widths", "openai_models", "upload_modes"}:
                self._merge_dict_field(name, value)
            elif name == "patterns":
                self._merge_patterns(value)
            elif name in {"case_insensitive", "space_as_whitespace"}:
                self._merge_bool_dict(name, value)
            elif name == "year_patterns":
                self.year_patterns = self._coerce_str_list(value, fallback=self.year_patterns)
            elif name in {"year_case_insensitive", "year_space_as_whitespace"}:
                setattr(self, name, bool(value))

        # Ensure scrape widths stay within reasonable bounds.
        widths: Dict[str, int] = {}
        for key, default_value in _default_scrape_column_widths().items():
            raw = self.scrape_column_widths.get(key, default_value)
            widths[key] = self._coerce_int(raw, default_value)
        self.scrape_column_widths = widths

        # Guarantee note colors retain defaults when absent.
        merged_colors = dict(DEFAULT_NOTE_COLOR_SCHEME)
        merged_colors.update({
            k: str(v).strip()
            for k, v in self.note_colors.items()
            if isinstance(k, str) and isinstance(v, str) and v.strip()
        })
        self.note_colors = merged_colors

        # Normalize dictionaries keyed by columns.
        patterns: Dict[str, List[str]] = {}
        for column in COLUMNS:
            values = self.patterns.get(column, [])
            if isinstance(values, Iterable) and not isinstance(values, (str, bytes)):
                cleaned = [str(item).strip() for item in values if str(item).strip()]
            else:
                cleaned = []
            patterns[column] = cleaned
        self.patterns = patterns

        case_flags: Dict[str, bool] = {}
        for column in COLUMNS:
            case_flags[column] = bool(self.case_insensitive.get(column, True))
        self.case_insensitive = case_flags

        whitespace_flags: Dict[str, bool] = {}
        for column in COLUMNS:
            whitespace_flags[column] = bool(self.space_as_whitespace.get(column, True))
        self.space_as_whitespace = whitespace_flags

        openai_models: Dict[str, str] = {}
        for column in COLUMNS:
            value = str(self.openai_models.get(column, DEFAULT_OPENAI_MODEL)).strip() or DEFAULT_OPENAI_MODEL
            openai_models[column] = value
        self.openai_models = openai_models

        upload_modes: Dict[str, str] = {}
        for column in COLUMNS:
            value = str(self.upload_modes.get(column, "pdf")).strip() or "pdf"
            upload_modes[column] = value
        self.upload_modes = upload_modes

    def _merge_dict_field(self, name: str, value: Any) -> None:
        if not isinstance(value, Mapping):
            return
        current = dict(getattr(self, name))
        for key, entry in value.items():
            if not isinstance(key, str):
                continue
            current[key] = entry
        setattr(self, name, current)

    def _merge_patterns(self, value: Any) -> None:
        if not isinstance(value, Mapping):
            return
        current = dict(self.patterns)
        for key, entry in value.items():
            if not isinstance(key, str):
                continue
            if isinstance(entry, Iterable) and not isinstance(entry, (str, bytes)):
                cleaned = [str(item) for item in entry if isinstance(item, (str, int, float))]
                current[key] = cleaned
        self.patterns = current

    def _merge_bool_dict(self, name: str, value: Any) -> None:
        if not isinstance(value, Mapping):
            return
        current = dict(getattr(self, name))
        for key, entry in value.items():
            if isinstance(key, str):
                current[key] = bool(entry)
        setattr(self, name, current)

    def _coerce_str_list(self, value: Any, *, fallback: List[str]) -> List[str]:
        if isinstance(value, Iterable) and not isinstance(value, (str, bytes)):
            coerced = [str(item).strip() for item in value if isinstance(item, (str, int, float))]
            coerced = [item for item in coerced if item]
            if coerced:
                return coerced
        return list(fallback)

    def _coerce_int(self, value: Any, default: int) -> int:
        try:
            coerced = int(value)
            if coerced > 0:
                return coerced
        except (TypeError, ValueError):
            pass
        return default

    def as_dict(self) -> Dict[str, Any]:
        """Return a JSON-serialisable view of the configuration."""

        return {
            "api_key": self.api_key,
            "downloads_dir": self.downloads_dir,
            "thread_count": int(self.thread_count),
            "auto_load_last_company": bool(self.auto_load_last_company),
            "last_company": self.last_company,
            "downloads_minutes": int(self.downloads_minutes),
            "note_colors": dict(self.note_colors),
            "scrape_column_widths": dict(self.scrape_column_widths),
            "scrape_row_height": int(self.scrape_row_height),
            "patterns": {k: list(v) for k, v in self.patterns.items()},
            "case_insensitive": dict(self.case_insensitive),
            "space_as_whitespace": dict(self.space_as_whitespace),
            "year_patterns": list(self.year_patterns),
            "year_case_insensitive": bool(self.year_case_insensitive),
            "year_space_as_whitespace": bool(self.year_space_as_whitespace),
            "openai_models": dict(self.openai_models),
            "upload_modes": dict(self.upload_modes),
        }

    def save(self) -> None:
        """Persist the configuration to disk atomically."""

        path = self.config_path()
        payload = json.dumps(self.as_dict(), indent=2)
        tmp_path = path.with_suffix(".tmp")
        with tmp_path.open("w", encoding="utf-8") as fh:
            fh.write(payload)
        tmp_path.replace(path)
