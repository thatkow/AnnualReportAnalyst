"""Data model for scrape result tables."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from app_logging import get_logger
from pdf_utils import PDFEntry, normalize_header_row

logger = get_logger()


class ScrapeTableModel:
    """Encapsulate scrape table data operations and state."""

    def __init__(
        self,
        app: "ReportAppV2",
        entry: PDFEntry,
        category: str,
        target_dir: Path,
    ) -> None:
        self.app = app
        self.entry = entry
        self.category = category
        self.target_dir = target_dir

        self.csv_path = target_dir / f"{category}.csv"
        self.multiplier_path = target_dir / f"{category}_multiplier.txt"
        self.raw_path = target_dir / f"{category}_raw.txt"

        self.has_csv_data = False
        self._updating_multiplier = False

        self.row_states: Dict[str, str] = {}
        self.row_keys: Dict[str, Tuple[str, str, str]] = {}

    # ------------------------------------------------------------------
    # CSV helpers
    # ------------------------------------------------------------------
    def load_csv_rows(self) -> Tuple[Optional[List[str]], List[List[str]]]:
        rows: List[List[str]] = []
        if self.csv_path.exists():
            try:
                with self.csv_path.open("r", encoding="utf-8", newline="") as fh:
                    reader = csv.reader(fh)
                    for raw_row in reader:
                        if any(cell.strip() for cell in raw_row):
                            rows.append([cell.strip() for cell in raw_row])
            except OSError:
                rows = []

        header: Optional[List[str]] = None
        data_rows: List[List[str]] = []
        if rows:
            candidate = normalize_header_row(rows[0])
            if candidate is not None:
                header = candidate
                data_rows = rows[1:]
            else:
                data_rows = rows

        self.has_csv_data = bool(data_rows)
        return header, data_rows

    def save_table(self, columns: List[str], rows: List[List[str]]) -> None:
        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            logger.exception(
                "Unable to ensure scrape directory exists: %s", self.target_dir
            )
            return

        try:
            with self.csv_path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, quoting=csv.QUOTE_MINIMAL)
                writer.writerow(columns)
                expected_len = len(columns)
                for values in rows:
                    row = list(values[:expected_len])
                    if len(row) < expected_len:
                        row.extend([""] * (expected_len - len(row)))
                    writer.writerow(row)
        except OSError:
            logger.exception(
                "Unable to persist CSV after table modification for %s - %s",
                self.entry.path.name,
                self.category,
            )

    # ------------------------------------------------------------------
    # Multiplier helpers
    # ------------------------------------------------------------------
    def load_multiplier_value(self) -> str:
        if self.multiplier_path.exists():
            try:
                return self.multiplier_path.read_text(encoding="utf-8").strip()
            except OSError:
                return ""
        return ""

    def save_multiplier_value(self, value: str) -> None:
        if self._updating_multiplier:
            return

        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            logger.exception(
                "Unable to ensure scrape directory exists: %s", self.target_dir
            )
            return

        try:
            if value:
                self.multiplier_path.write_text(value, encoding="utf-8")
            elif self.multiplier_path.exists():
                self.multiplier_path.unlink()
        except OSError:
            logger.exception(
                "Unable to persist multiplier for %s - %s",
                self.entry.path.name,
                self.category,
            )

    def begin_multiplier_update(self) -> None:
        self._updating_multiplier = True

    def end_multiplier_update(self) -> None:
        self._updating_multiplier = False

    def is_updating_multiplier(self) -> bool:
        return self._updating_multiplier

    # ------------------------------------------------------------------
    # Row registration/state helpers
    # ------------------------------------------------------------------
    def clear_row_tracking(self) -> None:
        self.row_states.clear()
        self.row_keys.clear()

    def register_row_key(self, item_id: str, key: Tuple[str, str, str]) -> None:
        self.row_keys[item_id] = key

    def unregister_row(self, item_id: str) -> None:
        self.row_states.pop(item_id, None)
        self.row_keys.pop(item_id, None)

    def set_row_state(self, item_id: str, state: Optional[str]) -> None:
        if state in (None, ""):
            self.row_states.pop(item_id, None)
        else:
            self.row_states[item_id] = state

    def get_row_state(self, item_id: str) -> Optional[str]:
        return self.row_states.get(item_id)


if False:  # pragma: no cover - type checking helper
    from report_app import ReportAppV2  # noqa: F401
