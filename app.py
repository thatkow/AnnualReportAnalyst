import base64
import io
import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter
from dash import ALL, Dash, Input, Output, State, callback_context, dcc, html, no_update
from flask import abort, send_file


COLUMNS = ["Financial", "Income", "Shares"]
DEFAULT_PATTERNS = {
    "Financial": ["statement of financial position"],
    "Income": ["statement of profit or loss"],
    "Shares": ["Movements in issued capital"],
}
BASE_DIR = Path(__file__).resolve().parent
COMPANIES_DIR = BASE_DIR / "companies"


@dataclass
class MatchRecord:
    page_index: int
    source: str
    pattern: Optional[str]
    image_b64: str

    def to_dict(self) -> Dict[str, Optional[str]]:
        return {
            "page_index": self.page_index,
            "source": self.source,
            "pattern": self.pattern,
            "image": self.image_b64,
        }


def _render_page_base64(doc: fitz.Document, page_index: int, *, max_width: int = 420) -> str:
    page = doc.load_page(page_index)
    pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6))
    mode = "RGBA" if pix.alpha else "RGB"
    image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    if image.width > max_width:
        ratio = max_width / image.width
        new_size = (int(image.width * ratio), int(image.height * ratio))
        image = image.resize(new_size, Image.LANCZOS)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return encoded


def _render_page_base64_from_path(pdf_path: Path, page_index: int) -> str:
    with fitz.open(pdf_path) as doc:
        return _render_page_base64(doc, page_index)


def _gather_patterns(pattern_inputs: Dict[str, str]) -> Tuple[Optional[Dict[str, List[re.Pattern[str]]]], Optional[str]]:
    compiled: Dict[str, List[re.Pattern[str]]] = {}
    for column in COLUMNS:
        raw = pattern_inputs.get(column, "")
        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        if not lines:
            lines = DEFAULT_PATTERNS[column]
        column_patterns: List[re.Pattern[str]] = []
        for value in lines:
            try:
                column_patterns.append(re.compile(value, re.IGNORECASE))
            except re.error as exc:
                return None, f"Invalid regex '{value}' for {column}: {exc}"
        compiled[column] = column_patterns
    return compiled, None


def _load_company_data(company: str, pattern_inputs: Dict[str, str]) -> Tuple[Optional[Dict], str]:
    if not company:
        return None, "Please choose a company to scan."

    folder = COMPANIES_DIR / company / "raw"
    if not folder.exists():
        return None, f"Folder '{folder}' does not exist."

    compiled_patterns, error = _gather_patterns(pattern_inputs)
    if error:
        return None, error

    pdf_paths = sorted(folder.rglob("*.pdf"))
    if not pdf_paths:
        return {
            "folder": str(folder),
            "entries": [],
        }, "No PDF files were found for the selected company."

    entries = []
    for pdf_path in pdf_paths:
        try:
            doc = fitz.open(pdf_path)
        except Exception as exc:  # pragma: no cover - safety net
            return None, f"Could not open '{pdf_path.name}': {exc}"
        matches: Dict[str, List[MatchRecord]] = {column: [] for column in COLUMNS}
        try:
            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                text = page.get_text("text")
                for column, patterns in compiled_patterns.items():
                    for pattern in patterns:
                        if pattern.search(text):
                            image_b64 = _render_page_base64(doc, page_index)
                            matches[column].append(
                                MatchRecord(page_index=page_index, source="regex", pattern=pattern.pattern, image_b64=image_b64)
                            )
                            break
        finally:
            doc.close()

        reader = PdfReader(str(pdf_path))
        entries.append(
            {
                "path": str(pdf_path),
                "display_name": str(pdf_path.relative_to(folder)),
                "total_pages": len(reader.pages),
                "matches": {column: [match.to_dict() for match in match_list] for column, match_list in matches.items()},
                "current": {column: (0 if matches[column] else None) for column in COLUMNS},
            }
        )

    return {"folder": str(folder), "entries": entries}, f"Loaded {len(entries)} PDF(s) from {folder}."


def _cycle_match(state: Dict, entry_idx: int, category: str, forward: bool) -> Tuple[Dict, str]:
    state_copy = deepcopy(state)
    entry = state_copy["entries"][entry_idx]
    matches = entry["matches"][category]
    current = entry["current"].get(category)
    if not matches or current is None:
        return state, f"No matches available for {category} in {Path(entry['path']).name}."

    new_index = current + (1 if forward else -1)
    if new_index < 0:
        return state, "Reached the first matched page."
    if new_index >= len(matches):
        return state, "Reached the last matched page."

    entry["current"][category] = new_index
    return state_copy, ""


def _apply_manual_selection(state: Dict, entry_idx: int, category: str, page_number: int) -> Tuple[Dict, str]:
    if page_number <= 0:
        return state, "Page number must be positive."

    state_copy = deepcopy(state)
    entry = state_copy["entries"][entry_idx]
    total_pages = entry["total_pages"]
    if page_number > total_pages:
        return state, f"Page {page_number} exceeds total pages ({total_pages})."

    page_index = page_number - 1
    matches = entry["matches"][category]
    for idx, match in enumerate(matches):
        if match["page_index"] == page_index:
            entry["current"][category] = idx
            return state_copy, ""

    pdf_path = Path(entry["path"])
    image_b64 = _render_page_base64_from_path(pdf_path, page_index)
    matches.append(
        {
            "page_index": page_index,
            "source": "manual",
            "pattern": None,
            "image": image_b64,
        }
    )
    entry["current"][category] = len(matches) - 1
    return state_copy, ""


def _export_selected_pages(state: Dict) -> str:
    folder = Path(state["folder"])
    output_dir = folder / "cut"
    output_dir.mkdir(parents=True, exist_ok=True)

    for entry in state["entries"]:
        reader = PdfReader(entry["path"])
        for category in COLUMNS:
            current_index = entry["current"].get(category)
            matches = entry["matches"].get(category, [])
            if current_index is None or not matches:
                continue
            current_index = max(0, min(current_index, len(matches) - 1))
            match = matches[current_index]
            page_index = match["page_index"]
            if page_index >= len(reader.pages):
                continue
            writer = PdfWriter()
            writer.add_page(reader.pages[page_index])
            output_path = output_dir / f"{Path(entry['path']).stem}_{category}.pdf"
            with output_path.open("wb") as fh:
                writer.write(fh)
    return f"Selected pages saved to '{output_dir}'."


app = Dash(__name__)
app.title = "Annual Report Analyst"


@app.server.route("/pdf/<path:relative_path>")
def serve_pdf(relative_path: str):
    pdf_path = (BASE_DIR / relative_path).resolve()
    if BASE_DIR not in pdf_path.parents and pdf_path != BASE_DIR:
        abort(404)
    if not pdf_path.exists():
        abort(404)
    return send_file(str(pdf_path), mimetype="application/pdf")


def _company_options() -> List[Dict[str, str]]:
    if not COMPANIES_DIR.exists():
        return []
    options = []
    for child in sorted(COMPANIES_DIR.iterdir()):
        if child.is_dir():
            options.append({"label": child.name, "value": child.name})
    return options


def _default_text(column: str) -> str:
    return "\n".join(DEFAULT_PATTERNS[column])


company_options = _company_options()
default_company = company_options[0]["value"] if company_options else None
pattern_defaults = {column: _default_text(column) for column in COLUMNS}
default_state, default_message = (None, "")
if default_company:
    default_state, default_message = _load_company_data(default_company, pattern_defaults)


def _build_cell(entry_idx: int, category: str, entry: Dict) -> html.Div:
    matches: List[Dict] = entry["matches"][category]
    current_index = entry["current"].get(category)
    info_text = "No matches found"
    image_component = html.Div("No preview", className="no-preview")
    if matches and current_index is not None and 0 <= current_index < len(matches):
        match = matches[current_index]
        info_parts = [f"Match {current_index + 1}/{len(matches)}", f"Page {match['page_index'] + 1}"]
        if match["source"] == "manual":
            info_parts.append("(manual)")
        if match.get("pattern"):
            info_parts.append(match["pattern"])
        info_text = " | ".join(info_parts)
        image_component = html.Img(
            src=f"data:image/png;base64,{match['image']}",
            id={"type": "nav", "entry": entry_idx, "category": category, "action": "next"},
            className="page-preview",
            title="Click to go to the next match",
        )

    controls = html.Div(
        [
            html.Button(
                "Prev",
                id={"type": "nav", "entry": entry_idx, "category": category, "action": "prev"},
                className="nav-button",
            ),
            html.Button(
                "Next",
                id={"type": "nav", "entry": entry_idx, "category": category, "action": "next_button"},
                className="nav-button",
            ),
        ],
        className="nav-controls",
    )

    manual_controls = html.Div(
        [
            dcc.Input(
                id={"type": "manual-input", "entry": entry_idx, "category": category},
                type="number",
                min=1,
                placeholder="Page #",
                className="manual-input",
            ),
            html.Button(
                "Set Page",
                id={"type": "manual-set", "entry": entry_idx, "category": category},
                className="nav-button",
            ),
            html.A(
                "Open PDF",
                href=f"/pdf/{Path(entry['path']).relative_to(BASE_DIR)}",
                target="_blank",
                className="open-link",
            ),
        ],
        className="manual-controls",
    )

    return html.Div(
        [
            image_component,
            html.Div(info_text, className="cell-info"),
            controls,
            manual_controls,
        ],
        className="grid-cell",
    )


app.layout = html.Div(
    [
        html.Style(
            """
            body { font-family: Arial, sans-serif; background-color: #f4f6f8; margin: 0; }
            .app-container { max-width: 1400px; margin: 0 auto; padding: 24px; }
            .controls { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 16px; align-items: center; }
            .controls > * { flex: 1 1 200px; }
            .pattern-area { width: 100%; min-height: 120px; }
            .pattern-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px; margin-bottom: 12px; }
            .grid { display: grid; grid-template-columns: 1.2fr repeat(3, 1fr); gap: 12px; }
            .grid-header { font-weight: bold; }
            .grid-row { background: #fff; border-radius: 8px; padding: 12px; display: contents; }
            .grid-label { background: #fff; border-radius: 8px; padding: 12px; display: flex; align-items: center; font-weight: 600; }
            .grid-cell { background: #fff; border-radius: 8px; padding: 12px; display: flex; flex-direction: column; gap: 8px; align-items: center; }
            .page-preview { max-width: 100%; border-radius: 4px; box-shadow: 0 2px 6px rgba(0,0,0,0.12); cursor: pointer; }
            .cell-info { font-size: 13px; text-align: center; color: #374151; }
            .nav-controls, .manual-controls { display: flex; gap: 8px; justify-content: center; flex-wrap: wrap; }
            .nav-button { padding: 6px 12px; border-radius: 4px; border: 1px solid #d1d5db; background: #e5e7eb; cursor: pointer; }
            .nav-button:hover { background: #d1d5db; }
            .manual-input { width: 90px; padding: 4px 6px; }
            .open-link { align-self: center; font-size: 13px; }
            .no-preview { font-size: 13px; color: #6b7280; }
            .status { margin-top: 16px; font-size: 14px; color: #2563eb; }
            .confirm-container { margin-top: 24px; display: flex; gap: 12px; align-items: center; }
            .confirm-button { padding: 8px 16px; border-radius: 6px; border: none; background: #2563eb; color: #fff; cursor: pointer; }
            .confirm-button:disabled { background: #9ca3af; cursor: not-allowed; }
            .header { margin-bottom: 16px; }
            """
        ),
        html.Div(
            [
                html.H1("Annual Report Analyst", className="header"),
                html.Div(
                    [
                        dcc.Dropdown(
                            id="company-dropdown",
                            options=company_options,
                            placeholder="Select a company",
                            value=default_company,
                            clearable=False,
                        ),
                        html.Button("Apply Patterns", id="apply-button", className="nav-button"),
                    ],
                    className="controls",
                ),
                html.Div(
                    [
                        html.Div(
                            [
                                html.Label(column),
                                dcc.Textarea(
                                    id={"type": "pattern", "column": column},
                                    value=pattern_defaults[column],
                                    className="pattern-area",
                                ),
                            ],
                            className="pattern-block",
                        )
                        for column in COLUMNS
                    ],
                    className="pattern-container",
                ),
                html.Div(id="grid-container", className="grid"),
                html.Div(
                    [
                        html.Button("Confirm", id="confirm-button", className="confirm-button"),
                        html.Div(default_message, id="status-message", className="status"),
                    ],
                    className="confirm-container",
                ),
                dcc.Store(id="app-state", data=default_state),
            ],
            className="app-container",
        ),
    ]
)


@app.callback(
    Output("app-state", "data"),
    Output("status-message", "children"),
    Input("apply-button", "n_clicks"),
    State("company-dropdown", "value"),
    State({"type": "pattern", "column": ALL}, "value"),
    prevent_initial_call=True,
)
def load_data(_n_clicks, company, pattern_values):
    pattern_inputs = {column: pattern_values[idx] if pattern_values and idx < len(pattern_values) else "" for idx, column in enumerate(COLUMNS)}
    state, message = _load_company_data(company, pattern_inputs)
    return state, message


@app.callback(
    Output("grid-container", "children"),
    Input("app-state", "data"),
)
def rebuild_grid(state):
    headers = [html.Div("PDF", className="grid-header")] + [html.Div(column, className="grid-header") for column in COLUMNS]
    if not state or not state.get("entries"):
        return [html.Div(headers, className="grid-row")]

    rows: List[html.Div] = [html.Div(headers, className="grid-row")]
    for entry_idx, entry in enumerate(state["entries"]):
        label = html.Div(entry["display_name"], className="grid-label")
        cells = [label]
        for category in COLUMNS:
            cells.append(_build_cell(entry_idx, category, entry))
        rows.append(html.Div(cells, className="grid-row"))
    return rows


@app.callback(
    Output("app-state", "data", allow_duplicate=True),
    Output("status-message", "children", allow_duplicate=True),
    Input({"type": "nav", "entry": ALL, "category": ALL, "action": ALL}, "n_clicks"),
    State("app-state", "data"),
    prevent_initial_call=True,
)
def handle_navigation(_clicks, state):
    if not state:
        return no_update, no_update
    ctx = callback_context
    if not ctx.triggered:
        return no_update, no_update
    triggered_id = ctx.triggered_id
    if not isinstance(triggered_id, dict):
        return no_update, no_update
    entry_idx = int(triggered_id.get("entry", 0))
    category = triggered_id.get("category")
    action = triggered_id.get("action")
    forward = action in ("next", "next_button")
    if category not in COLUMNS:
        return no_update, no_update
    new_state, message = _cycle_match(state, entry_idx, category, forward=forward)
    if new_state is state:
        return no_update, message or no_update
    return new_state, message


@app.callback(
    Output("app-state", "data", allow_duplicate=True),
    Output("status-message", "children", allow_duplicate=True),
    Input({"type": "manual-set", "entry": ALL, "category": ALL}, "n_clicks"),
    State({"type": "manual-input", "entry": ALL, "category": ALL}, "value"),
    State("app-state", "data"),
    prevent_initial_call=True,
)
def handle_manual_selection(_clicks, values, state):
    if not state:
        return no_update, no_update
    ctx = callback_context
    if not ctx.triggered:
        return no_update, no_update
    triggered_id = ctx.triggered_id
    if not isinstance(triggered_id, dict):
        return no_update, no_update
    entry_idx = int(triggered_id.get("entry", 0))
    category = triggered_id.get("category")
    if category not in COLUMNS:
        return no_update, no_update

    index = None
    if values:
        total_categories = len(COLUMNS)
        index = entry_idx * total_categories + COLUMNS.index(category)
        if index >= len(values):
            index = None
    page_value = values[index] if values and index is not None else None
    if page_value is None:
        return no_update, "Enter a page number before applying."
    new_state, message = _apply_manual_selection(state, entry_idx, category, int(page_value))
    if new_state is state:
        return no_update, message or no_update
    return new_state, message


@app.callback(
    Output("status-message", "children", allow_duplicate=True),
    Input("confirm-button", "n_clicks"),
    State("app-state", "data"),
    prevent_initial_call=True,
)
def confirm_selection(_n_clicks, state):
    if not state:
        return "Load PDFs before confirming selections."
    message = _export_selected_pages(state)
    return message


if __name__ == "__main__":
    app.run_server(debug=False, host="0.0.0.0")
