"""Deterministic workbook digest for the judge prompt.

Builds a compact, structural summary of a workbook using openpyxl (no Microsoft
Excel, no recalculation). Injected once into the judge prompt so the judge agent
can locate sheets, headers, and formulas without spending tool-call turns
re-discovering structure for every rubric criterion.

Values are intentionally not included: openpyxl returns only cached values
(often ``None`` for agent-written workbooks), so the agent should still use the
Excel MCP tools for actual computed values. The digest is a navigation aid.
"""

from __future__ import annotations

import importlib
from pathlib import Path
from typing import Any

_DIGEST_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
_MAX_CELLS_SCANNED = 20_000
_MAX_CELL_TEXT = 60


def build_workbook_digest(
    path: Path,
    *,
    max_sheets: int = 30,
    max_formula_samples: int = 15,
    max_preview_columns: int = 20,
    max_named_ranges: int = 50,
) -> str | None:
    """Return a compact structural digest of *path*, or ``None`` if unavailable.

    Returns ``None`` for unsupported extensions (``.xls``/``.xlsb``) or any read
    error, so callers can fall back to no digest without failing the run.
    """
    if path.suffix.lower() not in _DIGEST_EXTENSIONS:
        return None
    try:
        openpyxl = importlib.import_module("openpyxl")
        workbook = openpyxl.load_workbook(path, data_only=False)
    except Exception:  # noqa: BLE001 - any load failure falls back to no digest
        return None

    try:
        sheets = workbook.worksheets
        lines: list[str] = [f"Sheets ({len(sheets)}):"]
        for worksheet in sheets[:max_sheets]:
            lines.extend(_sheet_lines(worksheet, max_formula_samples, max_preview_columns))
        if len(sheets) > max_sheets:
            lines.append(f"- ... and {len(sheets) - max_sheets} more sheet(s) not shown")
        named_ranges = _named_ranges(workbook, max_named_ranges)
        if named_ranges:
            lines.append(f"Named ranges ({len(named_ranges)} shown):")
            lines.extend(f"- {name} -> {destination}" for name, destination in named_ranges)
        return "\n".join(lines)
    finally:
        close = getattr(workbook, "close", None)
        if callable(close):
            close()


def _sheet_lines(worksheet: Any, max_formula_samples: int, max_preview_columns: int) -> list[str]:
    dimensions = str(worksheet.calculate_dimension())
    lines = [
        f'- "{worksheet.title}" ({worksheet.sheet_state}) dims {dimensions} '
        f"max_row={worksheet.max_row} max_col={worksheet.max_column}"
    ]
    preview = _header_preview(worksheet, max_preview_columns)
    if preview:
        lines.append(f"    row1: {preview}")
    for coordinate, formula in _formula_samples(worksheet, max_formula_samples):
        lines.append(f"    formula {coordinate}: {formula}")
    return lines


def _header_preview(worksheet: Any, max_preview_columns: int) -> str:
    # Cap to the existing max_column so reading the preview never extends the
    # worksheet dimension (openpyxl materialises cells on access).
    width = min(max_preview_columns, worksheet.max_column)
    if width < 1 or worksheet.max_row < 1:
        return ""
    first_row = next(worksheet.iter_rows(min_row=1, max_row=1, max_col=width), ())
    cells = [_cell_text(cell.value) for cell in first_row]
    return " | ".join(cells) if any(cells) else ""


def _formula_samples(worksheet: Any, max_formula_samples: int) -> list[tuple[str, str]]:
    if max_formula_samples <= 0:
        return []
    samples: list[tuple[str, str]] = []
    scanned = 0
    for row in worksheet.iter_rows():
        for cell in row:
            scanned += 1
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                samples.append((cell.coordinate, _cell_text(value)))
                if len(samples) >= max_formula_samples:
                    return samples
        if scanned >= _MAX_CELLS_SCANNED:
            break
    return samples


def _named_ranges(workbook: Any, max_named_ranges: int) -> list[tuple[str, str]]:
    if max_named_ranges <= 0:
        return []
    names: list[tuple[str, str]] = []
    try:
        items = workbook.defined_names.items()
    except Exception:  # noqa: BLE001 - older/odd openpyxl shapes simply skip names
        return []
    for name, defined_name in items:
        destination = getattr(defined_name, "value", None) or str(defined_name)
        names.append((str(name), _cell_text(str(destination))))
        if len(names) >= max_named_ranges:
            break
    return names


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").strip()
    if len(text) > _MAX_CELL_TEXT:
        return text[: _MAX_CELL_TEXT - 1] + "…"
    return text
