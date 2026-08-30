"""Shared defaults for workbook inspection backends."""

from typing import Literal

DEFAULT_EXCEL_BACKEND: Literal["mac_excel"] = "mac_excel"
DEFAULT_EXCEL_TIMEOUT_SECONDS: int = 120
DEFAULT_EXCEL_MAX_CELLS_PER_CALL: int = 10_000
DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL: int = 1_000
DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS: int = 86_400
EXCEL_MAX_COLUMN_INDEX: int = 16_384
EXCEL_MAX_COLUMN_LABEL: str = "XFD"
EXCEL_AUTH_HEADER: str = "X-Gandalf-Excel-Token"
