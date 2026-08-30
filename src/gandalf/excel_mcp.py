"""Microsoft Excel-backed MCP tools for workbook inspection.

The MCP server is intentionally small: it exposes a safe JSON interface over
Excel-backed workbook facts, while all rubric scoring remains in the grader.
"""

from __future__ import annotations

import argparse
import atexit
import base64
import copy
import hashlib
import importlib
import json
import os
import re
import secrets
import shutil
import subprocess
import tempfile
import threading
import time as time_module
import urllib.error
import urllib.request
from collections import OrderedDict
from collections.abc import Callable, Iterator, Sequence
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from datetime import date, datetime, time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Protocol

from fastmcp import FastMCP

from gandalf.excel_defaults import (
    DEFAULT_EXCEL_BACKEND,
    DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
    DEFAULT_EXCEL_MAX_CELLS_PER_CALL,
    DEFAULT_EXCEL_TIMEOUT_SECONDS,
    DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS,
    EXCEL_AUTH_HEADER,
    EXCEL_MAX_COLUMN_INDEX,
    EXCEL_MAX_COLUMN_LABEL,
)

JsonDict = dict[str, Any]

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")
_RANGE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$")
_MACRO_EXTENSIONS = {".xls", ".xlsb", ".xlsm", ".xlam"}
_EXCEL_EXTENSIONS = {".xls", ".xlsb", ".xlsm", ".xlsx", ".xltx", ".xltm"}
_OPENPYXL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
_XL_NORMAL_LOAD = 0
_XL_REPAIR_FILE = 1
_REPAIR_ERROR_MARKERS = (
    "repair",
    "repaired",
    "corrupt",
    "corrupted",
    "unreadable",
    "removed records",
    "recovered",
)
_CORRUPT_LOAD_UNSUPPORTED_MARKER = "corrupt_load is not supported"
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_LIBREOFFICE_PREPARED_WORKBOOK_NAME = "prepared.xlsx"
_LIBREOFFICE_LOADED_WORKBOOK_CACHE_SIZE = 4
_LIBREOFFICE_LOCK_POLL_SECONDS = 0.05
_EXCEL_METRIC_SUM_FIELDS = (
    "windows_vm_client_request_ms",
    "windows_vm_client_total_ms",
    "windows_vm_client_payload_bytes",
    "windows_vm_client_workbook_bytes_sent",
    "windows_vm_client_ensure_ms",
    "windows_vm_worker_total_ms",
    "windows_vm_worker_operation_ms",
    "windows_vm_worker_cache_prepare_ms",
    "windows_vm_worker_uploaded_bytes",
    "libreoffice_conversion_ms",
    "libreoffice_openpyxl_load_ms",
    "libreoffice_hash_ms",
    "libreoffice_cache_lock_wait_ms",
)


@dataclass
class _WorkerCacheLockState:
    """Lock plus waiter/user count for one content-addressed cache key."""

    lock: threading.Lock
    references: int = 0


@dataclass(frozen=True)
class _LibreOfficePreparedWorkbook:
    """A workbook path prepared for openpyxl reads."""

    path: Path
    converted: bool
    recalculated: bool
    metrics: JsonDict


@dataclass(frozen=True)
class _LibreOfficeCacheKey:
    """Stable key for a LibreOffice prepared-workbook cache."""

    content_sha256: str
    suffix: str
    force_recalculate: bool


@dataclass
class _LibreOfficeOpenWorkbookPair:
    """Formula/value openpyxl workbooks loaded from one prepared workbook."""

    formula_workbook: Any
    value_workbook: Any


@dataclass(frozen=True)
class _WorkbookStatKey:
    """Local workbook identity used for short-lived operation caches."""

    path: str
    size: int
    mtime_ns: int


@dataclass(frozen=True)
class _RemoteWorkbookRef:
    """A workbook already known to the Windows VM worker."""

    workbook_name: str
    sha256: str
    legacy_full_upload: bool = False


@dataclass(frozen=True)
class _WorkerWorkbookSession:
    """Resolved worker-side cached workbook for one request."""

    workbook_path: Path
    sha256: str
    cache_hit: bool
    uploaded_bytes: int
    cache_prepare_ms: float


class ExcelMetricsCollector:
    """Collect aggregate Excel tool diagnostics for one MCP server process."""

    def __init__(self, metrics_path: Path | None = None) -> None:
        self.metrics_path = metrics_path
        self._lock = threading.Lock()
        self._summary = empty_excel_metrics_summary()

    def record(self, operation: str, metrics: JsonDict) -> None:
        with self._lock:
            update_excel_metrics_summary(self._summary, operation, metrics)
            if self.metrics_path is not None:
                write_excel_metrics_summary(self.metrics_path, self._summary)

    def summary(self) -> JsonDict:
        with self._lock:
            return copy.deepcopy(self._summary)


_WORKER_CACHE_LOCKS: dict[str, _WorkerCacheLockState] = {}
_WORKER_CACHE_LOCKS_GUARD = threading.Lock()


class ExcelBackend(Protocol):
    """Backend contract for workbook inspection."""

    def workbook_summary(self, workbook_path: Path, *, recalculate: bool) -> JsonDict:
        """Return workbook-level metadata."""

    def inspect_range(
        self,
        workbook_path: Path,
        *,
        sheet: str,
        range_address: str,
        include_values: bool,
        include_formulas: bool,
        include_formats: bool,
        recalculate: bool,
    ) -> JsonDict:
        """Return facts for a single Excel range."""

    def find_cells(
        self,
        workbook_path: Path,
        *,
        query: str,
        search_values: bool,
        search_formulas: bool,
        max_results: int,
        max_cells_per_call: int,
    ) -> JsonDict:
        """Find cells whose values or formulas contain *query*."""

    def recalculate(self, workbook_path: Path) -> JsonDict:
        """Recalculate a workbook."""

    def workbook_repair_check(self, workbook_path: Path) -> JsonDict:
        """Detect whether Microsoft Excel must repair the workbook before opening it."""


@dataclass(frozen=True)
class ExcelServerConfig:
    """Runtime settings consumed by the MCP server."""

    workdir: Path
    backend: str = DEFAULT_EXCEL_BACKEND
    windows_url: str | None = None
    auth_token: str | None = None
    timeout_seconds: int = DEFAULT_EXCEL_TIMEOUT_SECONDS
    visible: bool = False
    allow_macros: bool = False
    max_cells_per_call: int = DEFAULT_EXCEL_MAX_CELLS_PER_CALL
    max_format_cells_per_call: int = DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL
    metrics_path: Path | None = None
    debug_metrics: bool = False
    cache_dir: Path | None = None


class ExcelService:
    """Validated operations exposed through MCP."""

    def __init__(
        self,
        *,
        workdir: Path,
        backend: ExcelBackend,
        allow_macros: bool,
        max_cells_per_call: int,
        max_format_cells_per_call: int = DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
        metrics_path: Path | None = None,
        debug_metrics: bool = False,
    ) -> None:
        require_positive_int(max_cells_per_call, "max_cells_per_call")
        require_positive_int(max_format_cells_per_call, "max_format_cells_per_call")
        self.workdir = workdir.resolve()
        self.backend = backend
        self.allow_macros = allow_macros
        self.max_cells_per_call = max_cells_per_call
        self.max_format_cells_per_call = max_format_cells_per_call
        self._cache_lock = threading.Lock()
        self._operation_cache: dict[tuple[Any, ...], JsonDict] = {}
        self._metrics = ExcelMetricsCollector(metrics_path)
        self.debug_metrics = debug_metrics

    def workbook_summary(self, path: str, *, recalculate: bool = True) -> JsonDict:
        workbook_path = self._resolve_workbook_path(path)
        cache_key = ("workbook_summary", workbook_stat_key(workbook_path), bool(recalculate))
        return self._cached_operation(
            cache_key,
            lambda: self.backend.workbook_summary(workbook_path, recalculate=recalculate),
        )

    def inspect_range(
        self,
        path: str,
        *,
        sheet: str,
        range_address: str,
        include_values: bool = True,
        include_formulas: bool = True,
        include_formats: bool = False,
        recalculate: bool = True,
    ) -> JsonDict:
        workbook_path = self._resolve_workbook_path(path)
        cell_count = range_cell_count(range_address)
        if cell_count > self.max_cells_per_call:
            msg = (
                f"Range {range_address!r} contains {cell_count} cells, exceeding "
                f"max_cells_per_call={self.max_cells_per_call}"
            )
            raise ValueError(msg)
        if include_formats and cell_count > self.max_format_cells_per_call:
            msg = (
                f"Range {range_address!r} contains {cell_count} cells, exceeding "
                f"max_format_cells_per_call={self.max_format_cells_per_call}; "
                "pass include_formats=false or request a smaller range"
            )
            raise ValueError(msg)
        cache_key = (
            "inspect_range",
            workbook_stat_key(workbook_path),
            sheet,
            range_address,
            bool(include_values),
            bool(include_formulas),
            bool(include_formats),
            bool(recalculate),
        )
        return self._cached_operation(
            cache_key,
            lambda: self.backend.inspect_range(
                workbook_path,
                sheet=sheet,
                range_address=range_address,
                include_values=include_values,
                include_formulas=include_formulas,
                include_formats=include_formats,
                recalculate=recalculate,
            ),
        )

    def find_cells(
        self,
        path: str,
        *,
        query: str,
        search_values: bool = True,
        search_formulas: bool = True,
        max_results: int = 100,
    ) -> JsonDict:
        if not query:
            msg = "query cannot be empty"
            raise ValueError(msg)
        if max_results < 1:
            msg = "max_results must be at least 1"
            raise ValueError(msg)
        workbook_path = self._resolve_workbook_path(path)
        cache_key = (
            "find_cells",
            workbook_stat_key(workbook_path),
            query,
            bool(search_values),
            bool(search_formulas),
            int(max_results),
            self.max_cells_per_call,
        )
        return self._cached_operation(
            cache_key,
            lambda: self.backend.find_cells(
                workbook_path,
                query=query,
                search_values=search_values,
                search_formulas=search_formulas,
                max_results=max_results,
                max_cells_per_call=self.max_cells_per_call,
            ),
        )

    def recalculate(self, path: str) -> JsonDict:
        workbook_path = self._resolve_workbook_path(path)
        result = self.backend.recalculate(workbook_path)
        self._clear_operation_cache()
        return self._finalise_result("recalculate", copy.deepcopy(result))

    def workbook_repair_check(self, path: str) -> JsonDict:
        workbook_path = self._resolve_workbook_path(path)
        return self._finalise_result(
            "workbook_repair_check",
            copy.deepcopy(self.backend.workbook_repair_check(workbook_path)),
        )

    def _resolve_workbook_path(self, path: str) -> Path:
        raw = Path(path)
        resolved = raw.resolve() if raw.is_absolute() else (self.workdir / raw).resolve()
        try:
            resolved.relative_to(self.workdir)
        except ValueError as e:
            msg = f"Workbook path {path!r} is outside the allowed workdir"
            raise ValueError(msg) from e
        if not resolved.is_file():
            msg = f"Workbook path {path!r} does not exist"
            raise FileNotFoundError(msg)
        suffix = resolved.suffix.lower()
        if suffix not in _EXCEL_EXTENSIONS:
            msg = f"Workbook path {path!r} is not an Excel workbook"
            raise ValueError(msg)
        if not self.allow_macros and suffix in _MACRO_EXTENSIONS:
            msg = f"Workbook path {path!r} may contain macros; set allow_macros=true to inspect it"
            raise ValueError(msg)
        return resolved

    def _cached_operation(self, cache_key: tuple[Any, ...], compute: Callable[[], JsonDict]) -> JsonDict:
        with self._cache_lock:
            cached = self._operation_cache.get(cache_key)
        if cached is not None:
            result = copy.deepcopy(cached)
            return self._finalise_result(
                str(cache_key[0]),
                result,
                {"excel_service_cache_hit": True},
            )

        result = copy.deepcopy(compute())
        metrics = pop_result_metrics(result)
        metrics["excel_service_cache_hit"] = False
        with self._cache_lock:
            self._operation_cache[cache_key] = copy.deepcopy(result)
        return self._finalise_result(str(cache_key[0]), result, metrics)

    def _clear_operation_cache(self) -> None:
        with self._cache_lock:
            self._operation_cache.clear()

    def _finalise_result(
        self,
        operation: str,
        result: JsonDict,
        extra_metrics: JsonDict | None = None,
    ) -> JsonDict:
        metrics = pop_result_metrics(result)
        if extra_metrics:
            metrics.update(extra_metrics)
        self._metrics.record(operation, metrics)
        if metrics and self.debug_metrics:
            result["metrics"] = metrics
        return result


def workbook_stat_key(workbook_path: Path) -> _WorkbookStatKey:
    """Return a cheap local identity for workbook cache keys."""
    resolved = workbook_path.resolve()
    stat = resolved.stat()
    return _WorkbookStatKey(path=str(resolved), size=stat.st_size, mtime_ns=stat.st_mtime_ns)


def add_result_metrics(result: JsonDict, metrics: JsonDict) -> None:
    """Attach small diagnostic metrics without overwriting existing result fields."""
    existing = result.get("metrics")
    if not isinstance(existing, dict):
        existing = {}
        result["metrics"] = existing
    existing.update(metrics)


def pop_result_metrics(result: JsonDict) -> JsonDict:
    """Remove and return a result's diagnostic metrics, if present."""
    raw_metrics = result.pop("metrics", None)
    return dict(raw_metrics) if isinstance(raw_metrics, dict) else {}


def empty_excel_metrics_summary() -> JsonDict:
    """Return an empty aggregate metrics object."""
    return {
        "call_count": 0,
        "operation_counts": {},
        "excel_service_cache_hits": 0,
        "excel_service_cache_misses": 0,
        "windows_vm_client_workbook_cache_hits": 0,
        "windows_vm_client_workbook_cache_misses": 0,
        "windows_vm_worker_workbook_cache_hits": 0,
        "windows_vm_worker_workbook_cache_misses": 0,
        "windows_vm_client_cache_miss_retries": 0,
        "windows_vm_client_legacy_full_uploads": 0,
        "libreoffice_prepared_cache_hits": 0,
        "libreoffice_prepared_cache_misses": 0,
        "libreoffice_loaded_workbook_cache_hits": 0,
        "libreoffice_loaded_workbook_cache_misses": 0,
    }


def increment_summary_counter(summary: JsonDict, key: str, amount: int = 1) -> None:
    """Increment an integer counter in an aggregate metrics object."""
    current = summary.get(key, 0)
    if not isinstance(current, int):
        current = 0
    summary[key] = current + amount


def add_summary_float(summary: JsonDict, key: str, amount: float) -> None:
    """Add a numeric amount to an aggregate metrics object."""
    current = summary.get(key, 0.0)
    if not isinstance(current, (int, float)):
        current = 0.0
    summary[key] = round(float(current) + amount, 3)


def update_excel_metrics_summary(summary: JsonDict, operation: str, metrics: JsonDict) -> None:
    """Add one Excel tool call's diagnostics to an aggregate metrics object."""
    increment_summary_counter(summary, "call_count")
    operation_counts = summary.setdefault("operation_counts", {})
    if not isinstance(operation_counts, dict):
        operation_counts = {}
        summary["operation_counts"] = operation_counts
    operation_counts[operation] = int(operation_counts.get(operation, 0)) + 1

    service_cache_hit = metrics.get("excel_service_cache_hit")
    if service_cache_hit is True:
        increment_summary_counter(summary, "excel_service_cache_hits")
    elif service_cache_hit is False:
        increment_summary_counter(summary, "excel_service_cache_misses")

    client_cache_hit = metrics.get("windows_vm_client_workbook_cache_hit")
    if client_cache_hit is True:
        increment_summary_counter(summary, "windows_vm_client_workbook_cache_hits")
    elif client_cache_hit is False:
        increment_summary_counter(summary, "windows_vm_client_workbook_cache_misses")

    worker_cache_hit = metrics.get("windows_vm_worker_workbook_cache_hit")
    if worker_cache_hit is True:
        increment_summary_counter(summary, "windows_vm_worker_workbook_cache_hits")
    elif worker_cache_hit is False:
        increment_summary_counter(summary, "windows_vm_worker_workbook_cache_misses")

    if metrics.get("windows_vm_client_retried_after_cache_miss") is True:
        increment_summary_counter(summary, "windows_vm_client_cache_miss_retries")
    if metrics.get("windows_vm_client_legacy_full_upload") is True:
        increment_summary_counter(summary, "windows_vm_client_legacy_full_uploads")

    prepared_cache_hit = metrics.get("libreoffice_prepared_cache_hit")
    if prepared_cache_hit is True:
        increment_summary_counter(summary, "libreoffice_prepared_cache_hits")
    elif prepared_cache_hit is False:
        increment_summary_counter(summary, "libreoffice_prepared_cache_misses")

    loaded_cache_hit = metrics.get("libreoffice_loaded_workbook_cache_hit")
    if loaded_cache_hit is True:
        increment_summary_counter(summary, "libreoffice_loaded_workbook_cache_hits")
    elif loaded_cache_hit is False:
        increment_summary_counter(summary, "libreoffice_loaded_workbook_cache_misses")

    for field in _EXCEL_METRIC_SUM_FIELDS:
        value = metrics.get(field)
        if isinstance(value, (int, float)):
            add_summary_float(summary, field, float(value))


def merge_excel_metrics_summaries(summaries: Sequence[JsonDict]) -> JsonDict:
    """Merge aggregate metrics from multiple MCP server processes."""
    merged = empty_excel_metrics_summary()
    for summary in summaries:
        if not isinstance(summary, dict):
            continue
        increment_summary_counter(merged, "call_count", int(summary.get("call_count", 0)))
        operation_counts = summary.get("operation_counts")
        merged_operation_counts = merged.setdefault("operation_counts", {})
        if isinstance(operation_counts, dict) and isinstance(merged_operation_counts, dict):
            for operation, count in operation_counts.items():
                if isinstance(operation, str) and isinstance(count, int):
                    merged_operation_counts[operation] = int(merged_operation_counts.get(operation, 0)) + count
        for key in (
            "excel_service_cache_hits",
            "excel_service_cache_misses",
            "windows_vm_client_workbook_cache_hits",
            "windows_vm_client_workbook_cache_misses",
            "windows_vm_worker_workbook_cache_hits",
            "windows_vm_worker_workbook_cache_misses",
            "windows_vm_client_cache_miss_retries",
            "windows_vm_client_legacy_full_uploads",
            "libreoffice_prepared_cache_hits",
            "libreoffice_prepared_cache_misses",
            "libreoffice_loaded_workbook_cache_hits",
            "libreoffice_loaded_workbook_cache_misses",
        ):
            value = summary.get(key)
            if isinstance(value, int):
                increment_summary_counter(merged, key, value)
        for field in _EXCEL_METRIC_SUM_FIELDS:
            value = summary.get(field)
            if isinstance(value, (int, float)):
                add_summary_float(merged, field, float(value))
    return merged


def write_excel_metrics_summary(metrics_path: Path, summary: JsonDict) -> None:
    """Write an aggregate metrics file atomically enough for one MCP process."""
    metrics_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = metrics_path.with_suffix(metrics_path.suffix + ".tmp")
    temp_path.write_text(json.dumps(summary, indent=2))
    os.replace(temp_path, metrics_path)
    # Make the sidecar world-readable so a sandbox_user judge's metrics can be
    # collected by the grader user from the shared clone directory.
    with suppress(OSError):
        os.chmod(metrics_path, 0o644)


def merge_windows_vm_ensure_metrics(previous: JsonDict | None, current: JsonDict) -> JsonDict:
    """Merge repeated Windows VM ensure metrics after operation-time cache recovery."""
    if previous is None:
        return dict(current)
    merged = dict(current)
    previous_bytes = previous.get("windows_vm_client_workbook_bytes_sent", 0)
    current_bytes = current.get("windows_vm_client_workbook_bytes_sent", 0)
    if isinstance(previous_bytes, int) and isinstance(current_bytes, int):
        merged["windows_vm_client_workbook_bytes_sent"] = previous_bytes + current_bytes
    previous_ms = previous.get("windows_vm_client_ensure_ms", 0.0)
    current_ms = current.get("windows_vm_client_ensure_ms", 0.0)
    if isinstance(previous_ms, (int, float)) and isinstance(current_ms, (int, float)):
        merged["windows_vm_client_ensure_ms"] = round(float(previous_ms) + float(current_ms), 3)
    merged["windows_vm_client_previous_workbook_cache_hit"] = previous.get("windows_vm_client_workbook_cache_hit")
    return merged


class LocalExcelBackend:
    """Desktop Microsoft Excel backend powered by xlwings."""

    def __init__(self, *, visible: bool = False) -> None:
        self.visible = visible

    @contextmanager
    def _open_workbook(self, workbook_path: Path, *, recalculate: bool) -> Any:
        try:
            xw = importlib.import_module("xlwings")
        except ImportError as e:
            msg = "xlwings is required for the local Microsoft Excel backend"
            raise RuntimeError(msg) from e

        app = xw.App(visible=self.visible, add_book=False)
        workbook = None
        try:
            app.display_alerts = False
            app.screen_updating = False
            workbook = app.books.open(
                str(workbook_path),
                update_links=False,
                read_only=True,
                ignore_read_only_recommended=True,
            )
            if recalculate:
                app.calculate()
            yield workbook
        finally:
            if workbook is not None:
                workbook.close()
            app.quit()

    def workbook_summary(self, workbook_path: Path, *, recalculate: bool) -> JsonDict:
        with self._open_workbook(workbook_path, recalculate=recalculate) as workbook:
            sheets: list[JsonDict] = []
            for sheet in workbook.sheets:
                used = sheet.used_range
                sheets.append(
                    {
                        "name": sheet.name,
                        "used_range": used.address,
                        "rows": used.rows.count,
                        "columns": used.columns.count,
                    }
                )
            return {
                "path": workbook_path.name,
                "backend": "mac_excel",
                "recalculated": recalculate,
                "sheets": sheets,
            }

    def inspect_range(
        self,
        workbook_path: Path,
        *,
        sheet: str,
        range_address: str,
        include_values: bool,
        include_formulas: bool,
        include_formats: bool,
        recalculate: bool,
    ) -> JsonDict:
        rows, columns = range_shape(range_address)
        with self._open_workbook(workbook_path, recalculate=recalculate) as workbook:
            try:
                worksheet = workbook.sheets[sheet]
            except Exception as e:  # noqa: BLE001
                msg = f"Worksheet {sheet!r} does not exist"
                raise ValueError(msg) from e
            cell_range = worksheet.range(range_address)
            result: JsonDict = {
                "path": workbook_path.name,
                "backend": "mac_excel",
                "sheet": sheet,
                "range": range_address,
                "rows": rows,
                "columns": columns,
                "recalculated": recalculate,
            }
            if include_values:
                result["values"] = to_2d(cell_range.options(ndim=2).value, rows=rows, columns=columns)
            if include_formulas:
                result["formulas"] = to_2d(cell_range.options(ndim=2).formula, rows=rows, columns=columns)
            if include_formats:
                result["number_formats"] = range_number_formats(
                    worksheet,
                    cell_range,
                    rows=rows,
                    columns=columns,
                )
            return result

    def find_cells(
        self,
        workbook_path: Path,
        *,
        query: str,
        search_values: bool,
        search_formulas: bool,
        max_results: int,
        max_cells_per_call: int,
    ) -> JsonDict:
        query_lc = query.lower()
        matches: list[JsonDict] = []
        inspected_cells = 0
        with self._open_workbook(workbook_path, recalculate=True) as workbook:
            for worksheet in workbook.sheets:
                used = worksheet.used_range
                rows = used.rows.count
                columns = used.columns.count
                inspected_cells += rows * columns
                if inspected_cells > max_cells_per_call:
                    msg = f"Workbook search exceeds max_cells_per_call={max_cells_per_call}"
                    raise ValueError(msg)
                values = to_2d(used.options(ndim=2).value, rows=rows, columns=columns)
                formulas = to_2d(used.options(ndim=2).formula, rows=rows, columns=columns)
                for row_idx in range(rows):
                    for col_idx in range(columns):
                        value = values[row_idx][col_idx]
                        formula = formulas[row_idx][col_idx]
                        haystacks: list[tuple[str, Any]] = []
                        if search_values:
                            haystacks.append(("value", value))
                        if search_formulas:
                            haystacks.append(("formula", formula))
                        for match_type, haystack in haystacks:
                            if haystack is not None and query_lc in str(haystack).lower():
                                matches.append(
                                    {
                                        "sheet": worksheet.name,
                                        "cell": cell_address(
                                            used.row + row_idx,
                                            used.column + col_idx,
                                        ),
                                        "match_type": match_type,
                                        "value": value,
                                        "formula": formula,
                                    }
                                )
                                break
                        if len(matches) >= max_results:
                            return {
                                "path": workbook_path.name,
                                "backend": "mac_excel",
                                "query": query,
                                "matches": matches,
                                "truncated": True,
                            }
        return {
            "path": workbook_path.name,
            "backend": "mac_excel",
            "query": query,
            "matches": matches,
            "truncated": False,
        }

    def recalculate(self, workbook_path: Path) -> JsonDict:
        with self._open_workbook(workbook_path, recalculate=True):
            return {
                "path": workbook_path.name,
                "backend": "mac_excel",
                "recalculated": True,
            }

    def workbook_repair_check(self, workbook_path: Path) -> JsonDict:
        """Detect whether Excel requires repair mode to open *workbook_path*.

        The normal probe uses Excel's CorruptLoad=xlNormalLoad.  If that fails
        but CorruptLoad=xlRepairFile succeeds, Excel would have shown a repair
        dialogue in an interactive open.
        """
        normal_probe = self._probe_open(workbook_path, corrupt_load=_XL_NORMAL_LOAD)
        repair_probe: JsonDict | None = None
        repair_dialog_detected = False
        opened = bool(normal_probe["opened"])

        if not opened:
            repair_probe = self._probe_open(workbook_path, corrupt_load=_XL_REPAIR_FILE)
            repair_dialog_detected = bool(repair_probe["opened"]) or error_mentions_repair(normal_probe.get("error"))
            opened = bool(repair_probe["opened"])

        details: JsonDict = {"normal_probe": normal_probe}
        if repair_probe is not None:
            details["repair_probe"] = repair_probe

        return {
            "path": workbook_path.name,
            "backend": "mac_excel",
            "checked": True,
            "opened": opened,
            "repair_dialog_detected": repair_dialog_detected,
            "details": details,
        }

    def _probe_open(self, workbook_path: Path, *, corrupt_load: int) -> JsonDict:
        try:
            xw = importlib.import_module("xlwings")
        except ImportError as e:
            msg = "xlwings is required for the local Microsoft Excel backend"
            raise RuntimeError(msg) from e

        app = xw.App(visible=self.visible, add_book=False)
        workbook = None
        try:
            app.display_alerts = False
            app.screen_updating = False
            workbook = open_excel_workbook(
                app,
                workbook_path,
                corrupt_load=corrupt_load,
            )
            return {
                "opened": True,
                "corrupt_load": corrupt_load,
                "sheet_count": len(workbook.sheets),
            }
        except Exception as e:  # noqa: BLE001
            return {
                "opened": False,
                "corrupt_load": corrupt_load,
                "error": str(e),
                "error_type": e.__class__.__name__,
            }
        finally:
            if workbook is not None:
                workbook.close()
            app.quit()


@contextmanager
def libreoffice_cache_lock(lock_dir: Path, *, timeout_seconds: int) -> Iterator[None]:
    """Acquire a simple cross-process directory lock for a cache entry."""
    lock_dir.parent.mkdir(parents=True, exist_ok=True)
    deadline = time_module.monotonic() + max(timeout_seconds, 1)
    stale_after = max(timeout_seconds * 2, 60)
    while True:
        try:
            lock_dir.mkdir()
            with suppress(OSError):
                os.chmod(lock_dir, 0o777)
            try:
                yield
            finally:
                shutil.rmtree(lock_dir, ignore_errors=True)
            return
        except FileExistsError:
            with suppress(OSError):
                if time_module.time() - lock_dir.stat().st_mtime > stale_after:
                    shutil.rmtree(lock_dir, ignore_errors=True)
                    continue
            if time_module.monotonic() >= deadline:
                msg = f"Timed out waiting for LibreOffice cache lock {lock_dir}"
                raise TimeoutError(msg)
            time_module.sleep(_LIBREOFFICE_LOCK_POLL_SECONDS)


def file_sha256(path: Path) -> str:
    """Return the SHA-256 digest of a local file."""
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class LibreOfficeBackend:
    """Headless LibreOffice backend using openpyxl for structured reads."""

    def __init__(
        self,
        *,
        soffice_path: str | None = None,
        timeout_seconds: int = DEFAULT_EXCEL_TIMEOUT_SECONDS,
        cache_dir: Path | None = None,
        loaded_workbook_cache_size: int = _LIBREOFFICE_LOADED_WORKBOOK_CACHE_SIZE,
    ) -> None:
        self.soffice_path = soffice_path
        self.timeout_seconds = require_positive_int(timeout_seconds, "timeout_seconds")
        self.loaded_workbook_cache_size = require_positive_int(
            loaded_workbook_cache_size,
            "loaded_workbook_cache_size",
        )
        if cache_dir is None:
            self._cache_root = Path(tempfile.mkdtemp(prefix="gandalf_libreoffice_cache_"))
            atexit.register(shutil.rmtree, self._cache_root, ignore_errors=True)
        else:
            self._cache_root = cache_dir
            self._cache_root.mkdir(parents=True, exist_ok=True)
        with suppress(OSError):
            os.chmod(self._cache_root, 0o777)
        self._converted_cache: dict[_LibreOfficeCacheKey, Path] = {}
        self._hash_cache: dict[_WorkbookStatKey, str] = {}
        self._loaded_workbook_cache: OrderedDict[_WorkbookStatKey, _LibreOfficeOpenWorkbookPair] = OrderedDict()
        self._cache_lock = threading.Lock()
        atexit.register(self.close)

    def workbook_summary(self, workbook_path: Path, *, recalculate: bool) -> JsonDict:
        with self._open_workbooks(workbook_path, recalculate=recalculate) as (
            formula_workbook,
            _value_workbook,
            prepared,
        ):
            sheets: list[JsonDict] = []
            for worksheet in formula_workbook.worksheets:
                used_range = str(worksheet.calculate_dimension())
                rows, columns = range_shape(used_range)
                sheets.append(
                    {
                        "name": worksheet.title,
                        "used_range": used_range,
                        "rows": rows,
                        "columns": columns,
                        "visible": worksheet.sheet_state == "visible",
                    }
                )
            result = {
                "path": workbook_path.name,
                "backend": "libreoffice",
                "converted": prepared.converted,
                "recalculated": prepared.recalculated,
                "sheets": sheets,
            }
            add_result_metrics(result, prepared.metrics)
            return result

    def inspect_range(
        self,
        workbook_path: Path,
        *,
        sheet: str,
        range_address: str,
        include_values: bool,
        include_formulas: bool,
        include_formats: bool,
        recalculate: bool,
    ) -> JsonDict:
        rows, columns = range_shape(range_address)
        start_row, start_column, _end_row, _end_column = range_bounds(range_address)
        with self._open_workbooks(workbook_path, recalculate=recalculate) as (
            formula_workbook,
            value_workbook,
            prepared,
        ):
            if sheet not in formula_workbook.sheetnames:
                msg = f"Worksheet {sheet!r} does not exist"
                raise ValueError(msg)
            formula_sheet = formula_workbook[sheet]
            value_sheet = value_workbook[sheet]
            result: JsonDict = {
                "path": workbook_path.name,
                "backend": "libreoffice",
                "sheet": sheet,
                "range": range_address,
                "rows": rows,
                "columns": columns,
                "converted": prepared.converted,
                "recalculated": prepared.recalculated,
            }
            if include_values:
                result["values"] = [
                    [
                        jsonable(value_sheet.cell(row=start_row + row_idx, column=start_column + column_idx).value)
                        for column_idx in range(columns)
                    ]
                    for row_idx in range(rows)
                ]
            if include_formulas:
                result["formulas"] = [
                    [
                        jsonable(formula_sheet.cell(row=start_row + row_idx, column=start_column + column_idx).value)
                        for column_idx in range(columns)
                    ]
                    for row_idx in range(rows)
                ]
            if include_formats:
                result["number_formats"] = [
                    [
                        jsonable(
                            formula_sheet.cell(
                                row=start_row + row_idx,
                                column=start_column + column_idx,
                            ).number_format
                        )
                        for column_idx in range(columns)
                    ]
                    for row_idx in range(rows)
                ]
            add_result_metrics(result, prepared.metrics)
            return result

    def find_cells(
        self,
        workbook_path: Path,
        *,
        query: str,
        search_values: bool,
        search_formulas: bool,
        max_results: int,
        max_cells_per_call: int,
    ) -> JsonDict:
        query_lc = query.lower()
        matches: list[JsonDict] = []
        inspected_cells = 0
        with self._open_workbooks(workbook_path, recalculate=True) as (
            formula_workbook,
            value_workbook,
            prepared,
        ):
            for formula_sheet in formula_workbook.worksheets:
                used_range = str(formula_sheet.calculate_dimension())
                rows, columns = range_shape(used_range)
                inspected_cells += rows * columns
                if inspected_cells > max_cells_per_call:
                    msg = f"Workbook search exceeds max_cells_per_call={max_cells_per_call}"
                    raise ValueError(msg)
                start_row, start_column, _end_row, _end_column = range_bounds(used_range)
                value_sheet = value_workbook[formula_sheet.title]
                for row_idx in range(rows):
                    for col_idx in range(columns):
                        row = start_row + row_idx
                        column = start_column + col_idx
                        value = jsonable(value_sheet.cell(row=row, column=column).value)
                        formula = jsonable(formula_sheet.cell(row=row, column=column).value)
                        haystacks: list[tuple[str, Any]] = []
                        if search_values:
                            haystacks.append(("value", value))
                        if search_formulas:
                            haystacks.append(("formula", formula))
                        for match_type, haystack in haystacks:
                            if haystack is not None and query_lc in str(haystack).lower():
                                matches.append(
                                    {
                                        "sheet": formula_sheet.title,
                                        "cell": cell_address(row, column),
                                        "match_type": match_type,
                                        "value": value,
                                        "formula": formula,
                                    }
                                )
                                break
                        if len(matches) >= max_results:
                            result = {
                                "path": workbook_path.name,
                                "backend": "libreoffice",
                                "converted": prepared.converted,
                                "recalculated": prepared.recalculated,
                                "query": query,
                                "matches": matches,
                                "truncated": True,
                            }
                            add_result_metrics(result, prepared.metrics)
                            return result
        result = {
            "path": workbook_path.name,
            "backend": "libreoffice",
            "converted": prepared.converted,
            "recalculated": prepared.recalculated,
            "query": query,
            "matches": matches,
            "truncated": False,
        }
        add_result_metrics(result, prepared.metrics)
        return result

    def recalculate(self, workbook_path: Path) -> JsonDict:
        with self._prepared_workbook_path(workbook_path, recalculate=True) as prepared:
            result = {
                "path": workbook_path.name,
                "backend": "libreoffice",
                "converted": prepared.converted,
                "recalculated": prepared.recalculated,
            }
            add_result_metrics(result, prepared.metrics)
            return result

    def workbook_repair_check(self, workbook_path: Path) -> JsonDict:
        return {
            "path": workbook_path.name,
            "backend": "libreoffice",
            "checked": False,
            "opened": None,
            "repair_dialog_detected": None,
            "skipped_reason": "repair dialogue detection requires Microsoft Excel",
        }

    @contextmanager
    def _open_workbooks(
        self,
        workbook_path: Path,
        *,
        recalculate: bool,
    ) -> Iterator[tuple[Any, Any, _LibreOfficePreparedWorkbook]]:
        with self._prepared_workbook_path(workbook_path, recalculate=recalculate) as prepared:
            workbook_pair, metrics = self._openpyxl_workbook_pair(prepared.path)
            merged_metrics = dict(prepared.metrics)
            merged_metrics.update(metrics)
            yield (
                workbook_pair.formula_workbook,
                workbook_pair.value_workbook,
                _LibreOfficePreparedWorkbook(
                    path=prepared.path,
                    converted=prepared.converted,
                    recalculated=prepared.recalculated,
                    metrics=merged_metrics,
                ),
            )

    @contextmanager
    def _prepared_workbook_path(self, workbook_path: Path, *, recalculate: bool) -> Iterator[_LibreOfficePreparedWorkbook]:
        if recalculate or workbook_path.suffix.lower() not in _OPENPYXL_EXTENSIONS:
            prepared_path, metrics = self._cached_converted_workbook(
                workbook_path,
                force_recalculate=recalculate,
            )
            yield _LibreOfficePreparedWorkbook(
                path=prepared_path,
                converted=True,
                recalculated=recalculate,
                metrics=metrics,
            )
        else:
            yield _LibreOfficePreparedWorkbook(
                path=workbook_path,
                converted=False,
                recalculated=False,
                metrics={},
            )

    def _cached_converted_workbook(self, workbook_path: Path, *, force_recalculate: bool) -> tuple[Path, JsonDict]:
        key, metrics = self._cache_key(workbook_path, force_recalculate=force_recalculate)
        with self._cache_lock:
            cached = self._converted_cache.get(key)
            if cached is not None and cached.is_file():
                metrics["libreoffice_prepared_cache_hit"] = True
                return cached, metrics

        cache_dir = self._prepared_cache_dir(key)
        prepared_path = cache_dir / _LIBREOFFICE_PREPARED_WORKBOOK_NAME
        if prepared_path.is_file():
            with self._cache_lock:
                self._converted_cache[key] = prepared_path
            metrics["libreoffice_prepared_cache_hit"] = True
            return prepared_path, metrics

        lock_started = time_module.perf_counter()
        with libreoffice_cache_lock(cache_dir.with_name(f"{cache_dir.name}.lock"), timeout_seconds=self.timeout_seconds):
            metrics["libreoffice_cache_lock_wait_ms"] = round((time_module.perf_counter() - lock_started) * 1000, 3)
            if prepared_path.is_file():
                with self._cache_lock:
                    self._converted_cache[key] = prepared_path
                metrics["libreoffice_prepared_cache_hit"] = True
                return prepared_path, metrics

            metrics["libreoffice_prepared_cache_hit"] = False
            temp_dir = cache_dir.with_name(f"{cache_dir.name}.tmp.{secrets.token_hex(8)}")
            temp_dir.mkdir(parents=True)
            with suppress(OSError):
                os.chmod(temp_dir, 0o777)
            try:
                started = time_module.perf_counter()
                converted = self._convert_with_libreoffice(
                    workbook_path,
                    temp_dir,
                    force_recalculate=force_recalculate,
                )
                metrics["libreoffice_conversion_ms"] = round((time_module.perf_counter() - started) * 1000, 3)
                cache_dir.mkdir(parents=True, exist_ok=True)
                with suppress(OSError):
                    os.chmod(cache_dir, 0o777)
                os.replace(converted, prepared_path)
                with suppress(OSError):
                    os.chmod(prepared_path, 0o666)
            except BaseException:
                shutil.rmtree(cache_dir, ignore_errors=True)
                raise
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)
            with self._cache_lock:
                self._converted_cache[key] = prepared_path
            return prepared_path, metrics

    def _openpyxl_workbook_pair(self, prepared_path: Path) -> tuple[_LibreOfficeOpenWorkbookPair, JsonDict]:
        key = workbook_stat_key(prepared_path)
        with self._cache_lock:
            # FastMCP stdio currently serialises tool calls for one MCP server.
            # If that ever changes, cached openpyxl workbooks need a lease or
            # read lock so eviction cannot close a workbook while another call
            # is still reading it.
            cached = self._loaded_workbook_cache.get(key)
            if cached is not None:
                self._loaded_workbook_cache.move_to_end(key)
                return cached, {"libreoffice_loaded_workbook_cache_hit": True}

            started = time_module.perf_counter()
            formula_workbook = load_openpyxl_workbook(prepared_path, data_only=False)
            try:
                value_workbook = load_openpyxl_workbook(prepared_path, data_only=True)
            except BaseException:
                close_workbook(formula_workbook)
                raise
            pair = _LibreOfficeOpenWorkbookPair(
                formula_workbook=formula_workbook,
                value_workbook=value_workbook,
            )
            self._loaded_workbook_cache[key] = pair
            self._evict_loaded_workbook_cache()
            return pair, {
                "libreoffice_loaded_workbook_cache_hit": False,
                "libreoffice_openpyxl_load_ms": round((time_module.perf_counter() - started) * 1000, 3),
            }

    def _evict_loaded_workbook_cache(self) -> None:
        while len(self._loaded_workbook_cache) > self.loaded_workbook_cache_size:
            _key, pair = self._loaded_workbook_cache.popitem(last=False)
            close_workbook(pair.formula_workbook)
            close_workbook(pair.value_workbook)

    def _cache_key(self, workbook_path: Path, *, force_recalculate: bool) -> tuple[_LibreOfficeCacheKey, JsonDict]:
        started = time_module.perf_counter()
        stat_key = workbook_stat_key(workbook_path)
        with self._cache_lock:
            digest = self._hash_cache.get(stat_key)
        if digest is None:
            digest = file_sha256(workbook_path)
            with self._cache_lock:
                self._hash_cache[stat_key] = digest
        metrics = {"libreoffice_hash_ms": round((time_module.perf_counter() - started) * 1000, 3)}
        return _LibreOfficeCacheKey(
            content_sha256=digest,
            suffix=workbook_path.suffix.lower(),
            force_recalculate=force_recalculate,
        ), metrics

    def _prepared_cache_dir(self, key: _LibreOfficeCacheKey) -> Path:
        return self._cache_root / hashlib.sha256(repr(key).encode("utf-8")).hexdigest()

    def close(self) -> None:
        """Close cached openpyxl workbooks held by this backend instance."""
        with self._cache_lock:
            pairs = list(self._loaded_workbook_cache.values())
            self._loaded_workbook_cache.clear()
        for pair in pairs:
            close_workbook(pair.formula_workbook)
            close_workbook(pair.value_workbook)

    def _convert_with_libreoffice(self, workbook_path: Path, temp_dir: Path, *, force_recalculate: bool) -> Path:
        soffice_path = self._resolve_soffice_path()
        source_dir = temp_dir / "source"
        output_dir = temp_dir / "output"
        profile_dir = temp_dir / "profile"
        source_dir.mkdir()
        output_dir.mkdir()
        profile_dir.mkdir()
        if force_recalculate:
            write_libreoffice_recalc_profile(profile_dir)
        source_path = source_dir / workbook_path.name
        shutil.copy2(workbook_path, source_path)
        command = [
            soffice_path,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_dir),
            str(source_path),
        ]
        try:
            completed = subprocess.run(  # noqa: S603
                command,
                check=False,
                capture_output=True,
                text=True,
                timeout=self.timeout_seconds,
            )
        except subprocess.TimeoutExpired as e:
            msg = f"LibreOffice conversion timed out after {self.timeout_seconds}s"
            raise RuntimeError(msg) from e
        if completed.returncode != 0:
            output = "\n".join(part for part in (completed.stdout.strip(), completed.stderr.strip()) if part)
            msg = f"LibreOffice conversion failed with exit code {completed.returncode}: {output}"
            raise RuntimeError(msg)
        expected_path = output_dir / f"{workbook_path.stem}.xlsx"
        if expected_path.is_file():
            return expected_path
        candidates = sorted(output_dir.glob("*.xlsx"))
        if len(candidates) == 1:
            return candidates[0]
        output = "\n".join(part for part in (completed.stdout.strip(), completed.stderr.strip()) if part)
        msg = f"LibreOffice conversion did not produce an .xlsx file: {output}"
        raise RuntimeError(msg)

    def _resolve_soffice_path(self) -> str:
        if self.soffice_path:
            return self.soffice_path
        soffice_path = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice_path is None:
            msg = "LibreOffice backend requires 'soffice' or 'libreoffice' on PATH"
            raise RuntimeError(msg)
        return soffice_path


class WindowsVMExcelBackend:
    """Proxy backend for a remote Windows VM Excel worker."""

    def __init__(self, *, url: str, timeout_seconds: int, auth_token: str | None = None) -> None:
        self.url = url.rstrip("/")
        self.timeout_seconds = timeout_seconds
        self.auth_token = auth_token
        self._remote_workbook_lock = threading.Lock()
        self._remote_workbooks: dict[_WorkbookStatKey, _RemoteWorkbookRef] = {}

    def workbook_summary(self, workbook_path: Path, *, recalculate: bool) -> JsonDict:
        return self._post(workbook_path, "workbook_summary", {"recalculate": recalculate})

    def inspect_range(
        self,
        workbook_path: Path,
        *,
        sheet: str,
        range_address: str,
        include_values: bool,
        include_formulas: bool,
        include_formats: bool,
        recalculate: bool,
    ) -> JsonDict:
        return self._post(
            workbook_path,
            "inspect_range",
            {
                "sheet": sheet,
                "range_address": range_address,
                "include_values": include_values,
                "include_formulas": include_formulas,
                "include_formats": include_formats,
                "recalculate": recalculate,
            },
        )

    def find_cells(
        self,
        workbook_path: Path,
        *,
        query: str,
        search_values: bool,
        search_formulas: bool,
        max_results: int,
        max_cells_per_call: int,
    ) -> JsonDict:
        return self._post(
            workbook_path,
            "find_cells",
            {
                "query": query,
                "search_values": search_values,
                "search_formulas": search_formulas,
                "max_results": max_results,
                "max_cells_per_call": max_cells_per_call,
            },
        )

    def recalculate(self, workbook_path: Path) -> JsonDict:
        return self._post(workbook_path, "recalculate", {})

    def workbook_repair_check(self, workbook_path: Path) -> JsonDict:
        return self._post(workbook_path, "workbook_repair_check", {})

    def _post(self, workbook_path: Path, operation: str, args: JsonDict) -> JsonDict:
        started = time_module.perf_counter()
        remote_ref, ensure_metrics = self._ensure_workbook(workbook_path)
        operation_workbook_bytes_sent = 0
        operation_ms = 0.0
        payload_bytes = 0
        retried_after_cache_miss = False
        while True:
            payload = {
                "operation": operation,
                "args": args,
                "workbook_name": remote_ref.workbook_name,
                "workbook_sha256": remote_ref.sha256,
                "workbook_id": remote_ref.sha256,
            }
            if remote_ref.legacy_full_upload:
                workbook_bytes = workbook_path.read_bytes()
                payload["workbook_b64"] = base64.b64encode(workbook_bytes).decode("ascii")
                operation_workbook_bytes_sent += len(workbook_bytes)
            data, request_ms, request_payload_bytes = self._send_json("/inspect", payload)
            operation_ms += request_ms
            payload_bytes += request_payload_bytes
            if data.get("ok") is True:
                break
            error = str(data.get("error", "unknown error"))
            if retried_after_cache_miss or remote_ref.legacy_full_upload or "workbook_not_cached" not in error:
                msg = f"Windows Excel worker error: {data.get('error', 'unknown error')}"
                raise RuntimeError(msg)
            self._forget_workbook(workbook_path)
            remote_ref, ensure_metrics = self._ensure_workbook(
                workbook_path,
                previous_metrics=ensure_metrics,
            )
            retried_after_cache_miss = True
        result = data.get("result")
        if not isinstance(result, dict):
            msg = "Windows Excel worker response missing object result"
            raise RuntimeError(msg)
        result["backend"] = "windows_vm"
        ensure_metrics = dict(ensure_metrics)
        ensure_workbook_bytes_sent = ensure_metrics.pop("windows_vm_client_workbook_bytes_sent", 0)
        if not isinstance(ensure_workbook_bytes_sent, int):
            ensure_workbook_bytes_sent = 0
        add_result_metrics(
            result,
            {
                "windows_vm_client_operation": operation,
                "windows_vm_client_request_ms": round(operation_ms, 3),
                "windows_vm_client_total_ms": round((time_module.perf_counter() - started) * 1000, 3),
                "windows_vm_client_payload_bytes": payload_bytes,
                "windows_vm_client_workbook_bytes_sent": (
                    ensure_workbook_bytes_sent + operation_workbook_bytes_sent
                ),
                "windows_vm_client_legacy_full_upload": remote_ref.legacy_full_upload,
                "windows_vm_client_retried_after_cache_miss": retried_after_cache_miss,
                **ensure_metrics,
            },
        )
        return result

    def _ensure_workbook(
        self,
        workbook_path: Path,
        *,
        previous_metrics: JsonDict | None = None,
    ) -> tuple[_RemoteWorkbookRef, JsonDict]:
        local_key = workbook_stat_key(workbook_path)
        with self._remote_workbook_lock:
            cached = self._remote_workbooks.get(local_key)
        if cached is not None:
            return cached, merge_windows_vm_ensure_metrics(
                previous_metrics,
                {
                    "windows_vm_client_workbook_cache_hit": True,
                    "windows_vm_client_workbook_bytes_sent": 0,
                    "windows_vm_worker_workbook_cache_hit": None,
                    "windows_vm_client_ensure_ms": 0.0,
                },
            )

        workbook_bytes = workbook_path.read_bytes()
        workbook_sha256 = hashlib.sha256(workbook_bytes).hexdigest()
        ref = _RemoteWorkbookRef(workbook_name=workbook_path.name, sha256=workbook_sha256)
        metadata_payload = {
            "workbook_name": ref.workbook_name,
            "workbook_sha256": ref.sha256,
            "workbook_id": ref.sha256,
        }
        try:
            data, ensure_ms, _payload_bytes = self._send_json("/workbook/ensure", metadata_payload)
        except RuntimeError as e:
            if "HTTP Error 404" not in str(e):
                raise
            legacy_ref = _RemoteWorkbookRef(
                workbook_name=workbook_path.name,
                sha256=workbook_sha256,
                legacy_full_upload=True,
            )
            with self._remote_workbook_lock:
                self._remote_workbooks[local_key] = legacy_ref
            return legacy_ref, merge_windows_vm_ensure_metrics(
                previous_metrics,
                {
                    "windows_vm_client_workbook_cache_hit": False,
                    "windows_vm_client_workbook_bytes_sent": 0,
                    "windows_vm_worker_workbook_cache_hit": None,
                    "windows_vm_client_ensure_ms": 0.0,
                },
            )
        bytes_sent = 0
        worker_cache_hit: bool | None = None
        if data.get("ok") is True:
            result = data.get("result")
            if not isinstance(result, dict):
                msg = "Windows Excel worker response missing object result"
                raise RuntimeError(msg)
            worker_cache_hit = bool(result.get("cache_hit", False))
        elif "workbook_not_cached" in str(data.get("error", "")):
            upload_payload = {
                **metadata_payload,
                "workbook_b64": base64.b64encode(workbook_bytes).decode("ascii"),
            }
            data, upload_ms, _payload_bytes = self._send_json("/workbook/ensure", upload_payload)
            ensure_ms += upload_ms
            bytes_sent = len(workbook_bytes)
            if data.get("ok") is not True:
                msg = f"Windows Excel worker error: {data.get('error', 'unknown error')}"
                raise RuntimeError(msg)
            result = data.get("result")
            if not isinstance(result, dict):
                msg = "Windows Excel worker response missing object result"
                raise RuntimeError(msg)
            worker_cache_hit = bool(result.get("cache_hit", False))
        else:
            msg = f"Windows Excel worker error: {data.get('error', 'unknown error')}"
            raise RuntimeError(msg)

        with self._remote_workbook_lock:
            self._remote_workbooks[local_key] = ref
        return ref, merge_windows_vm_ensure_metrics(
            previous_metrics,
            {
                "windows_vm_client_workbook_cache_hit": False,
                "windows_vm_client_workbook_bytes_sent": bytes_sent,
                "windows_vm_worker_workbook_cache_hit": worker_cache_hit,
                "windows_vm_client_ensure_ms": round(ensure_ms, 3),
            },
        )

    def _forget_workbook(self, workbook_path: Path) -> None:
        local_key = workbook_stat_key(workbook_path)
        with self._remote_workbook_lock:
            self._remote_workbooks.pop(local_key, None)

    def _send_json(self, endpoint: str, payload: JsonDict) -> tuple[JsonDict, float, int]:
        body = json.dumps(payload).encode("utf-8")
        headers = {"Content-Type": "application/json"}
        if self.auth_token is not None:
            headers[EXCEL_AUTH_HEADER] = self.auth_token
        request = urllib.request.Request(
            f"{self.url}{endpoint}",
            data=body,
            headers=headers,
            method="POST",
        )
        started = time_module.perf_counter()
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                data = json.loads(response.read().decode("utf-8"))
        except urllib.error.URLError as e:
            msg = f"Windows Excel worker request failed: {e}"
            raise RuntimeError(msg) from e
        elapsed_ms = (time_module.perf_counter() - started) * 1000
        if not isinstance(data, dict):
            msg = "Windows Excel worker returned a non-object response"
            raise RuntimeError(msg)
        return data, elapsed_ms, len(body)


def create_service(config: ExcelServerConfig) -> ExcelService:
    """Create an ExcelService from runtime config."""
    if config.backend == "windows_vm":
        if config.windows_url is None:
            msg = "GANDALF_EXCEL_WINDOWS_URL is required for windows_vm backend"
            raise RuntimeError(msg)
        backend: ExcelBackend = WindowsVMExcelBackend(
            url=config.windows_url,
            timeout_seconds=config.timeout_seconds,
            auth_token=config.auth_token,
        )
    elif config.backend == "mac_excel":
        backend = LocalExcelBackend(visible=config.visible)
    elif config.backend == "libreoffice":
        backend = LibreOfficeBackend(timeout_seconds=config.timeout_seconds, cache_dir=config.cache_dir)
    else:
        msg = f"Unsupported Excel backend {config.backend!r}"
        raise RuntimeError(msg)
    return ExcelService(
        workdir=config.workdir,
        backend=backend,
        allow_macros=config.allow_macros,
        max_cells_per_call=config.max_cells_per_call,
        max_format_cells_per_call=config.max_format_cells_per_call,
        metrics_path=config.metrics_path,
        debug_metrics=config.debug_metrics,
    )


def create_mcp(service: ExcelService) -> FastMCP:
    """Build the FastMCP server for Excel workbook inspection."""
    mcp = FastMCP("gandalf-excel")

    @mcp.tool
    def excel_workbook_summary(path: str, recalculate: bool = True) -> JsonDict:
        """Summarise workbook sheets and used ranges using the configured workbook backend."""
        return service.workbook_summary(path, recalculate=recalculate)

    @mcp.tool
    def excel_inspect_range(
        path: str,
        sheet: str,
        range: str,  # noqa: A002 - user-facing Excel term
        include_values: bool = True,
        include_formulas: bool = True,
        include_formats: bool = False,
    ) -> JsonDict:
        """Inspect a worksheet range; pass include_formats=true for number formats or cell formatting."""
        return service.inspect_range(
            path,
            sheet=sheet,
            range_address=range,
            include_values=include_values,
            include_formulas=include_formulas,
            include_formats=include_formats,
        )

    @mcp.tool
    def excel_find_cells(
        path: str,
        query: str,
        search_values: bool = True,
        search_formulas: bool = True,
        max_results: int = 100,
    ) -> JsonDict:
        """Search workbook values and formulas using the configured workbook backend."""
        return service.find_cells(
            path,
            query=query,
            search_values=search_values,
            search_formulas=search_formulas,
            max_results=max_results,
        )

    @mcp.tool
    def excel_recalculate(path: str) -> JsonDict:
        """Open and recalculate a workbook using the configured workbook backend."""
        return service.recalculate(path)

    @mcp.tool
    def excel_workbook_repair_check(path: str) -> JsonDict:
        """Detect whether Microsoft Excel must repair a workbook before opening it, when supported."""
        return service.workbook_repair_check(path)

    return mcp


def config_from_env() -> ExcelServerConfig:
    """Build MCP server config from judge-process environment variables."""
    workdir = os.environ.get("GANDALF_EXCEL_WORKDIR")
    if not workdir:
        msg = "GANDALF_EXCEL_WORKDIR is required"
        raise RuntimeError(msg)
    return ExcelServerConfig(
        workdir=Path(workdir),
        backend=os.environ.get("GANDALF_EXCEL_BACKEND", DEFAULT_EXCEL_BACKEND),
        windows_url=os.environ.get("GANDALF_EXCEL_WINDOWS_URL"),
        auth_token=os.environ.get("GANDALF_EXCEL_AUTH_TOKEN"),
        timeout_seconds=int(os.environ.get("GANDALF_EXCEL_TIMEOUT_SECONDS", str(DEFAULT_EXCEL_TIMEOUT_SECONDS))),
        visible=parse_bool(os.environ.get("GANDALF_EXCEL_VISIBLE", "false")),
        allow_macros=parse_bool(os.environ.get("GANDALF_EXCEL_ALLOW_MACROS", "false")),
        max_cells_per_call=parse_positive_env_int(
            "GANDALF_EXCEL_MAX_CELLS_PER_CALL",
            DEFAULT_EXCEL_MAX_CELLS_PER_CALL,
        ),
        max_format_cells_per_call=parse_positive_env_int(
            "GANDALF_EXCEL_MAX_FORMAT_CELLS_PER_CALL",
            DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
        ),
        metrics_path=(
            Path(metrics_path)
            if (metrics_path := os.environ.get("GANDALF_EXCEL_METRICS_PATH"))
            else None
        ),
        debug_metrics=parse_bool(os.environ.get("GANDALF_EXCEL_DEBUG_METRICS", "false")),
        cache_dir=(
            Path(cache_dir)
            if (cache_dir := os.environ.get("GANDALF_EXCEL_CACHE_DIR"))
            else None
        ),
    )


def parse_bool(value: str) -> bool:
    """Parse a simple boolean environment value."""
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def parse_positive_env_int(name: str, default: int) -> int:
    """Parse a positive integer from an environment variable."""
    raw = os.environ.get(name, str(default))
    try:
        value = int(raw)
    except ValueError as e:
        msg = f"{name} must be a positive integer"
        raise RuntimeError(msg) from e
    return require_positive_int(value, name)


def positive_int(value: str) -> int:
    """Argparse type for positive integers."""
    try:
        parsed = int(value)
    except ValueError as e:
        msg = "must be a positive integer"
        raise argparse.ArgumentTypeError(msg) from e
    try:
        return require_positive_int(parsed, "value")
    except ValueError as e:
        raise argparse.ArgumentTypeError(str(e)) from e


def require_positive_int(value: int, name: str) -> int:
    """Return *value* when positive, otherwise raise a clear validation error."""
    if value <= 0:
        msg = f"{name} must be positive"
        raise ValueError(msg)
    return value


def open_excel_workbook(app: Any, workbook_path: Path, *, corrupt_load: int) -> Any:
    """Open a workbook with a requested Excel CorruptLoad mode."""
    kwargs: JsonDict = {
        "update_links": False,
        "read_only": True,
        "ignore_read_only_recommended": True,
        "corrupt_load": corrupt_load,
    }
    try:
        return app.books.open(str(workbook_path), **kwargs)
    except Exception as e:
        # Older xlwings versions and xlwings on macOS may not expose
        # CorruptLoad. Fall back to the normal open path only for xlNormalLoad;
        # repair-mode probing is unavailable in those environments.
        if corrupt_load != _XL_NORMAL_LOAD or not corrupt_load_is_unsupported(e):
            raise
        kwargs.pop("corrupt_load", None)
        return app.books.open(str(workbook_path), **kwargs)


def error_mentions_repair(error: Any) -> bool:
    """Return whether an Excel/open error explicitly looks repair-related."""
    if error is None:
        return False
    error_text = str(error).lower()
    if _CORRUPT_LOAD_UNSUPPORTED_MARKER in error_text:
        return False
    return any(marker in error_text for marker in _REPAIR_ERROR_MARKERS)


def corrupt_load_is_unsupported(error: Any) -> bool:
    """Return whether the host xlwings/Excel layer lacks CorruptLoad support."""
    return _CORRUPT_LOAD_UNSUPPORTED_MARKER in str(error).lower()


def column_to_index(column: str) -> int:
    """Convert an Excel column label to a 1-based index."""
    total = 0
    for char in column.upper():
        if not ("A" <= char <= "Z"):
            msg = f"Invalid Excel column {column!r}"
            raise ValueError(msg)
        total = total * 26 + (ord(char) - ord("A") + 1)
    if total > EXCEL_MAX_COLUMN_INDEX:
        msg = f"Excel column {column!r} exceeds maximum Excel column {EXCEL_MAX_COLUMN_LABEL}"
        raise ValueError(msg)
    return total


def index_to_column(index: int) -> str:
    """Convert a 1-based column index to an Excel column label."""
    if index < 1:
        msg = "Column index must be positive"
        raise ValueError(msg)
    if index > EXCEL_MAX_COLUMN_INDEX:
        msg = f"Column index {index} exceeds maximum Excel column {EXCEL_MAX_COLUMN_LABEL}"
        raise ValueError(msg)
    label = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        label = chr(ord("A") + remainder) + label
    return label


def cell_address(row: int, column: int) -> str:
    """Return an A1 address for 1-based row and column indices."""
    return f"{index_to_column(column)}{row}"


def range_shape(range_address: str) -> tuple[int, int]:
    """Return row and column counts for an A1 range."""
    start_row, start_col_idx, end_row, end_col_idx = range_bounds(range_address)
    return end_row - start_row + 1, end_col_idx - start_col_idx + 1


def range_bounds(range_address: str) -> tuple[int, int, int, int]:
    """Return 1-based start/end row and column indexes for an A1 range."""
    match = _RANGE_RE.match(range_address)
    if not match:
        msg = f"Invalid Excel range address {range_address!r}"
        raise ValueError(msg)
    start_col, start_row_raw, end_col, end_row_raw = match.groups()
    start_row = int(start_row_raw)
    start_col_idx = column_to_index(start_col)
    end_row = int(end_row_raw) if end_row_raw is not None else start_row
    end_col_idx = column_to_index(end_col) if end_col is not None else start_col_idx
    if end_row < start_row or end_col_idx < start_col_idx:
        msg = f"Invalid reversed Excel range address {range_address!r}"
        raise ValueError(msg)
    return start_row, start_col_idx, end_row, end_col_idx


def range_cell_count(range_address: str) -> int:
    """Return the number of cells in an A1 range."""
    rows, columns = range_shape(range_address)
    return rows * columns


def to_2d(value: Any, *, rows: int, columns: int) -> list[list[Any]]:
    """Normalise scalar/list Excel values to a 2D list."""
    if rows == 1 and columns == 1:
        return [[jsonable(value)]]
    if isinstance(value, list | tuple):
        if rows == 1:
            return [[jsonable(item) for item in value]]
        if value and isinstance(value[0], list | tuple):
            return [[jsonable(item) for item in row] for row in value]
        return [[jsonable(item)] for item in value]
    return [[jsonable(value) for _column in range(columns)] for _row in range(rows)]


def range_number_formats(worksheet: Any, cell_range: Any, *, rows: int, columns: int) -> list[list[Any]]:
    """Return number formats as a true per-cell 2D grid.

    xlwings does not expose a reliable bulk per-cell format API across Excel
    hosts, so callers must enforce a conservative cell budget before calling.
    """
    start_row = cell_range.row
    start_column = cell_range.column
    return [
        [
            jsonable(worksheet.range((start_row + row_idx, start_column + column_idx)).number_format)
            for column_idx in range(columns)
        ]
        for row_idx in range(rows)
    ]


def jsonable(value: Any) -> Any:
    """Convert common Excel/Python values to JSON-safe values."""
    if isinstance(value, datetime | date | time):
        return value.isoformat()
    if isinstance(value, tuple):
        return [jsonable(item) for item in value]
    return value


def load_openpyxl_workbook(path: Path, *, data_only: bool) -> Any:
    """Load a workbook with openpyxl, raising a clear backend error if missing."""
    try:
        openpyxl = importlib.import_module("openpyxl")
    except ImportError as e:
        msg = "openpyxl is required for the LibreOffice workbook backend"
        raise RuntimeError(msg) from e
    return openpyxl.load_workbook(path, data_only=data_only)


def write_libreoffice_recalc_profile(profile_dir: Path) -> None:
    """Seed a LibreOffice user profile that always recalculates imported formulas."""
    user_dir = profile_dir / "user"
    user_dir.mkdir(parents=True, exist_ok=True)
    (user_dir / "registrymodifications.xcu").write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
""",
        encoding="utf-8",
    )


def close_workbook(workbook: Any) -> None:
    """Close an openpyxl workbook when the installed version exposes close()."""
    close = getattr(workbook, "close", None)
    if callable(close):
        close()


def run_worker(
    *,
    host: str,
    port: int,
    temp_dir: Path,
    auth_token: str | None,
    visible: bool,
    allow_macros: bool,
    max_cells_per_call: int,
    max_format_cells_per_call: int = DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
    cache_ttl_seconds: int = DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS,
) -> None:
    """Run a minimal HTTP Excel worker for a Windows VM."""
    server = build_worker_server(
        host=host,
        port=port,
        temp_dir=temp_dir,
        auth_token=auth_token,
        visible=visible,
        allow_macros=allow_macros,
        max_cells_per_call=max_cells_per_call,
        max_format_cells_per_call=max_format_cells_per_call,
        cache_ttl_seconds=cache_ttl_seconds,
    )
    server.serve_forever()


def build_worker_server(
    *,
    host: str,
    port: int,
    temp_dir: Path,
    auth_token: str | None,
    visible: bool,
    allow_macros: bool,
    max_cells_per_call: int,
    max_format_cells_per_call: int = DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
    cache_ttl_seconds: int = DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS,
) -> ThreadingHTTPServer:
    """Build (but do not start) the worker HTTP server.

    Split out from :func:`run_worker` so the full HTTP round-trip can be tested
    on an ephemeral port. Pass ``port=0`` to let the OS choose a free port; read
    the chosen port from ``server.server_address[1]``.
    """
    require_positive_int(max_cells_per_call, "max_cells_per_call")
    require_positive_int(max_format_cells_per_call, "max_format_cells_per_call")
    require_positive_int(cache_ttl_seconds, "cache_ttl_seconds")
    temp_dir.mkdir(parents=True, exist_ok=True)

    class WorkerHandler(BaseHTTPRequestHandler):
        def do_POST(self) -> None:  # noqa: N802
            if self.path not in {"/inspect", "/workbook/ensure"}:
                self._send_json(404, {"ok": False, "error": "not found"})
                return
            if not worker_request_is_authorised(auth_token, self.headers.get(EXCEL_AUTH_HEADER)):
                self._send_json(401, {"ok": False, "error": "unauthorised"})
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                if self.path == "/workbook/ensure":
                    result = handle_worker_ensure(
                        payload,
                        temp_dir=temp_dir,
                        cache_ttl_seconds=cache_ttl_seconds,
                    )
                else:
                    result = handle_worker_request(
                        payload,
                        temp_dir=temp_dir,
                        visible=visible,
                        allow_macros=allow_macros,
                        max_cells_per_call=max_cells_per_call,
                        max_format_cells_per_call=max_format_cells_per_call,
                        cache_ttl_seconds=cache_ttl_seconds,
                    )
            except Exception as e:  # noqa: BLE001
                self._send_json(200, {"ok": False, "error": str(e)})
            else:
                self._send_json(200, {"ok": True, "result": result})

        def log_message(self, format: str, *args: object) -> None:  # noqa: A002
            return

        def _send_json(self, status: int, payload: JsonDict) -> None:
            body = json.dumps(payload).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    return ThreadingHTTPServer((host, port), WorkerHandler)


def worker_request_is_authorised(auth_token: str | None, supplied_token: str | None) -> bool:
    """Return whether an incoming worker request has the configured token."""
    if auth_token is None:
        return True
    if supplied_token is None:
        return False
    return secrets.compare_digest(auth_token.encode("utf-8"), supplied_token.encode("utf-8"))


@contextmanager
def worker_cache_lock(cache_key: str, *, blocking: bool = True) -> Iterator[bool]:
    """Lease the per-cache-key lock used to protect worker session dirs.

    The registry entry is removed when no active holder or waiter remains, so a
    long-lived worker does not retain one lock object per distinct workbook.
    """
    with _WORKER_CACHE_LOCKS_GUARD:
        state = _WORKER_CACHE_LOCKS.get(cache_key)
        if state is None:
            state = _WorkerCacheLockState(lock=threading.Lock())
            _WORKER_CACHE_LOCKS[cache_key] = state
        state.references += 1

    acquired = False
    try:
        acquired = state.lock.acquire(blocking=blocking)
        yield acquired
    finally:
        if acquired:
            state.lock.release()
        with _WORKER_CACHE_LOCKS_GUARD:
            state.references -= 1
            if state.references == 0 and _WORKER_CACHE_LOCKS.get(cache_key) is state:
                del _WORKER_CACHE_LOCKS[cache_key]


def prune_worker_cache(
    temp_dir: Path,
    *,
    ttl_seconds: int = DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS,
    now: float | None = None,
) -> None:
    """Remove stale content-addressed workbook cache directories."""
    if ttl_seconds <= 0:
        msg = "ttl_seconds must be positive"
        raise ValueError(msg)
    if not temp_dir.exists():
        return
    cutoff = (time_module.time() if now is None else now) - ttl_seconds
    for entry in temp_dir.iterdir():
        if not entry.is_dir():
            continue
        with worker_cache_lock(entry.name, blocking=False) as acquired:
            if not acquired:
                continue
            try:
                if entry.stat().st_mtime < cutoff:
                    shutil.rmtree(entry)
            except FileNotFoundError:
                continue


def handle_worker_ensure(
    payload: JsonDict,
    *,
    temp_dir: Path,
    cache_ttl_seconds: int = DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS,
) -> JsonDict:
    """Ensure the worker has a cached copy of a workbook and return its cache id."""
    session = ensure_worker_workbook(payload, temp_dir=temp_dir, cache_ttl_seconds=cache_ttl_seconds)
    return {
        "workbook_id": session.sha256,
        "workbook_name": session.workbook_path.name,
        "cache_hit": session.cache_hit,
        "uploaded_bytes": session.uploaded_bytes,
        "cache_prepare_ms": round(session.cache_prepare_ms, 3),
    }


def handle_worker_request(
    payload: JsonDict,
    *,
    temp_dir: Path,
    visible: bool,
    allow_macros: bool,
    max_cells_per_call: int,
    max_format_cells_per_call: int = DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
    cache_ttl_seconds: int = DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS,
) -> JsonDict:
    """Handle one Windows worker request payload."""
    started = time_module.perf_counter()
    operation = payload.get("operation")
    args = payload.get("args")
    if not isinstance(operation, str) or not isinstance(args, dict):
        msg = "operation and args are required"
        raise ValueError(msg)
    session = ensure_worker_workbook(payload, temp_dir=temp_dir, cache_ttl_seconds=cache_ttl_seconds)
    workbook_path = session.workbook_path
    service = ExcelService(
        workdir=workbook_path.parent,
        backend=LocalExcelBackend(visible=visible),
        allow_macros=allow_macros,
        max_cells_per_call=max_cells_per_call,
        max_format_cells_per_call=max_format_cells_per_call,
    )
    operation_started = time_module.perf_counter()
    if operation == "workbook_summary":
        result = service.workbook_summary(workbook_path.name, recalculate=bool(args.get("recalculate", True)))
    elif operation == "inspect_range":
        result = service.inspect_range(
            workbook_path.name,
            sheet=str(args["sheet"]),
            range_address=str(args["range_address"]),
            include_values=bool(args.get("include_values", True)),
            include_formulas=bool(args.get("include_formulas", True)),
            include_formats=bool(args.get("include_formats", False)),
            recalculate=bool(args.get("recalculate", True)),
        )
    elif operation == "find_cells":
        result = service.find_cells(
            workbook_path.name,
            query=str(args["query"]),
            search_values=bool(args.get("search_values", True)),
            search_formulas=bool(args.get("search_formulas", True)),
            max_results=int(args.get("max_results", 100)),
        )
    elif operation == "recalculate":
        result = service.recalculate(workbook_path.name)
    elif operation == "workbook_repair_check":
        result = service.workbook_repair_check(workbook_path.name)
    else:
        msg = f"Unsupported operation {operation!r}"
        raise ValueError(msg)
    add_result_metrics(
        result,
        {
            "windows_vm_worker_operation": operation,
            "windows_vm_worker_total_ms": round((time_module.perf_counter() - started) * 1000, 3),
            "windows_vm_worker_operation_ms": round((time_module.perf_counter() - operation_started) * 1000, 3),
            "windows_vm_worker_cache_prepare_ms": round(session.cache_prepare_ms, 3),
            "windows_vm_worker_workbook_cache_hit": session.cache_hit,
            "windows_vm_worker_uploaded_bytes": session.uploaded_bytes,
        },
    )
    return result


def ensure_worker_workbook(
    payload: JsonDict,
    *,
    temp_dir: Path,
    cache_ttl_seconds: int = DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS,
) -> _WorkerWorkbookSession:
    """Resolve a worker request's workbook to a content-addressed cache path."""
    started = time_module.perf_counter()
    workbook_name = payload.get("workbook_name")
    workbook_sha256 = payload.get("workbook_sha256")
    workbook_id = payload.get("workbook_id")
    workbook_b64 = payload.get("workbook_b64")
    if not isinstance(workbook_name, str) or not workbook_name:
        msg = "workbook_name is required"
        raise ValueError(msg)
    if not isinstance(workbook_sha256, str):
        msg = "workbook_sha256 is required"
        raise ValueError(msg)
    if not _SHA256_RE.fullmatch(workbook_sha256):
        msg = "workbook_sha256 must be a lowercase SHA-256 hex digest"
        raise ValueError(msg)
    if workbook_id is not None and workbook_id != workbook_sha256:
        msg = "workbook_id must match workbook_sha256"
        raise ValueError(msg)

    workbook_bytes: bytes | None = None
    if workbook_b64 is not None:
        if not isinstance(workbook_b64, str):
            msg = "workbook_b64 must be a string when supplied"
            raise ValueError(msg)
        workbook_bytes = base64.b64decode(workbook_b64)
        actual_sha = hashlib.sha256(workbook_bytes).hexdigest()
        if actual_sha != workbook_sha256:
            msg = "workbook_sha256 does not match workbook bytes"
            raise ValueError(msg)

    temp_dir.mkdir(parents=True, exist_ok=True)
    with worker_cache_lock(workbook_sha256):
        prune_worker_cache(temp_dir, ttl_seconds=cache_ttl_seconds)
        session_dir = temp_dir / workbook_sha256
        workbook_path = session_dir / Path(workbook_name).name
        cache_hit = workbook_path.exists()
        if not cache_hit:
            if workbook_bytes is None:
                msg = "workbook_not_cached"
                raise FileNotFoundError(msg)
            session_dir.mkdir(parents=True, exist_ok=True)
            workbook_path.write_bytes(workbook_bytes)
        os.utime(session_dir, None)
        os.utime(workbook_path, None)
    return _WorkerWorkbookSession(
        workbook_path=workbook_path,
        sha256=workbook_sha256,
        cache_hit=cache_hit,
        uploaded_bytes=len(workbook_bytes) if workbook_bytes is not None and not cache_hit else 0,
        cache_prepare_ms=(time_module.perf_counter() - started) * 1000,
    )


def main() -> None:
    """Run the Excel MCP server over stdio."""
    service = create_service(config_from_env())
    mcp = create_mcp(service)
    mcp.run(transport="stdio", show_banner=False)


def worker_main(argv: Sequence[str] | None = None) -> None:
    """Run the Windows VM Excel HTTP worker."""
    parser = argparse.ArgumentParser(description="Gandalf Excel worker service")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=positive_int, default=8765)
    parser.add_argument("--temp-dir", default=os.path.join(tempfile.gettempdir(), "gandalf_excel_worker"))
    parser.add_argument("--auth-token", default=os.environ.get("GANDALF_EXCEL_AUTH_TOKEN"))
    parser.add_argument("--visible", action="store_true")
    parser.add_argument("--allow-macros", action="store_true")
    parser.add_argument("--max-cells-per-call", type=positive_int, default=DEFAULT_EXCEL_MAX_CELLS_PER_CALL)
    parser.add_argument(
        "--max-format-cells-per-call",
        type=positive_int,
        default=DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
    )
    parser.add_argument("--cache-ttl-seconds", type=positive_int, default=DEFAULT_EXCEL_WORKER_CACHE_TTL_SECONDS)
    args = parser.parse_args(argv)
    run_worker(
        host=args.host,
        port=args.port,
        temp_dir=Path(args.temp_dir),
        auth_token=args.auth_token,
        visible=args.visible,
        allow_macros=args.allow_macros,
        max_cells_per_call=args.max_cells_per_call,
        max_format_cells_per_call=args.max_format_cells_per_call,
        cache_ttl_seconds=args.cache_ttl_seconds,
    )


if __name__ == "__main__":
    main()
