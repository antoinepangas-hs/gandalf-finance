"""Outer grader orchestrator.

Runs as the grader user and spawns the inner judge as the sandbox user
(via sudo) to evaluate rubric criteria using an OpenHands agent-as-judge.

Supports two evaluation modes (configured via ``mode`` in the TOML config):
  - **batch** (default): criteria evaluated in one or more batch sessions.
  - **individual**: one agent session per rubric criterion.

Batch mode auto-splits by default and adaptively backs off when rate limits are
detected.  ``batch_splits`` and ``max_concurrency`` can still be set for fixed
parallelism.

Produces (in ``output_dir``):
  reward.json  - Reward file ([0,1] reward)
  info.json    - Detailed per-criterion results + LLM usage
"""

import argparse
import contextlib
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field as dataclass_field
from pathlib import Path
from typing import Any, Callable, Generic, Iterator, Literal, TypeVar, cast

from pydantic import TypeAdapter

from gandalf.excel_mcp import ExcelServerConfig, create_service, merge_excel_metrics_summaries
from gandalf.models import (
    BatchJudgeInput,
    CriterionResult,
    ExcelBackendConfig,
    EvaluationInfo,
    GraderConfig,
    JudgeInput,
    LLMUsage,
    RubricItem,
    SectionGateResult,
    SectionResult,
    Verdict,
    WorkbookRepairCheckResult,
    load_config,
    load_rubric,
)
from gandalf.paths import ensure_work_root, excel_access_root, is_protected_path
from gandalf.workbook_digest import build_workbook_digest

GOLDEN_CHECK_GUIDANCE = (
    "Golden check mode is enabled. Section-level tolerance bands are disabled for this run; "
    "evaluate numeric values exactly unless an individual criterion explicitly states its own tolerance."
)
MICROSOFT_EXCEL_BACKEND_GUIDANCE = (
    "Microsoft Excel-backed workbook inspection is available through the MCP server named 'excel'. "
    "Use the Excel MCP tools for workbook values, formulas, 3D references, external links, formats, "
    "and recalculation-sensitive checks. Do not use LibreOffice recalculation for Excel workbooks; "
    "LibreOffice can mis-handle Excel-specific formulas such as 3D references. "
    "Range inspection omits per-cell formats unless include_formats=true; set include_formats=true "
    "for number-formatting or formatting-specific rubric checks."
)
LIBREOFFICE_BACKEND_GUIDANCE = (
    "Structured workbook inspection is available through the MCP server named 'excel'. "
    "For this run those tools are backed by LibreOffice headless plus openpyxl. Use the MCP tools "
    "for workbook sheets, values, formulas, formats, searches, and LibreOffice recalculation. "
    "Range inspection omits per-cell formats unless include_formats=true; set include_formats=true "
    "for number-formatting or formatting-specific rubric checks. "
    "When recalculation is requested, Gandalf configures LibreOffice to always recalculate imported "
    "OOXML/ODF formulas before reading cached values, but those values still come from LibreOffice's "
    "formula engine. "
    "Do not treat LibreOffice results as guaranteed Microsoft Excel parity for Excel-specific "
    "features such as 3D references, external links, macros, or repair dialogue detection."
)
RATE_LIMIT_MARKERS = ("429", "rate limit", "rate_limit_exceeded", "too many requests")
BATCH_RECOVERY_ERROR_MARKERS = (
    "empty verdict file",
    "did not write a verdict file",
    "invalid json",
    "failed to read judge output",
    "judge execution timed out",
    "judge process failed",
    "judge execution error",
    "batch split failed",
    "no reasoning provided",
    "missing",
    "timed out",
    "timeout",
    "empty response",
    "stuck",
)
T = TypeVar("T")
CliExcelBackend = Literal[
    "config",
    "mac-excel",
    "mac_excel",
    "windows-vm",
    "windows_vm",
    "excel-vm",
    "excel_vm",
    "libreoffice",
    "libre-office",
    "disabled",
    "none",
]


def load_trajectory_final_output(path: str) -> str:
    """Load an ATIF trajectory file and extract the final agent message."""
    with open(path) as f:
        data = json.load(f)

    steps = data.get("steps", [])

    # Extract final agent message (last with non-empty content, no tool calls)
    final_output = ""
    for step in reversed(steps):
        if step.get("source") == "agent" and not step.get("tool_calls"):
            msg = step.get("message", "")
            if msg.strip():
                final_output = msg
                break

    return final_output


# Environment variables forwarded to the inner judge subprocess (via sudo).
# Only these are passed — everything else is stripped to avoid leaking secrets
# or host-specific state into the sandbox.
JUDGE_ENV_ALLOWLIST = frozenset(
    {
        "PATH",
        "LLM_API_KEY",
        "LLM_BASE_URL",
        "PYTHONPATH",
        "UV_TOOL_DIR",
        "UV_TOOL_BIN_DIR",
        "UV_PYTHON_INSTALL_DIR",
        "GANDALF_EXCEL_CACHE_DIR",
        "GANDALF_EXCEL_DEBUG_METRICS",
        # OpenTelemetry — forwarded so the inner judge can export traces
        # to any OTEL-compatible backend (e.g. Langfuse, Jaeger, Honeycomb).
        "OTEL_EXPORTER_OTLP_ENDPOINT",
        "OTEL_EXPORTER_OTLP_HEADERS",
        "OTEL_EXPORTER_OTLP_TRACES_PROTOCOL",
    }
)


def judge_env_vars() -> list[str]:
    """Build the ``KEY=VALUE`` list for the judge subprocess environment."""
    return [f"{k}={v}" for k, v in os.environ.items() if k in JUDGE_ENV_ALLOWLIST and v]


def excel_metrics_path_for_trace(trace_path: str) -> str:
    """Return the output-side metrics sidecar path associated with a judge trace."""
    trace = Path(trace_path)
    return str(trace.with_name(f"{trace.stem}_excel_metrics.json"))


def clear_excel_metrics(output_dir: str) -> None:
    """Remove stale Excel metrics sidecars from a previous run in the same output dir."""
    output_path = Path(output_dir)
    if not output_path.exists():
        return
    for metrics_path in output_path.glob("*_excel_metrics.json"):
        with contextlib.suppress(OSError):
            metrics_path.unlink()


@contextlib.contextmanager
def libreoffice_cache_environment(config: GraderConfig) -> Iterator[None]:
    """Expose a run-scoped LibreOffice cache directory to judge subprocesses."""
    if not config.excel_backend.enabled or config.excel_backend.backend != "libreoffice":
        yield
        return

    cache_dir = tempfile.mkdtemp(prefix="gandalf_libreoffice_run_cache_")
    with contextlib.suppress(OSError):
        os.chmod(cache_dir, 0o777)  # noqa: S103
    previous = os.environ.get("GANDALF_EXCEL_CACHE_DIR")
    os.environ["GANDALF_EXCEL_CACHE_DIR"] = cache_dir
    try:
        yield
    finally:
        if previous is None:
            os.environ.pop("GANDALF_EXCEL_CACHE_DIR", None)
        else:
            os.environ["GANDALF_EXCEL_CACHE_DIR"] = previous
        shutil.rmtree(cache_dir, ignore_errors=True)


def persist_excel_metrics(clone_metrics_path: str, trace_path: str) -> None:
    """Copy one judge clone's Excel metrics beside its trace, if metrics were emitted."""
    source = Path(clone_metrics_path)
    if not source.is_file():
        return
    target = Path(excel_metrics_path_for_trace(trace_path))
    with contextlib.suppress(OSError):
        shutil.copyfile(source, target)


def collect_excel_metrics(output_dir: str) -> dict[str, object] | None:
    """Return merged Excel metrics sidecars for an output directory."""
    summaries: list[dict[str, Any]] = []
    for metrics_path in sorted(Path(output_dir).glob("*_excel_metrics.json")):
        try:
            data = json.loads(metrics_path.read_text())
        except (OSError, json.JSONDecodeError):
            continue
        if isinstance(data, dict):
            summaries.append(data)
    if not summaries:
        return None
    merged = merge_excel_metrics_summaries(summaries)
    return merged if merged.get("call_count", 0) else None


def resolve_optional_file(
    inline: str | None,
    path: str | None,
    label: str,
) -> str | None:
    """Return *inline* content, or read from *path*, or ``None``.

    The caller is expected to ensure *inline* and *path* are mutually
    exclusive (enforced by ``GraderConfig``'s model validator).  If a
    path is given but does not exist, exits with a clear error.
    """
    if inline is not None:
        return inline
    if not path:
        return None
    if not os.path.isfile(path):
        print(  # noqa: T201
            f"ERROR: File not found: {path}\n  Configured via: {label}",
            file=sys.stderr,
        )
        sys.exit(1)
    with open(path) as f:
        return f.read()


def resolve_config_value(
    inline: str | None,
    config_path: str | None,
    env_var: str,
    config_label: str,
) -> str | None:
    """Resolve a config value: inline content → config path → env var → ``None``."""
    path = config_path or os.environ.get(env_var)
    source = f"{config_label} in grader config" if config_path else f"{env_var} env var"
    return resolve_optional_file(inline, path, source)


def resolve_instructions(config: GraderConfig) -> str:
    """Resolve task instructions (inline, path, or env var).

    Resolution order:
      1. config.instructions (inline in TOML)
      2. config.instructions_path (from TOML)
      3. GRADER_INSTRUCTIONS_PATH env var
      4. Error — instructions are required
    """
    result = resolve_config_value(
        config.instructions,
        config.instructions_path,
        "GRADER_INSTRUCTIONS_PATH",
        "instructions_path",
    )
    if not result:
        print(  # noqa: T201
            "ERROR: No instructions provided. Set 'instructions' or 'instructions_path' "
            "in the config, or the GRADER_INSTRUCTIONS_PATH env var.",
            file=sys.stderr,
        )
        sys.exit(1)
    return result


def resolve_judge_prompt(config: GraderConfig) -> str | None:
    """Resolve the custom judge prompt template (inline, path, or env var).

    Resolution order:
      1. config.judge_prompt (inline in TOML)
      2. config.judge_prompt_path (from TOML)
      3. GRADER_JUDGE_PROMPT_PATH env var
      4. No custom template (returns None, uses built-in)
    """
    return resolve_config_value(
        config.judge_prompt,
        config.judge_prompt_path,
        "GRADER_JUDGE_PROMPT_PATH",
        "judge_prompt_path",
    )


def resolve_judge_guidance(config: GraderConfig) -> str:
    """Resolve judge guidance content (inline, path, or env var).

    Resolution order:
      1. config.judge_guidance (inline in TOML)
      2. config.judge_guidance_path (from TOML)
      3. GRADER_JUDGE_GUIDANCE_PATH env var
      4. No guidance (empty string)
    """
    return (
        resolve_config_value(
            config.judge_guidance,
            config.judge_guidance_path,
            "GRADER_JUDGE_GUIDANCE_PATH",
            "judge_guidance_path",
        )
        or ""
    )


def effective_judge_guidance(
    judge_guidance: str,
    *,
    golden_check: bool,
    excel_backend: ExcelBackendConfig | None = None,
    workbook_digest: str | None = None,
) -> str:
    """Return judge guidance with any runtime-mode instructions applied."""
    guidance_parts: list[str] = []
    if golden_check:
        guidance_parts.append(GOLDEN_CHECK_GUIDANCE)
    if excel_backend is not None and excel_backend.enabled:
        guidance_parts.append(workbook_backend_guidance(excel_backend))
    if workbook_digest:
        guidance_parts.append(workbook_digest)
    if judge_guidance:
        guidance_parts.append(judge_guidance)
    return "\n\n".join(guidance_parts)


def workbook_backend_guidance(excel_backend: ExcelBackendConfig) -> str:
    """Return judge guidance for the configured workbook backend."""
    if excel_backend.backend == "libreoffice":
        return LIBREOFFICE_BACKEND_GUIDANCE
    return MICROSOFT_EXCEL_BACKEND_GUIDANCE


def apply_golden_check_defaults(config: GraderConfig) -> GraderConfig:
    """Enable golden-check performance defaults unless the config explicitly overrides them."""
    if not config.golden_check:
        return config

    if "enabled" in config.workbook_digest.model_fields_set:
        return config
    return config.model_copy(
        update={"workbook_digest": config.workbook_digest.model_copy(update={"enabled": True})}
    )


def workbook_digest_text(config: GraderConfig) -> str | None:
    """Return the judge-prompt workbook digest for *config*, or None when disabled/unavailable."""
    digest_config = config.workbook_digest
    if not digest_config.enabled:
        return None
    workbook = digest_config.workbook_path or default_repair_check_workbook(config.workdir)
    if workbook is None:
        return None
    workdir = Path(config.workdir).resolve()
    path = Path(workbook)
    path = path.resolve() if path.is_absolute() else (workdir / path).resolve()
    try:
        path.relative_to(workdir)
    except ValueError as e:
        msg = f"workbook_digest.workbook_path {workbook!r} is outside the allowed workdir"
        raise ValueError(msg) from e
    body = build_workbook_digest(
        path,
        max_sheets=digest_config.max_sheets,
        max_formula_samples=digest_config.max_formula_samples,
        max_preview_columns=digest_config.max_preview_columns,
        max_named_ranges=digest_config.max_named_ranges,
    )
    if not body:
        return None
    return (
        f"Workbook digest for {path.name} (deterministic openpyxl read; cached values are NOT "
        "recalculated, so use the Excel MCP tools for actual values). Use it to locate sheets, "
        "headers, and formulas before inspecting cells:\n" + body
    )


def apply_excel_backend_cli_override(
    config: GraderConfig,
    *,
    backend: CliExcelBackend | None,
    windows_url: str | None = None,
    auth_token: str | None = None,
) -> GraderConfig:
    """Return *config* with CLI Excel backend overrides applied."""
    if backend is None and windows_url is None and auth_token is None:
        return config

    data = config.excel_backend.model_dump()
    if windows_url is not None:
        data["windows_url"] = windows_url
    if auth_token is not None:
        data["auth_token"] = auth_token

    if backend in {"libreoffice", "libre-office"}:
        data["enabled"] = True
        data["backend"] = "libreoffice"
    elif backend in {"disabled", "none"}:
        data["enabled"] = False
        data["backend"] = "mac_excel"
    elif backend in {"mac-excel", "mac_excel"}:
        data["enabled"] = True
        data["backend"] = "mac_excel"
    elif backend in {"windows-vm", "windows_vm", "excel-vm", "excel_vm"}:
        data["enabled"] = True
        data["backend"] = "windows_vm"
    elif backend == "config" or backend is None:
        pass
    else:
        msg = f"Unsupported --excel-backend value {backend!r}"
        raise ValueError(msg)

    try:
        excel_backend = ExcelBackendConfig.model_validate(data)
    except Exception as e:  # noqa: BLE001
        msg = f"Invalid Excel backend override: {e}"
        raise ValueError(msg) from e

    if excel_backend.enabled and any(server.name == "excel" for server in config.mcp_servers):
        msg = (
            "excel_backend.enabled automatically attaches an MCP server named 'excel'; "
            "remove or rename the configured MCP server named 'excel'"
        )
        raise ValueError(msg)

    return config.model_copy(update={"excel_backend": excel_backend})


def clone_workspace(src: str, *, parent_dir: str | None = None) -> str:
    """Clone workspace into a temp directory accessible to the sandbox user.

    Walks the source tree once, skipping unreadable directories and files with
    a warning.  Each directory and file is made world-accessible inline so no
    second pass is needed.

    ``shutil.copytree`` is not used because its ``copy_function`` hook only
    covers per-file errors — directory listing errors (e.g. a 0o700 dir owned
    by the agent) cannot be caught there.
    """
    if parent_dir is not None:
        os.makedirs(parent_dir, exist_ok=True)
    clone_dir = tempfile.mkdtemp(prefix="judge_workspace_", dir=parent_dir)
    # Root dir is created by mkdtemp at 0o700; open it up immediately so
    # sandbox_user can traverse and write to it.
    os.chmod(clone_dir, 0o777)  # noqa: S103
    skipped: list[str] = []

    def on_walk_error(err: OSError) -> None:
        skipped.append(err.filename or str(err))

    for dirpath, _dirnames, filenames in os.walk(src, onerror=on_walk_error):
        rel = os.path.relpath(dirpath, src)
        dst_dir = os.path.join(clone_dir, rel)
        os.makedirs(dst_dir, exist_ok=True)
        os.chmod(dst_dir, 0o777)  # noqa: S103

        for fname in filenames:
            src_file = os.path.join(dirpath, fname)
            dst_file = os.path.join(dst_dir, fname)
            try:
                shutil.copyfile(src_file, dst_file)
                # Preserve execute bits from source so scripts/binaries
                # remain runnable, while granting world read/write.
                src_mode = os.stat(src_file).st_mode
                os.chmod(dst_file, 0o666 | (src_mode & 0o111))
            except OSError:
                # Covers PermissionError, FileNotFoundError (broken symlinks),
                # IsADirectoryError (symlinks to dirs in filenames), etc.
                skipped.append(src_file)

    max_skipped_log = 20
    if skipped:
        print(  # noqa: T201
            f"[gandalf] workspace clone: skipped {len(skipped)} unreadable path(s):",
            file=sys.stderr,
        )
        for p in skipped[:max_skipped_log]:
            print(f"  - {p}", file=sys.stderr)  # noqa: T201
        if len(skipped) > max_skipped_log:
            print(f"  ... and {len(skipped) - max_skipped_log} more", file=sys.stderr)  # noqa: T201

    return clone_dir


def run_judge(
    judge_input: JudgeInput | BatchJudgeInput,
    sandbox_user: str | None,
    trace_path: str,
    timeout: int = 300,
    clone_parent_dir: str | None = None,
) -> tuple[list[Verdict], LLMUsage]:
    """Clone workspace, run the judge subprocess, and return parsed verdicts.

    Always returns a *list* of verdicts, even for a single-criterion
    ``JudgeInput`` (one-element list).  On any subprocess failure every
    verdict is set to ``met=None`` with the error message.
    """
    batch = isinstance(judge_input, BatchJudgeInput)
    n = len(judge_input.criteria) if isinstance(judge_input, BatchJudgeInput) else 1

    def fail(msg: str) -> tuple[list[Verdict], LLMUsage]:
        return Verdict.errors(n, msg), LLMUsage()

    try:
        clone_dir = clone_workspace(judge_input.workdir, parent_dir=clone_parent_dir)
    except Exception as e:  # noqa: BLE001
        return fail(f"Failed to clone workspace: {e}")

    cloned_input = judge_input.model_copy(update={"workdir": clone_dir})
    clone_excel_metrics_path = os.path.join(clone_dir, "excel_metrics.json")

    prefix = "judge_batch_" if batch else "judge_"
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".json",
        prefix=f"{prefix}input_",
        dir=clone_dir,
        delete=False,
    ) as input_f:
        input_f.write(cloned_input.model_dump_json())
        input_path = input_f.name

    # Pre-create the output file so sandbox_user can write to it without
    # needing general write access to /tmp (which may not be world-writable).
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".json",
        prefix=f"{prefix}output_",
        dir=clone_dir,
        delete=False,
    ) as output_f:
        output_path = output_f.name
    os.chmod(output_path, 0o666)  # noqa: S103

    try:
        os.chmod(input_path, 0o644)
        env_vars = [
            f"HOME={clone_dir}",
            *judge_env_vars(),
            f"GANDALF_EXCEL_METRICS_PATH={clone_excel_metrics_path}",
        ]

        cmd = []
        if sandbox_user is not None:
            cmd += ["sudo", "-u", sandbox_user]
        cmd += [
            "env",
            *env_vars,
            "gandalf-the-grader-judge",
            "--input",
            input_path,
            "--output",
            output_path,
        ]
        if batch:
            cmd.append("--batch")

        result = subprocess.run(
            cmd,
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=clone_dir,
        )

        save_trace(trace_path, result.stdout, result.stderr, result.returncode)
        persist_excel_metrics(clone_excel_metrics_path, trace_path)

        if result.returncode != 0:
            return fail(f"Judge process failed (exit {result.returncode}): {result.stderr[:500]}")

        with open(output_path) as f:
            data = json.load(f)

    except subprocess.TimeoutExpired:
        save_trace(trace_path, "", "Judge execution timed out.", -1)
        persist_excel_metrics(clone_excel_metrics_path, trace_path)
        return fail("Judge execution timed out.")
    except (json.JSONDecodeError, FileNotFoundError) as e:
        return fail(f"Failed to read judge output: {e}")
    else:
        if batch:
            verdicts = TypeAdapter(list[Verdict]).validate_python(data["verdicts"])
        else:
            verdicts = [Verdict.model_validate(data["verdict"])]
        usage = LLMUsage.model_validate(data["llm_usage"])
        return verdicts, usage
    finally:
        shutil.rmtree(clone_dir, ignore_errors=True)


def judge_clone_parent_dir(config: GraderConfig) -> str | None:
    """Return a stable clone parent for Mac Excel judge workspaces.

    Judge workspace clones are placed under the stable ``excel_access`` root (not
    the per-run, timestamped output dir) so a single one-time macOS "Grant
    Access" to that folder covers every future clone and Excel stops prompting.
    """
    if config.excel_backend.enabled and config.excel_backend.backend == "mac_excel":
        return str(excel_access_root() / "judge_workspaces")
    return None


def save_trace(trace_path: str, stdout: str, stderr: str, returncode: int) -> None:
    """Write the judge's stdout/stderr to a trace file."""
    with contextlib.suppress(OSError), open(trace_path, "w") as f:
        f.write(f"exit_code: {returncode}\n")
        f.write("=== stdout ===\n")
        f.write(stdout)
        f.write("\n=== stderr ===\n")
        f.write(stderr)


def format_status(*, met: bool | None) -> str:
    """Format criterion evaluation status for display."""
    if met is True:
        return "MET"
    if met is None:
        return "ERROR"
    return "UNMET"


def verdict_to_result(item: RubricItem, verdict: Verdict) -> CriterionResult:
    """Convert a Verdict into a CriterionResult for the given rubric item."""
    return CriterionResult(
        criterion=item.criterion,
        weight=item.weight,
        section=item.section,
        gate=item.gate,
        met=verdict.met,
        reasoning=verdict.reasoning,
        evidence=verdict.evidence,
        section_tolerance_pct=item.section_tolerance_pct,
    )


def skipped_result(item: RubricItem, reason: str) -> CriterionResult:
    """Build a skipped CriterionResult for criteria bypassed by a section gate."""
    return CriterionResult(
        criterion=item.criterion,
        weight=item.weight,
        section=item.section,
        gate=item.gate,
        met=None,
        reasoning=reason,
        section_tolerance_pct=item.section_tolerance_pct,
        skipped=True,
        skip_reason=reason,
    )


GATES_ONLY_SKIP_REASON = "Skipped because --gates-only was set."


def format_tolerance_pct(tolerance_pct: float) -> str:
    """Format a configured percentage tolerance for prompt text."""
    return f"{tolerance_pct:g}"


def criterion_for_judge(item: RubricItem, *, golden_check: bool) -> str:
    """Return criterion text as sent to the judge, including effective tolerance guidance."""
    if golden_check or item.gate or item.section_tolerance_pct is None:
        return item.criterion
    tolerance = format_tolerance_pct(item.section_tolerance_pct)
    return (
        f"{item.criterion}\n\n"
        "Section-level numeric tolerance: when this criterion compares submitted numeric values "
        f"to expected or reference numeric values, treat values within +/-{tolerance}% of the "
        "expected/reference value as meeting the numeric requirement. For an expected/reference "
        "value of 0, require exact zero unless this criterion explicitly states another tolerance. "
        "This tolerance does not affect qualitative, formatting, structural, presence, or "
        "nonnumeric checks."
    )


@dataclass(frozen=True)
class _AdaptiveRunResult(Generic[T]):
    """One scheduler task result."""

    value: T
    usage: LLMUsage
    rate_limited: bool = False


@dataclass(frozen=True)
class _AdaptiveWorkItem(Generic[T]):
    """One scheduler task to execute."""

    index: int
    label: str
    run: Callable[[], _AdaptiveRunResult[T]]


def is_rate_limit_error_text(text: str) -> bool:
    """Return whether text looks like a provider rate-limit failure."""
    lowered = text.lower()
    return any(marker in lowered for marker in RATE_LIMIT_MARKERS)


def verdicts_hit_rate_limit(verdicts: list[Verdict]) -> bool:
    """Return whether any unresolved verdict looks rate-limit related."""
    return any(verdict.met is None and is_rate_limit_error_text(verdict.reasoning) for verdict in verdicts)


def section_gate_results_hit_rate_limit(results: list[SectionGateResult]) -> bool:
    """Return whether any unresolved section-gate result looks rate-limit related."""
    return any(result.met is None and not result.skipped and is_rate_limit_error_text(result.reasoning) for result in results)


def batch_timeout_for_count(config: GraderConfig, count: int) -> int:
    """Return the timeout for a batch judge session with *count* criteria."""
    timeout = config.judge_timeout * count
    if config.batch_timeout is not None:
        timeout = min(timeout, config.batch_timeout)
    return timeout


def split_count_for_batch(config: GraderConfig, count: int) -> int:
    """Return the number of batch chunks to use for *count* criteria."""
    if count <= 1:
        return count
    if config.batch_splits is not None:
        return min(config.batch_splits, count)
    if not config.auto_parallel:
        return 1
    return min(count, max(1, math.ceil(count / config.auto_parallel_target_chunk_size)))


def split_count_for_auto_parallel(config: GraderConfig, count: int) -> int:
    """Return auto-parallel chunk count for non-rubric batch work."""
    if count <= 1:
        return count
    if not config.auto_parallel:
        return 1
    return min(count, max(1, math.ceil(count / config.auto_parallel_target_chunk_size)))


def chunk_indexed_items(items: list[T], splits: int) -> list[list[tuple[int, T]]]:
    """Split items into positional chunks preserving original indices."""
    if not items:
        return []
    chunk_size = math.ceil(len(items) / max(1, splits))
    return [
        [(i, items[i]) for i in range(start, min(start + chunk_size, len(items)))]
        for start in range(0, len(items), chunk_size)
    ]


def effective_concurrency_cap(config: GraderConfig, *, total: int, discovery_default_cap: int) -> int:
    """Return the effective maximum scheduler concurrency."""
    if total <= 0:
        return 0
    configured_cap = config.max_concurrency if config.max_concurrency is not None else discovery_default_cap
    return max(1, min(total, configured_cap))


def execute_adaptive_work(
    config: GraderConfig,
    work_items: list[_AdaptiveWorkItem[T]],
    *,
    label: str,
    fixed_concurrency: int,
    discovery_default_cap: int,
) -> tuple[list[T], LLMUsage]:
    """Execute work items with optional adaptive rate-limit discovery."""
    if not work_items:
        return [], LLMUsage()

    total_usage = LLMUsage()
    results_by_index: dict[int, T] = {}

    if not config.rate_limit_discovery:
        concurrency = max(1, min(len(work_items), config.max_concurrency or fixed_concurrency))
        print(f"[{label}] Running {len(work_items)} task(s) with max_concurrency={concurrency}")  # noqa: T201
        with ThreadPoolExecutor(max_workers=concurrency) as executor:
            futures = {executor.submit(item.run): item for item in work_items}
            for future, item in futures.items():
                run_result = future.result()
                results_by_index[item.index] = run_result.value
                total_usage = total_usage + run_result.usage
        return [results_by_index[item.index] for item in sorted(work_items, key=lambda item: item.index)], total_usage

    cap = effective_concurrency_cap(config, total=len(work_items), discovery_default_cap=discovery_default_cap)
    current_concurrency = max(1, min(config.rate_limit_initial_concurrency, cap))
    last_successful_concurrency = 1
    discovery_frozen = False
    pending = list(work_items)
    requeues = {item.index: 0 for item in work_items}

    while pending:
        wave = pending[:current_concurrency]
        pending = pending[current_concurrency:]
        print(  # noqa: T201
            f"[{label}] Running {len(wave)} task(s) at concurrency={current_concurrency} "
            f"(remaining={len(pending)})"
        )
        wave_rate_limited = False
        requeued_in_wave = 0
        with ThreadPoolExecutor(max_workers=min(current_concurrency, len(wave))) as executor:
            futures = {executor.submit(item.run): item for item in wave}
            for future, item in futures.items():
                run_result = future.result()
                total_usage = total_usage + run_result.usage
                if run_result.rate_limited:
                    wave_rate_limited = True
                    requeues[item.index] += 1
                    if requeues[item.index] <= config.rate_limit_max_requeues_per_chunk:
                        pending.append(item)
                        requeued_in_wave += 1
                        continue
                    print(  # noqa: T201
                        f"[{label}] Rate limit persisted for {item.label}; "
                        f"keeping errored result after {requeues[item.index]} attempts"
                    )
                results_by_index[item.index] = run_result.value

        if wave_rate_limited:
            discovery_frozen = True
            current_concurrency = max(1, min(last_successful_concurrency, cap))
            if requeued_in_wave:
                print(  # noqa: T201
                    f"[{label}] Rate limit detected; backing off for "
                    f"{config.rate_limit_backoff_seconds}s and continuing at concurrency={current_concurrency}"
                )
                time.sleep(config.rate_limit_backoff_seconds)
        elif not discovery_frozen:
            last_successful_concurrency = current_concurrency
            current_concurrency = min(cap, current_concurrency * 2)

    ordered_items = sorted(work_items, key=lambda item: item.index)
    return [results_by_index[item.index] for item in ordered_items], total_usage


def run_individual(
    config: GraderConfig,
    rubric: list[RubricItem],
    final_output: str,
    instructions: str,
    judge_guidance: str,
    judge_prompt: str | None,
    trace_suffix: str = "",
) -> tuple[list[CriterionResult], LLMUsage]:
    """Evaluate each rubric item in its own agent session.

    When max_concurrency > 1, up to N criteria are evaluated in parallel
    via a thread pool.  Results are always returned in rubric order.
    """
    n = len(rubric)

    def _work_item(i: int, item: RubricItem) -> _AdaptiveWorkItem[CriterionResult]:
        def _run() -> _AdaptiveRunResult[CriterionResult]:
            print(f"[{i + 1}/{n}] Evaluating: {item.criterion[:80]}...")  # noqa: T201
            judge_input = JudgeInput(
                model=config.model,
                reasoning_effort=config.reasoning_effort,
                instructions=instructions,
                final_output=final_output,
                criterion=criterion_for_judge(item, golden_check=config.golden_check),
                workdir=config.workdir,
                mcp_servers=config.mcp_servers,
                excel_backend=config.excel_backend,
                judge_guidance=judge_guidance,
                judge_prompt=judge_prompt,
            )
            trace_path = os.path.join(config.output_dir, f"judge_trace_{i}{trace_suffix}.txt")
            verdicts, usage = run_judge(
                judge_input,
                sandbox_user=config.sandbox_user,
                trace_path=trace_path,
                timeout=config.judge_timeout,
                clone_parent_dir=judge_clone_parent_dir(config),
            )
            result = verdict_to_result(item, verdicts[0])
            print(  # noqa: T201
                f"  [{i + 1}/{n}] {format_status(met=verdicts[0].met)}: {verdicts[0].reasoning[:120]}"
            )
            return _AdaptiveRunResult(
                value=result,
                usage=usage,
                rate_limited=verdicts_hit_rate_limit(verdicts),
            )

        return _AdaptiveWorkItem(index=i, label=f"criterion {i}", run=_run)

    work_items = [_work_item(i, item) for i, item in enumerate(rubric)]
    fixed_concurrency = config.max_concurrency or 1
    return execute_adaptive_work(
        config,
        work_items,
        label="individual",
        fixed_concurrency=fixed_concurrency,
        discovery_default_cap=fixed_concurrency,
    )


def run_batch(
    config: GraderConfig,
    rubric: list[RubricItem],
    final_output: str,
    instructions: str,
    judge_guidance: str,
    judge_prompt: str | None,
    trace_suffix: str = "",
) -> tuple[list[CriterionResult], LLMUsage]:
    """Evaluate all rubric items in a single agent session."""
    criteria = [criterion_for_judge(item, golden_check=config.golden_check) for item in rubric]
    n = len(criteria)

    batch_timeout = batch_timeout_for_count(config, n)

    print(f"[batch] Evaluating {n} criteria in one session (timeout={batch_timeout}s)...")  # noqa: T201
    judge_input = BatchJudgeInput(
        model=config.model,
        reasoning_effort=config.reasoning_effort,
        instructions=instructions,
        final_output=final_output,
        criteria=criteria,
        workdir=config.workdir,
        mcp_servers=config.mcp_servers,
        excel_backend=config.excel_backend,
        judge_guidance=judge_guidance,
        judge_prompt=judge_prompt,
    )
    trace_path = os.path.join(config.output_dir, f"judge_trace_batch{trace_suffix}.txt")
    verdicts, usage = run_judge(
        judge_input,
        sandbox_user=config.sandbox_user,
        trace_path=trace_path,
        timeout=batch_timeout,
        clone_parent_dir=judge_clone_parent_dir(config),
    )

    results: list[CriterionResult] = []
    for i, item in enumerate(rubric):
        v = verdicts[i] if i < len(verdicts) else Verdict(met=None, reasoning="No reasoning provided.")
        results.append(verdict_to_result(item, v))
        print(f"  [{i + 1}/{n}] {format_status(met=v.met)}: {v.reasoning[:120]}")  # noqa: T201
    return results, usage


def run_batch_concurrent(
    config: GraderConfig,
    rubric: list[RubricItem],
    final_output: str,
    instructions: str,
    judge_guidance: str,
    judge_prompt: str | None,
    trace_suffix: str = "",
) -> tuple[list[CriterionResult], LLMUsage]:
    """Split criteria into N positional chunks and evaluate each as a parallel batch.

    Each chunk is sent to its own judge subprocess.  All chunks run in parallel
    via a thread pool (each thread blocks on subprocess.run).  Results are merged
    back in original rubric order.
    """
    n = len(rubric)
    if n == 0:
        return [], LLMUsage()
    splits = split_count_for_batch(config, n)
    chunks = chunk_indexed_items(rubric, splits)

    print(  # noqa: T201
        f"[batch-concurrent] Splitting {n} criteria into {len(chunks)} chunks "
        f"(sizes: {', '.join(str(len(c)) for c in chunks)})"
    )

    def _run_split(
        split_idx: int, chunk: list[tuple[int, RubricItem]]
    ) -> _AdaptiveRunResult[list[tuple[int, CriterionResult]]]:
        # Use local 0-based indices for the judge — the prompt says
        # "0 through N-1" and read_batch_verdict filters by 0 <= idx < N.
        # Global rubric indices are restored when building indexed_results.
        criteria_list = [criterion_for_judge(item, golden_check=config.golden_check) for _orig_idx, item in chunk]

        n_criteria = len(criteria_list)
        batch_timeout = batch_timeout_for_count(config, n_criteria)

        print(  # noqa: T201
            f"  [split {split_idx + 1}/{len(chunks)}] {n_criteria} criteria (timeout={batch_timeout}s)..."
        )

        judge_input = BatchJudgeInput(
            model=config.model,
            reasoning_effort=config.reasoning_effort,
            instructions=instructions,
            final_output=final_output,
            criteria=criteria_list,
            workdir=config.workdir,
            mcp_servers=config.mcp_servers,
            excel_backend=config.excel_backend,
            judge_guidance=judge_guidance,
            judge_prompt=judge_prompt,
        )

        trace_path = os.path.join(config.output_dir, f"judge_trace_batch_split{split_idx}{trace_suffix}.txt")
        verdicts, usage = run_judge(
            judge_input,
            sandbox_user=config.sandbox_user,
            trace_path=trace_path,
            timeout=batch_timeout,
            clone_parent_dir=judge_clone_parent_dir(config),
        )

        indexed_results: list[tuple[int, CriterionResult]] = []
        for j, (orig_idx, item) in enumerate(chunk):
            v = verdicts[j] if j < len(verdicts) else Verdict(met=None, reasoning="No reasoning provided.")
            indexed_results.append((orig_idx, verdict_to_result(item, v)))
            print(  # noqa: T201
                f"    [{orig_idx + 1}/{n}] {format_status(met=v.met)}: {v.reasoning[:120]}"
            )

        return _AdaptiveRunResult(
            value=indexed_results,
            usage=usage,
            rate_limited=verdicts_hit_rate_limit(verdicts),
        )

    def _make_split_work_item(
        split_idx: int,
        chunk: list[tuple[int, RubricItem]],
    ) -> _AdaptiveWorkItem[list[tuple[int, CriterionResult]]]:
        def _run_bound_split() -> _AdaptiveRunResult[list[tuple[int, CriterionResult]]]:
            return _run_split(split_idx, chunk)

        return _AdaptiveWorkItem(index=split_idx, label=f"batch split {split_idx}", run=_run_bound_split)

    # Run all splits in parallel
    work_items = [_make_split_work_item(split_idx, chunk) for split_idx, chunk in enumerate(chunks)]

    try:
        split_results, total_usage = execute_adaptive_work(
            config,
            work_items,
            label="batch-concurrent",
            fixed_concurrency=len(chunks),
            discovery_default_cap=config.rate_limit_max_concurrency,
        )
    except Exception as exc:  # noqa: BLE001
        # All-or-nothing: if any split raises, we fail *all* criteria so
        # the hard-fail path in main() writes info.json but not reward.json.
        print(f"[batch-concurrent] Split failed unexpectedly: {exc}", file=sys.stderr)  # noqa: T201
        return (
            [
                CriterionResult(
                    criterion=item.criterion,
                    weight=item.weight,
                    section=item.section,
                    gate=item.gate,
                    met=None,
                    reasoning=f"Batch split failed: {exc}",
                    section_tolerance_pct=item.section_tolerance_pct,
                )
                for item in rubric
            ],
            LLMUsage(),
        )

    all_indexed_results = [indexed_result for split_result in split_results for indexed_result in split_result]

    # Sort back to original rubric order
    all_indexed_results.sort(key=lambda x: x[0])
    results = [r for _, r in all_indexed_results]

    return results, total_usage


def get_errored_indices(results: list[CriterionResult]) -> list[int]:
    """Return indices of criteria where met is None (infrastructure error)."""
    return [i for i, r in enumerate(results) if r.met is None and not r.skipped]


def apply_retries(
    results: list[CriterionResult],
    retry_results: list[CriterionResult],
    errored_indices: list[int],
) -> list[CriterionResult]:
    """Return a new results list with retry outcomes spliced in at *errored_indices*."""
    retry_map = dict(zip(errored_indices, retry_results, strict=False))
    return [retry_map.get(i, r) for i, r in enumerate(results)]


def _available_sections(rubric: list[RubricItem]) -> list[str]:
    """Return section names in first-seen rubric order."""
    sections: list[str] = []
    for item in rubric:
        if item.section is not None and item.section not in sections:
            sections.append(item.section)
    return sections


def filter_rubric_sections(
    rubric: list[RubricItem],
    include_sections: list[str] | None = None,
    exclude_sections: list[str] | None = None,
) -> list[RubricItem]:
    """Filter a normalised rubric by exact section names."""
    include_sections = include_sections or []
    exclude_sections = exclude_sections or []
    if not include_sections and not exclude_sections:
        return rubric

    available_sections = _available_sections(rubric)
    available_set = set(available_sections)
    requested_sections = include_sections + exclude_sections
    unknown_sections = sorted({section for section in requested_sections if section not in available_set})
    if unknown_sections:
        available = ", ".join(available_sections) if available_sections else "(none)"
        unknown = ", ".join(unknown_sections)
        msg = f"Unknown section filter value(s): {unknown}. Available sections: {available}"
        raise ValueError(msg)

    include_set = set(include_sections)
    exclude_set = set(exclude_sections)
    filtered = [
        item
        for item in rubric
        if (
            (item.section in include_set if include_sections else True)
            and not (item.section is not None and item.section in exclude_set)
        )
    ]
    if not filtered:
        msg = "Section filters left no rubric criteria to evaluate."
        raise ValueError(msg)
    return filtered


def collect_section_gates(rubric: list[RubricItem]) -> dict[str, list[str]]:
    """Return ordered section-level gates keyed by section name."""
    section_gates: dict[str, list[str]] = {}
    for item in rubric:
        if item.section is None:
            continue
        item_gates: list[str] = []
        if item.section_gate is not None:
            item_gates.append(item.section_gate)
        item_gates.extend(item.section_gates)
        if not item_gates:
            continue
        gates = section_gates.setdefault(item.section, [])
        for gate in item_gates:
            if gate not in gates:
                gates.append(gate)
    return section_gates


@dataclass(frozen=True)
class _SectionGateRequest:
    """One section-level gate to evaluate."""

    section: str
    index: int
    criterion: str


def evaluate_section_gate_batch(
    config: GraderConfig,
    requests: list[_SectionGateRequest],
    final_output: str,
    instructions: str,
    judge_guidance: str,
    judge_prompt: str | None,
    trace_suffix: str = "",
) -> tuple[dict[tuple[str, int], SectionGateResult], LLMUsage]:
    """Evaluate one batch of section-level gates."""
    if not requests:
        return {}, LLMUsage()

    n = len(requests)
    splits = split_count_for_auto_parallel(config, n)
    chunks = chunk_indexed_items(requests, splits)

    print(f"[section-gates] Evaluating {n} section gate(s) in {len(chunks)} chunk(s)...")  # noqa: T201

    def _run_split(
        split_idx: int,
        chunk: list[tuple[int, _SectionGateRequest]],
    ) -> _AdaptiveRunResult[dict[tuple[str, int], SectionGateResult]]:
        criteria = [request.criterion for _orig_idx, request in chunk]
        batch_timeout = batch_timeout_for_count(config, len(criteria))
        print(  # noqa: T201
            f"  [section gate split {split_idx + 1}/{len(chunks)}] "
            f"{len(criteria)} gate(s) (timeout={batch_timeout}s)..."
        )
        judge_input = BatchJudgeInput(
            model=config.model,
            reasoning_effort=config.reasoning_effort,
            instructions=instructions,
            final_output=final_output,
            criteria=criteria,
            workdir=config.workdir,
            mcp_servers=config.mcp_servers,
            excel_backend=config.excel_backend,
            judge_guidance=judge_guidance,
            judge_prompt=judge_prompt,
        )
        split_suffix = trace_suffix if len(chunks) == 1 else f"{trace_suffix}_split{split_idx}"
        trace_path = os.path.join(config.output_dir, f"judge_trace_section_gates{split_suffix}.txt")
        verdicts, usage = run_judge(
            judge_input,
            sandbox_user=config.sandbox_user,
            trace_path=trace_path,
            timeout=batch_timeout,
            clone_parent_dir=judge_clone_parent_dir(config),
        )

        results: dict[tuple[str, int], SectionGateResult] = {}
        for local_idx, (_orig_idx, request) in enumerate(chunk):
            verdict = (
                verdicts[local_idx]
                if local_idx < len(verdicts)
                else Verdict(met=None, reasoning="No reasoning provided.")
            )
            results[(request.section, request.index)] = SectionGateResult(
                section=request.section,
                index=request.index,
                criterion=request.criterion,
                met=verdict.met,
                reasoning=verdict.reasoning,
                evidence=verdict.evidence,
            )
            print(  # noqa: T201
                f"  [section gate {local_idx + 1}/{len(chunk)}] "
                f"{request.section}[{request.index}]: {format_status(met=verdict.met)}"
            )

        return _AdaptiveRunResult(
            value=results,
            usage=usage,
            rate_limited=verdicts_hit_rate_limit(verdicts) or section_gate_results_hit_rate_limit(list(results.values())),
        )

    def _make_section_gate_work_item(
        split_idx: int,
        chunk: list[tuple[int, _SectionGateRequest]],
    ) -> _AdaptiveWorkItem[dict[tuple[str, int], SectionGateResult]]:
        def _run_bound_split() -> _AdaptiveRunResult[dict[tuple[str, int], SectionGateResult]]:
            return _run_split(split_idx, chunk)

        return _AdaptiveWorkItem(index=split_idx, label=f"section gate split {split_idx}", run=_run_bound_split)

    work_items = [_make_section_gate_work_item(split_idx, chunk) for split_idx, chunk in enumerate(chunks)]
    try:
        split_results, total_usage = execute_adaptive_work(
            config,
            work_items,
            label="section-gates",
            fixed_concurrency=len(chunks),
            discovery_default_cap=config.rate_limit_max_concurrency,
        )
    except Exception as exc:  # noqa: BLE001
        msg = f"Section gate batch failed: {exc}"
        return (
            {
                (request.section, request.index): SectionGateResult(
                    section=request.section,
                    index=request.index,
                    criterion=request.criterion,
                    met=None,
                    reasoning=msg,
                )
                for request in requests
            },
            LLMUsage(),
        )

    merged_results: dict[tuple[str, int], SectionGateResult] = {}
    for split_result in split_results:
        merged_results.update(split_result)
    return merged_results, total_usage


def skipped_section_gate_result(
    section: str,
    index: int,
    criterion: str,
    reason: str,
) -> SectionGateResult:
    """Build a skipped SectionGateResult for gates bypassed by an earlier gate."""
    return SectionGateResult(
        section=section,
        index=index,
        criterion=criterion,
        met=None,
        reasoning=reason,
        skipped=True,
        skip_reason=reason,
    )


def evaluate_section_gates(
    config: GraderConfig,
    section_gates: dict[str, list[str]],
    final_output: str,
    instructions: str,
    judge_guidance: str,
    judge_prompt: str | None,
    trace_suffix: str = "",
) -> tuple[dict[str, list[SectionGateResult]], LLMUsage, int]:
    """Evaluate ordered section-level gates with per-section short-circuiting."""
    section_gate_results: dict[str, list[SectionGateResult]] = {section: [] for section in section_gates}
    if not section_gates:
        return section_gate_results, LLMUsage(), 0

    active_gate_indices = {section: 0 for section in section_gates}
    total_usage = LLMUsage()
    initial_errored = 0
    round_index = 0

    while active_gate_indices:
        requests = [
            _SectionGateRequest(section=section, index=index, criterion=section_gates[section][index])
            for section, index in active_gate_indices.items()
        ]
        round_results, usage = evaluate_section_gate_batch(
            config,
            requests,
            final_output,
            instructions,
            judge_guidance,
            judge_prompt,
            trace_suffix=f"_round{round_index}{trace_suffix}",
        )
        total_usage = total_usage + usage
        initial_errored += sum(1 for result in round_results.values() if result.met is None)

        for attempt in range(config.judge_retries):
            errored_requests = [
                request
                for request in requests
                if round_results[(request.section, request.index)].met is None
            ]
            if not errored_requests:
                break
            print(  # noqa: T201
                f"\n[retry {attempt + 1}/{config.judge_retries}] "
                f"Retrying {len(errored_requests)} errored section gate(s)..."
            )
            retry_results, retry_usage = evaluate_section_gate_batch(
                config,
                errored_requests,
                final_output,
                instructions,
                judge_guidance,
                judge_prompt,
                trace_suffix=f"_round{round_index}_retry{attempt + 1}{trace_suffix}",
            )
            round_results.update(retry_results)
            total_usage = total_usage + retry_usage

        next_active_gate_indices: dict[str, int] = {}
        for request in requests:
            result = round_results[(request.section, request.index)]
            section_gate_results[request.section].append(result)
            next_index = request.index + 1
            if result.met is True and next_index < len(section_gates[request.section]):
                next_active_gate_indices[request.section] = next_index
            elif result.met is not True:
                if result.met is False:
                    reason = f"Skipped because section gate {request.index} for {request.section!r} was not met."
                else:
                    reason = (
                        f"Skipped because section gate {request.index} for {request.section!r} "
                        "could not be evaluated."
                    )
                for skipped_index in range(next_index, len(section_gates[request.section])):
                    section_gate_results[request.section].append(
                        skipped_section_gate_result(
                            section=request.section,
                            index=skipped_index,
                            criterion=section_gates[request.section][skipped_index],
                            reason=reason,
                        )
                    )

        active_gate_indices = next_active_gate_indices
        round_index += 1

    return section_gate_results, total_usage, initial_errored


def get_errored_section_gate_results(
    section_gate_results: dict[str, list[SectionGateResult]],
) -> list[SectionGateResult]:
    """Return attempted section gates that still failed due to judge error."""
    return [
        result
        for results in section_gate_results.values()
        for result in results
        if result.met is None and not result.skipped
    ]


def split_rubric_by_section_gates(
    rubric: list[RubricItem],
    section_gate_results: dict[str, list[SectionGateResult]],
) -> tuple[list[tuple[int, RubricItem]], dict[int, CriterionResult]]:
    """Return criteria to evaluate and skipped criterion results for failed section gates."""
    gated_out_sections = {
        section
        for section, results in section_gate_results.items()
        if results and not all(section_gate_result.met is True for section_gate_result in results)
    }
    indexed_rubric: list[tuple[int, RubricItem]] = []
    skipped_results: dict[int, CriterionResult] = {}

    for index, item in enumerate(rubric):
        if item.section in gated_out_sections:
            blocking_result = next(
                (
                    section_gate_result
                    for section_gate_result in section_gate_results[item.section]
                    if section_gate_result.met is not True and not section_gate_result.skipped
                ),
                None,
            )
            if blocking_result is not None and blocking_result.met is False:
                reason = f"Skipped because section gate for {item.section!r} was not met."
            else:
                reason = f"Skipped because section gate for {item.section!r} could not be evaluated."
            skipped_results[index] = skipped_result(item, reason)
        else:
            indexed_rubric.append((index, item))

    return indexed_rubric, skipped_results


def merge_indexed_results(
    rubric: list[RubricItem],
    indexed_rubric: list[tuple[int, RubricItem]],
    evaluated_results: list[CriterionResult],
    skipped_results: dict[int, CriterionResult],
) -> list[CriterionResult]:
    """Merge evaluated and skipped results back into original rubric order."""
    result_by_index = dict(skipped_results)
    for (index, _item), result in zip(indexed_rubric, evaluated_results, strict=True):
        result_by_index[index] = result
    return [result_by_index[i] for i in range(len(rubric))]


def split_indexed_rubric_for_gates_only(
    indexed_rubric: list[tuple[int, RubricItem]],
) -> tuple[list[tuple[int, RubricItem]], dict[int, CriterionResult]]:
    """Return retained gate criteria and skipped non-gate results for gate-only runs."""
    gate_indexed_rubric: list[tuple[int, RubricItem]] = []
    skipped_results: dict[int, CriterionResult] = {}
    for index, item in indexed_rubric:
        if item.gate:
            gate_indexed_rubric.append((index, item))
        else:
            skipped_results[index] = skipped_result(item, GATES_ONLY_SKIP_REASON)
    return gate_indexed_rubric, skipped_results


def count_gate_check_statuses(
    results: list[CriterionResult],
    section_gate_results: dict[str, list[SectionGateResult]],
) -> tuple[int, int, int]:
    """Return pass/fail/error counts for attempted section gates and gate criteria."""
    passed = 0
    failed = 0
    errored = 0

    for section_results in section_gate_results.values():
        for result in section_results:
            if result.skipped:
                continue
            if result.met is True:
                passed += 1
            elif result.met is False:
                failed += 1
            else:
                errored += 1

    for criterion_result in results:
        if not criterion_result.gate or criterion_result.skipped:
            continue
        if criterion_result.met is True:
            passed += 1
        elif criterion_result.met is False:
            failed += 1
        else:
            errored += 1

    return passed, failed, errored


@dataclass
class _SectionScoreAccumulator:
    """Mutable scoring state for one named rubric section."""

    section: str
    score: float = 0.0
    minimum_score: float = 0.0
    maximum_score: float = 0.0
    gate_count: int = 0
    passed_gate_indices: list[int] = dataclass_field(default_factory=list)
    failed_gate_indices: list[int] = dataclass_field(default_factory=list)
    section_gate_results: list[SectionGateResult] = dataclass_field(default_factory=list)
    section_gate_gated_out: bool = False
    gated_out: bool = False
    section_tolerance_pct: float | None = None

    def to_result(self) -> SectionResult:
        single_section_gate = self.section_gate_results[0] if len(self.section_gate_results) == 1 else None
        passed_section_gate_indices = [
            result.index for result in self.section_gate_results if result.met is True and not result.skipped
        ]
        failed_section_gate_indices = [
            result.index for result in self.section_gate_results if result.met is False and not result.skipped
        ]
        errored_section_gate_indices = [
            result.index for result in self.section_gate_results if result.met is None and not result.skipped
        ]
        skipped_section_gate_indices = [result.index for result in self.section_gate_results if result.skipped]
        if not self.section_gate_results:
            section_gates_met = None
        elif errored_section_gate_indices:
            section_gates_met = None
        else:
            section_gates_met = all(result.met is True for result in self.section_gate_results)

        return SectionResult(
            section=self.section,
            score=round(self.score, 4),
            minimum_score=round(self.minimum_score, 4),
            maximum_score=round(self.maximum_score, 4),
            gate_count=self.gate_count,
            passed_gate_indices=self.passed_gate_indices,
            failed_gate_indices=self.failed_gate_indices,
            section_gate=single_section_gate.criterion if single_section_gate is not None else None,
            section_gate_met=single_section_gate.met if single_section_gate is not None else None,
            section_gate_reasoning=single_section_gate.reasoning if single_section_gate is not None else None,
            section_gate_evidence=single_section_gate.evidence if single_section_gate is not None else [],
            section_gate_count=len(self.section_gate_results),
            section_gate_results=self.section_gate_results,
            passed_section_gate_indices=passed_section_gate_indices,
            failed_section_gate_indices=failed_section_gate_indices,
            errored_section_gate_indices=errored_section_gate_indices,
            skipped_section_gate_indices=skipped_section_gate_indices,
            section_gates_met=section_gates_met,
            section_tolerance_pct=self.section_tolerance_pct,
            gated_out=self.gated_out,
        )


def score_results(
    results: list[CriterionResult],
    section_gate_results: dict[str, list[SectionGateResult]] | None = None,
) -> tuple[list[CriterionResult], list[SectionResult]]:
    """Apply section gate scoring and return scored criteria plus section summaries."""
    section_gate_results = section_gate_results or {}
    sections: dict[str, _SectionScoreAccumulator] = {}

    for index, result in enumerate(results):
        if result.section is None:
            continue
        section = sections.setdefault(result.section, _SectionScoreAccumulator(section=result.section))
        if section.section_tolerance_pct is None and result.section_tolerance_pct is not None:
            section.section_tolerance_pct = result.section_tolerance_pct
        if result.weight < 0:
            section.minimum_score += result.weight
        elif result.weight > 0:
            section.maximum_score += result.weight

        if result.gate and not result.skipped:
            section.gate_count += 1
            if result.met is True:
                section.passed_gate_indices.append(index)
            else:
                section.failed_gate_indices.append(index)

    for section_name, gate_results in section_gate_results.items():
        section = sections.setdefault(section_name, _SectionScoreAccumulator(section=section_name))
        section.section_gate_results = gate_results
        section.section_gate_gated_out = bool(gate_results) and not all(
            section_gate_result.met is True for section_gate_result in gate_results
        )

    for section in sections.values():
        section.gated_out = section.section_gate_gated_out or bool(section.failed_gate_indices)

    scored_results: list[CriterionResult] = []
    for result in results:
        contribution = 0.0
        if result.met is True and not result.skipped:
            section_gated_out = result.section is not None and sections[result.section].gated_out
            if not section_gated_out:
                contribution = result.weight

        scored = result.model_copy(update={"score_contribution": contribution})
        scored_results.append(scored)
        if scored.section is not None:
            sections[scored.section].score += contribution

    section_results = [section.to_result() for section in sections.values()]
    return scored_results, section_results


def excel_server_config(config: GraderConfig, *, workdir: Path | None = None) -> ExcelServerConfig:
    """Return the runtime Excel service config for *config*."""
    return ExcelServerConfig(
        workdir=workdir or Path(config.workdir),
        backend=config.excel_backend.backend,
        windows_url=config.excel_backend.windows_url,
        auth_token=config.excel_backend.auth_token,
        timeout_seconds=config.excel_backend.timeout_seconds,
        visible=config.excel_backend.visible,
        allow_macros=config.excel_backend.allow_macros,
        max_cells_per_call=config.excel_backend.max_cells_per_call,
        max_format_cells_per_call=config.excel_backend.max_format_cells_per_call,
    )


def configured_repair_check_workbook(config: GraderConfig) -> str | None:
    """Return the configured or default workbook path used by repair/preflight checks."""
    return config.workbook_repair_check.workbook_path or default_repair_check_workbook(config.workdir)


def repair_check_service_path(config: GraderConfig, workbook_path: str) -> tuple[ExcelServerConfig, str, str | None]:
    """Return the Excel service config/path for repair checking *workbook_path*.

    Local Mac Excel shows a "Grant Access" powerbox on first access to any path,
    so the workbook is always staged under the stable ``excel_access`` root.
    Granting that one folder once (see ``run_excel_preflight``) covers every
    per-run subdirectory, so later runs do not prompt. The returned cleanup
    path, when set, is a directory the caller removes.
    """
    if config.excel_backend.backend != "mac_excel":
        return excel_server_config(config), workbook_path, None

    original = Path(workbook_path)
    original_path = original if original.is_absolute() else Path(config.workdir) / original
    # Always stage under the stable excel_access root (even for non-protected
    # workdirs): per-run paths are unique, so a one-time folder grant only
    # persists when Excel opens from this single granted parent.
    staging_root = ensure_work_root(excel_access_root()) / "repair"
    staging_root.mkdir(parents=True, exist_ok=True)
    temp_dir = tempfile.mkdtemp(prefix="repair_", dir=str(staging_root))
    try:
        temp_path = Path(temp_dir) / original_path.name
        shutil.copy2(original_path, temp_path)
    except BaseException:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise
    return excel_server_config(config, workdir=Path(temp_dir)), temp_path.name, temp_dir


def workbook_repair_check_result_from_response(
    config: GraderConfig,
    workbook_path: str,
    result: dict[str, object],
) -> WorkbookRepairCheckResult:
    """Convert an Excel service response to the public info.json model."""
    return WorkbookRepairCheckResult(
        enabled=True,
        checked=bool(result.get("checked", True)),
        workbook_path=workbook_path,
        backend=str(result.get("backend") or config.excel_backend.backend),
        opened=bool(result["opened"]) if "opened" in result else None,
        repair_dialog_detected=(
            bool(result["repair_dialog_detected"]) if "repair_dialog_detected" in result else None
        ),
        details={
            key: value
            for key, value in result.items()
            if key
            not in {
                "path",
                "backend",
                "checked",
                "opened",
                "repair_dialog_detected",
                "skipped_reason",
                "error",
            }
        },
        skipped_reason=str(result["skipped_reason"]) if "skipped_reason" in result else None,
        error=str(result["error"]) if "error" in result else None,
    )


def run_workbook_repair_check(config: GraderConfig) -> WorkbookRepairCheckResult | None:
    """Run the deterministic Excel repair check for the submitted workbook."""
    repair_config = config.workbook_repair_check
    if not repair_config.enabled:
        return WorkbookRepairCheckResult(
            enabled=False,
            checked=False,
            skipped_reason="disabled",
        )

    if not config.excel_backend.enabled:
        return WorkbookRepairCheckResult(
            enabled=True,
            checked=False,
            workbook_path=configured_repair_check_workbook(config),
            skipped_reason="excel_backend disabled",
        )

    workbook_path = configured_repair_check_workbook(config)
    if workbook_path is None:
        return WorkbookRepairCheckResult(
            enabled=True,
            checked=False,
            skipped_reason="submitted_output.xlsx not found",
        )

    temp_dir: str | None = None
    try:
        service_config, service_workbook_path, temp_dir = repair_check_service_path(config, workbook_path)
        service = create_service(service_config)
        result = service.workbook_repair_check(service_workbook_path)
    except Exception as e:  # noqa: BLE001
        return WorkbookRepairCheckResult(
            enabled=True,
            checked=False,
            workbook_path=workbook_path,
            backend=config.excel_backend.backend,
            repair_dialog_detected=None,
            error=str(e),
        )
    finally:
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)

    return workbook_repair_check_result_from_response(config, workbook_path, result)


def run_excel_preflight(config: GraderConfig) -> int:
    """Open the submitted workbook through the configured Excel backend."""
    if not config.excel_backend.enabled:
        print("Excel preflight skipped: excel_backend is disabled.")  # noqa: T201
        return 0

    workbook_path = configured_repair_check_workbook(config)
    if workbook_path is None:
        print(  # noqa: T201
            "Excel preflight failed: submitted_output.xlsx not found and no "
            "[workbook_repair_check].workbook_path is configured.",
            file=sys.stderr,
        )
        return 1

    if config.excel_backend.backend == "mac_excel":
        print(  # noqa: T201
            f"If macOS shows a 'Grant Access' dialog, navigate up and grant the folder "
            f"{excel_access_root()} once — that covers all future grading runs.",
        )

    temp_dir: str | None = None
    try:
        service_config, service_workbook_path, temp_dir = repair_check_service_path(config, workbook_path)
        service = create_service(service_config)
        result = service.workbook_summary(service_workbook_path, recalculate=False)
    except Exception as e:  # noqa: BLE001
        print(f"Excel preflight failed for {workbook_path}: {e}", file=sys.stderr)  # noqa: T201
        return 1
    finally:
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)

    backend = result.get("backend", config.excel_backend.backend)
    sheets = result.get("sheets")
    sheet_count = len(sheets) if isinstance(sheets, list) else "unknown"
    print(  # noqa: T201
        f"Excel preflight passed for {workbook_path} "
        f"(backend={backend}, sheets={sheet_count})."
    )
    return 0


def default_repair_check_workbook(workdir: str) -> str | None:
    """Return the default submitted workbook path if present."""
    candidate = Path(workdir) / "submitted_output.xlsx"
    return candidate.name if candidate.is_file() else None


def warn_if_excel_workspace_protected(config: GraderConfig) -> None:
    """Warn when Mac Excel is enabled but run files live in a sandbox-protected folder.

    The repair check stages protected workbooks under the work root, but the
    judge's Excel MCP opens clones under ``output_dir``; if that is protected,
    Mac Excel can still show "Grant Access" prompts mid-run.
    """
    if not (config.excel_backend.enabled and config.excel_backend.backend == "mac_excel"):
        return
    protected = [
        label
        for label, value in (("output_dir", config.output_dir), ("workdir", config.workdir))
        if is_protected_path(value)
    ]
    if not protected:
        return
    print(  # noqa: T201
        f"WARNING: Microsoft Excel is enabled and {', '.join(protected)} is in a macOS "
        "sandbox-protected folder (~/Documents, ~/Desktop, ~/Downloads, iCloud, or a "
        "cloud-sync folder). Excel may show 'Grant Access' prompts during grading. Move the "
        "run under a non-protected path (e.g. ~/gandalf_runs), set [excel_backend] "
        "enabled = false, or use backend = 'windows_vm'.",
        file=sys.stderr,
    )


def write_info(
    config: GraderConfig,
    results: list[CriterionResult],
    llm_usage: LLMUsage,
    errored_criterion_count: int,
    section_gate_results: dict[str, list[SectionGateResult]] | None = None,
    errored_section_gate_count: int = 0,
    workbook_repair_check: WorkbookRepairCheckResult | None = None,
    excel_metrics: dict[str, Any] | None = None,
    batch_recovery: dict[str, Any] | None = None,
) -> tuple[float, float]:
    """Compute reward and raw score and write info.json. Returns (reward, raw_score).

    raw_score: sum of score contributions after section gates are applied.
    Errored criteria (met=None) contribute 0.

    reward: clip(0, 1, raw_score / sum_of_positive_weights), always in [0, 1].
    """
    scored_results, section_results = score_results(results, section_gate_results)

    raw_score = round(
        sum(r.score_contribution for r in scored_results),
        4,
    )

    minimum_score = round(sum(r.weight for r in scored_results if r.weight < 0), 4)
    maximum_score = round(sum(r.weight for r in scored_results if r.weight > 0), 4)

    unpenalized_reward = round(
        max(0.0, min(1.0, raw_score / maximum_score)) if maximum_score > 0 else 0.0,
        4,
    )
    repair_penalty_applied = (
        workbook_repair_check is not None
        and workbook_repair_check.repair_dialog_detected is True
        and config.workbook_repair_check.penalty_multiplier < 1.0
    )
    reward = (
        round(unpenalized_reward * config.workbook_repair_check.penalty_multiplier, 4)
        if repair_penalty_applied
        else unpenalized_reward
    )

    n_total = len(results)
    n_skipped = sum(1 for result in results if result.skipped)
    n_evaluated = n_total - errored_criterion_count - n_skipped
    evaluated_pct = round((n_evaluated / n_total * 100.0) if n_total > 0 else 100.0, 2)

    info = EvaluationInfo(
        reward=reward,
        raw_score=raw_score,
        unpenalized_reward=unpenalized_reward if repair_penalty_applied else None,
        repair_penalty_applied=repair_penalty_applied,
        repair_penalty_multiplier=(
            config.workbook_repair_check.penalty_multiplier if repair_penalty_applied else None
        ),
        workbook_repair_check=workbook_repair_check,
        minimum_score=minimum_score,
        maximum_score=maximum_score,
        criterion_results=scored_results,
        section_results=section_results,
        llm_usage=llm_usage,
        excel_metrics=excel_metrics,
        batch_recovery=batch_recovery,
        errored_criterion_count=errored_criterion_count,
        errored_section_gate_count=errored_section_gate_count,
        evaluated_criteria_pct=evaluated_pct,
    )
    info_data = info.model_dump(mode="json")
    strip_tolerance = config.golden_check
    for result in info_data["criterion_results"]:
        if strip_tolerance or result.get("section_tolerance_pct") is None:
            result.pop("section_tolerance_pct", None)
    for section in info_data["section_results"]:
        if strip_tolerance or section.get("section_tolerance_pct") is None:
            section.pop("section_tolerance_pct", None)
    if workbook_repair_check is None:
        info_data.pop("workbook_repair_check", None)
        info_data.pop("unpenalized_reward", None)
        info_data.pop("repair_penalty_applied", None)
        info_data.pop("repair_penalty_multiplier", None)
    elif not repair_penalty_applied:
        info_data.pop("unpenalized_reward", None)
        info_data.pop("repair_penalty_multiplier", None)
    if excel_metrics is None:
        info_data.pop("excel_metrics", None)
    if batch_recovery is None:
        info_data.pop("batch_recovery", None)
    with open(os.path.join(config.output_dir, "info.json"), "w") as f:
        json.dump(info_data, f, indent=2)

    return reward, raw_score


def remove_reward_file(output_dir: str) -> None:
    """Remove a stale reward.json when a run intentionally does not emit one."""
    with contextlib.suppress(FileNotFoundError):
        os.remove(os.path.join(output_dir, "reward.json"))


RunFn = Callable[..., tuple[list[CriterionResult], LLMUsage]]


@dataclass
class BatchRecoveryTracker:
    """Aggregate observability for batch-fragment recovery."""

    attempted_mini_batches: int = 0
    attempted_criteria: int = 0
    recovered_criteria: int = 0
    unresolved_criteria: int = 0
    rounds: int = 0
    extra_usage: LLMUsage = dataclass_field(default_factory=LLMUsage)

    def record_round(
        self,
        *,
        attempted_mini_batches: int,
        attempted_criteria: int,
        recovered_criteria: int,
        unresolved_criteria: int,
        usage: LLMUsage,
    ) -> None:
        self.rounds += 1
        self.attempted_mini_batches += attempted_mini_batches
        self.attempted_criteria += attempted_criteria
        self.recovered_criteria += recovered_criteria
        self.unresolved_criteria = unresolved_criteria
        self.extra_usage = self.extra_usage + usage

    def to_info(self) -> dict[str, Any] | None:
        """Return an info.json payload, or None if recovery never ran."""
        if self.attempted_mini_batches == 0:
            return None
        return {
            "attempted_mini_batches": self.attempted_mini_batches,
            "attempted_criteria": self.attempted_criteria,
            "recovered_criteria": self.recovered_criteria,
            "unresolved_criteria": self.unresolved_criteria,
            "rounds": self.rounds,
            "extra_usage": self.extra_usage.model_dump(mode="json"),
        }


def is_recoverable_batch_error_text(text: str) -> bool:
    """Return whether an unresolved verdict looks like recoverable batch fragmentation."""
    lowered = text.lower()
    if is_rate_limit_error_text(lowered):
        return False
    return any(marker in lowered for marker in BATCH_RECOVERY_ERROR_MARKERS)


def get_batch_recovery_indices(results: list[CriterionResult]) -> list[int]:
    """Return criterion indices eligible for batch-fragment recovery."""
    return [
        i
        for i, result in enumerate(results)
        if result.met is None and not result.skipped and is_recoverable_batch_error_text(result.reasoning)
    ]


def chunk_plain(items: list[T], size: int) -> list[list[T]]:
    """Return fixed-size chunks preserving order."""
    return [items[start : start + size] for start in range(0, len(items), size)]


def run_batch_recovery_chunks(
    config: GraderConfig,
    *,
    rubric: list[RubricItem],
    errored_indices: list[int],
    final_output: str,
    instructions: str,
    judge_guidance: str,
    judge_prompt: str | None,
    recovery_round: int,
    trace_suffix: str,
) -> tuple[list[tuple[int, CriterionResult]], LLMUsage, int]:
    """Rerun recoverable batch-fragment errors in small sequential mini-batches."""
    total_usage = LLMUsage()
    indexed_results: list[tuple[int, CriterionResult]] = []
    chunks = chunk_plain(errored_indices, config.batch_recovery_target_chunk_size)
    for chunk_idx, chunk in enumerate(chunks):
        criteria = [criterion_for_judge(rubric[i], golden_check=config.golden_check) for i in chunk]
        batch_timeout = batch_timeout_for_count(config, len(criteria))
        print(  # noqa: T201
            f"[batch-recovery] round {recovery_round}, chunk {chunk_idx + 1}/{len(chunks)}: "
            f"{len(criteria)} criteria (timeout={batch_timeout}s)..."
        )
        judge_input = BatchJudgeInput(
            model=config.model,
            reasoning_effort=config.reasoning_effort,
            instructions=instructions,
            final_output=final_output,
            criteria=criteria,
            workdir=config.workdir,
            mcp_servers=config.mcp_servers,
            excel_backend=config.excel_backend,
            judge_guidance=judge_guidance,
            judge_prompt=judge_prompt,
        )
        trace_path = os.path.join(
            config.output_dir,
            f"judge_trace_batch_recovery{recovery_round}_{chunk_idx}{trace_suffix}.txt",
        )
        verdicts, usage = run_judge(
            judge_input,
            sandbox_user=config.sandbox_user,
            trace_path=trace_path,
            timeout=batch_timeout,
            clone_parent_dir=judge_clone_parent_dir(config),
        )
        total_usage = total_usage + usage
        for local_idx, orig_idx in enumerate(chunk):
            verdict = (
                verdicts[local_idx]
                if local_idx < len(verdicts)
                else Verdict(met=None, reasoning="No reasoning provided.")
            )
            indexed_results.append((orig_idx, verdict_to_result(rubric[orig_idx], verdict)))
    return indexed_results, total_usage, len(chunks)


def evaluate_with_retries(
    config: GraderConfig,
    *,
    run: RunFn,
    rubric: list[RubricItem],
    indexed_rubric: list[tuple[int, RubricItem]],
    skipped_results: dict[int, CriterionResult],
    final_output: str,
    instructions: str,
    judge_guidance: str,
    judge_prompt: str | None,
    llm_usage: LLMUsage,
    retry_label: str = "errored criteria",
    batch_recovery: BatchRecoveryTracker | None = None,
) -> tuple[list[CriterionResult], LLMUsage, int]:
    """Evaluate *indexed_rubric* via *run*, then retry errored criteria.

    Returns the merged results (in full rubric order, including *skipped_results*),
    the accumulated LLM usage, and the pre-retry errored-criterion count (for
    observability). Retries reuse the same scheduler path as the initial run.
    """
    rubric_to_evaluate = [item for _index, item in indexed_rubric]
    if rubric_to_evaluate:
        evaluated_results, criterion_usage = run(
            config,
            rubric_to_evaluate,
            final_output,
            instructions,
            judge_guidance,
            judge_prompt,
        )
        llm_usage = llm_usage + criterion_usage
    else:
        evaluated_results = []
    results = merge_indexed_results(rubric, indexed_rubric, evaluated_results, skipped_results)

    initial_errored = len(get_errored_indices(results))

    remaining_retry_budget = config.judge_retries
    if config.mode == "batch" and config.batch_recovery_enabled and batch_recovery is not None:
        recovery_round_limit = min(config.batch_recovery_max_rounds, remaining_retry_budget)
        for recovery_round in range(1, recovery_round_limit + 1):
            recoverable_indices = get_batch_recovery_indices(results)
            if not recoverable_indices:
                break
            recovery_results, recovery_usage, attempted_mini_batches = run_batch_recovery_chunks(
                config,
                rubric=rubric,
                errored_indices=recoverable_indices,
                final_output=final_output,
                instructions=instructions,
                judge_guidance=judge_guidance,
                judge_prompt=judge_prompt,
                recovery_round=recovery_round,
                trace_suffix="",
            )
            recovery_map = dict(recovery_results)
            recovered_count = sum(
                1 for idx in recoverable_indices if recovery_map.get(idx, results[idx]).met is not None
            )
            for idx, result in recovery_results:
                results[idx] = result
            llm_usage = llm_usage + recovery_usage
            batch_recovery.record_round(
                attempted_mini_batches=attempted_mini_batches,
                attempted_criteria=len(recoverable_indices),
                recovered_criteria=recovered_count,
                unresolved_criteria=len(get_errored_indices(results)),
                usage=recovery_usage,
            )
            remaining_retry_budget -= 1

    for attempt in range(remaining_retry_budget):
        errored = get_errored_indices(results)
        if not errored:
            break
        print(  # noqa: T201
            f"\n[retry {attempt + 1}/{remaining_retry_budget}] Retrying {len(errored)} {retry_label}..."
        )
        retry_rubric = [rubric[i] for i in errored]
        retry_results, retry_usage = run(
            config,
            retry_rubric,
            final_output,
            instructions,
            judge_guidance,
            judge_prompt,
            trace_suffix=f"_retry{attempt + 1}",
        )
        results = apply_retries(results, retry_results, errored)
        llm_usage = llm_usage + retry_usage

    return results, llm_usage, initial_errored


def main() -> None:
    parser = argparse.ArgumentParser(description="Grader: evaluate agent output via agent-as-judge")
    parser.add_argument("--config", required=True, help="Path to grader config TOML file")
    parser.add_argument(
        "--golden-check",
        action="store_true",
        help="Disable configured section-level tolerance bands for exact golden-output checks",
    )
    parser.add_argument(
        "--include-section",
        action="append",
        default=[],
        help="Exact rubric section name to include. Repeat to include multiple sections.",
    )
    parser.add_argument(
        "--exclude-section",
        action="append",
        default=[],
        help="Exact rubric section name to exclude. Repeat to exclude multiple sections.",
    )
    parser.add_argument(
        "--gates-only",
        action="store_true",
        help="Evaluate only section-level gates and gate criteria, then exit non-zero on any gate failure.",
    )
    parser.add_argument(
        "--excel-preflight",
        action="store_true",
        help="Open the submitted workbook through the configured Excel backend, then exit without grading.",
    )
    parser.add_argument(
        "--excel-backend",
        choices=[
            "config",
            "mac-excel",
            "mac_excel",
            "windows-vm",
            "windows_vm",
            "excel-vm",
            "excel_vm",
            "libreoffice",
            "libre-office",
            "disabled",
            "none",
        ],
        help=(
            "Override [excel_backend] for this run. Use 'windows-vm' for the Excel VM, "
            "'mac-excel' for local Mac Excel, 'libreoffice' for LibreOffice-backed MCP tools, "
            "or 'disabled'/'none' to disable workbook MCP tools."
        ),
    )
    parser.add_argument(
        "--excel-windows-url",
        help="Windows Excel worker URL to use with --excel-backend windows-vm.",
    )
    parser.add_argument(
        "--excel-auth-token",
        help="Auth token for the Windows Excel worker.",
    )
    args = parser.parse_args()

    if args.excel_preflight:
        incompatible: list[str] = []
        if args.gates_only:
            incompatible.append("--gates-only")
        if args.include_section:
            incompatible.append("--include-section")
        if args.exclude_section:
            incompatible.append("--exclude-section")
        if incompatible:
            parser.error(f"--excel-preflight cannot be combined with {', '.join(incompatible)}")

    config = load_config(args.config)
    if args.golden_check:
        config.golden_check = True
    try:
        config = apply_excel_backend_cli_override(
            config,
            backend=cast("CliExcelBackend | None", args.excel_backend),
            windows_url=args.excel_windows_url,
            auth_token=args.excel_auth_token,
        )
    except ValueError as e:
        parser.error(str(e))
    config = apply_golden_check_defaults(config)

    if args.excel_preflight:
        sys.exit(run_excel_preflight(config))

    instructions = resolve_instructions(config)

    # The model validator guarantees exactly one of rubric / rubric_path is set.
    if config.rubric is not None:
        rubric = cast("list[RubricItem]", config.rubric)
    else:
        assert config.rubric_path is not None  # noqa: S101  # guaranteed by model validator
        rubric = load_rubric(config.rubric_path)
    try:
        rubric = filter_rubric_sections(
            rubric,
            include_sections=args.include_section,
            exclude_sections=args.exclude_section,
        )
    except ValueError as e:
        parser.error(str(e))
    final_output = load_trajectory_final_output(config.trajectory_path)
    try:
        digest_text = workbook_digest_text(config)
    except ValueError as e:
        parser.error(str(e))
    judge_guidance = effective_judge_guidance(
        resolve_judge_guidance(config),
        golden_check=config.golden_check,
        excel_backend=config.excel_backend,
        workbook_digest=digest_text,
    )
    judge_prompt = resolve_judge_prompt(config)

    os.makedirs(config.output_dir, exist_ok=True)
    clear_excel_metrics(config.output_dir)
    warn_if_excel_workspace_protected(config)
    batch_recovery = BatchRecoveryTracker()

    with libreoffice_cache_environment(config):
        if config.mode == "batch":
            run = run_batch_concurrent if config.batch_splits is not None or config.auto_parallel else run_batch
        else:
            run = run_individual

        # 1. Evaluate section-level gates before section criteria, so failed gates
        # can skip the expensive section-level criterion checks.
        section_gates = collect_section_gates(rubric)
        section_gate_results, llm_usage, initial_gate_errored = evaluate_section_gates(
            config,
            section_gates,
            final_output=final_output,
            instructions=instructions,
            judge_guidance=judge_guidance,
            judge_prompt=judge_prompt,
        )

        indexed_rubric, skipped_results = split_rubric_by_section_gates(rubric, section_gate_results)

        if args.gates_only:
            gate_indexed_rubric, gates_only_skipped_results = split_indexed_rubric_for_gates_only(indexed_rubric)
            skipped_results = {**skipped_results, **gates_only_skipped_results}

            results, llm_usage, _initial_errored = evaluate_with_retries(
                config,
                run=run,
                rubric=rubric,
                indexed_rubric=gate_indexed_rubric,
                skipped_results=skipped_results,
                final_output=final_output,
                instructions=instructions,
                judge_guidance=judge_guidance,
                judge_prompt=judge_prompt,
                llm_usage=llm_usage,
                retry_label="errored gate criteria",
                batch_recovery=batch_recovery,
            )

            final_errored = get_errored_indices(results)
            final_gate_errored = get_errored_section_gate_results(section_gate_results)
            errored_count = len(final_errored)
            gate_errored_count = len(final_gate_errored)
            write_info(
                config,
                results,
                llm_usage,
                errored_count,
                section_gate_results=section_gate_results,
                errored_section_gate_count=gate_errored_count,
                excel_metrics=collect_excel_metrics(config.output_dir),
                batch_recovery=batch_recovery.to_info(),
            )
            remove_reward_file(config.output_dir)

            passed_gates, failed_gates, errored_gates = count_gate_check_statuses(results, section_gate_results)
            total_gates = passed_gates + failed_gates + errored_gates
            if failed_gates or errored_gates:
                print(  # noqa: T201
                    f"\nGate check failed: {failed_gates} failed, {errored_gates} errored gate check(s).",
                    file=sys.stderr,
                )
                print(f"info.json written to {config.output_dir}/ (reward.json NOT written)", file=sys.stderr)  # noqa: T201
                sys.exit(1)

            if total_gates == 0:
                print("\nGate check passed: 0/0 gate checks passed. No gates to evaluate.")  # noqa: T201
            else:
                print(f"\nGate check passed: {passed_gates}/{total_gates} gate checks passed.")  # noqa: T201
            print(f"info.json written to {config.output_dir}/ (reward.json NOT written)")  # noqa: T201
            return

        # 2-4. Initial criterion evaluation + retry loop (initial_errored is the
        # pre-retry error count, kept for the recovered-after-retry summary).
        results, llm_usage, initial_errored = evaluate_with_retries(
            config,
            run=run,
            rubric=rubric,
            indexed_rubric=indexed_rubric,
            skipped_results=skipped_results,
            final_output=final_output,
            instructions=instructions,
            judge_guidance=judge_guidance,
            judge_prompt=judge_prompt,
            llm_usage=llm_usage,
            retry_label="errored criteria",
            batch_recovery=batch_recovery,
        )

    # 5. ALWAYS write info.json (even on hard fail)
    final_errored = get_errored_indices(results)
    final_gate_errored = get_errored_section_gate_results(section_gate_results)
    errored_count = len(final_errored)
    gate_errored_count = len(final_gate_errored)
    workbook_repair_check = run_workbook_repair_check(config)
    reward, raw_score = write_info(
        config,
        results,
        llm_usage,
        errored_count,
        section_gate_results=section_gate_results,
        errored_section_gate_count=gate_errored_count,
        workbook_repair_check=workbook_repair_check,
        excel_metrics=collect_excel_metrics(config.output_dir),
        batch_recovery=batch_recovery.to_info(),
    )

    # 6. If any criteria or section gates still errored: do NOT write reward.json, exit 1
    if final_errored or final_gate_errored:
        print(  # noqa: T201
            f"\nERROR: {errored_count} criteria and {gate_errored_count} section gate(s) "
            f"could not be evaluated (initial criterion errors: {initial_errored}, "
            f"initial section gate errors: {initial_gate_errored}; after retries: "
            f"{errored_count} criteria, {gate_errored_count} section gate(s)).",
            file=sys.stderr,
        )
        print(f"info.json written to {config.output_dir}/ (reward.json NOT written)", file=sys.stderr)  # noqa: T201
        sys.exit(1)

    # 7. All resolved — write reward.json
    with open(os.path.join(config.output_dir, "reward.json"), "w") as f:
        json.dump({"reward": reward}, f, indent=2)

    print(f"\nReward: {reward} (raw: {raw_score})")  # noqa: T201
    if llm_usage.cost_usd > 0:
        print(  # noqa: T201
            f"Grader LLM cost: ${llm_usage.cost_usd:.4f} "
            f"({len(rubric)} criteria, "
            f"{llm_usage.prompt_tokens} prompt + {llm_usage.completion_tokens} completion tokens)"
        )
    if config.mode == "batch" and config.batch_splits is not None:
        cap = config.max_concurrency or config.rate_limit_max_concurrency
        mode_display = f"batch (batch_splits={config.batch_splits}, max_concurrency_cap={cap})"
    elif config.mode == "batch" and config.auto_parallel:
        cap = config.max_concurrency or config.rate_limit_max_concurrency
        mode_display = (
            "batch "
            f"(auto_parallel_target_chunk_size={config.auto_parallel_target_chunk_size}, "
            f"max_concurrency_cap={cap})"
        )
    elif config.max_concurrency is not None and config.max_concurrency > 1:
        mode_display = f"{config.mode} (max_concurrency={config.max_concurrency})"
    else:
        mode_display = config.mode
    print(f"Mode: {mode_display}")  # noqa: T201
    if initial_errored > 0:
        print(f"Retried: {initial_errored} criteria recovered after retry")  # noqa: T201
    print(f"Results written to {config.output_dir}/")  # noqa: T201
