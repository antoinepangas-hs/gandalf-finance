"""Data models and configuration loaders for the grader."""

import math
import tomllib
from pathlib import Path
from typing import Any, Literal

from pydantic import (
    BaseModel,
    Field,
    TypeAdapter,
    ValidationInfo,
    field_validator,
    model_validator,
)

from gandalf.excel_defaults import (
    DEFAULT_EXCEL_BACKEND,
    DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL,
    DEFAULT_EXCEL_MAX_CELLS_PER_CALL,
    DEFAULT_EXCEL_TIMEOUT_SECONDS,
)


def _append_unique_strings(target: list[str], values: list[str]) -> None:
    """Append values that have not already appeared, preserving order."""
    for value in values:
        if value not in target:
            target.append(value)


def _validate_tolerance_pct(value: float | None, field_name: str) -> float | None:
    """Validate a percentage tolerance field."""
    if value is None:
        return None
    if not math.isfinite(value):
        msg = f"{field_name} must be finite"
        raise ValueError(msg)
    if value < 0:
        msg = f"{field_name} must be non-negative"
        raise ValueError(msg)
    return value


class MCPServer(BaseModel):
    """Configuration for an MCP server.

    Supports stdio (subprocess) and remote network transports
    (streamable-http, http, sse).  Stdio servers require ``command``;
    remote servers require ``url``.
    """

    name: str
    transport: Literal["stdio", "streamable-http", "http", "sse"] = "stdio"
    command: str | None = None
    args: list[str] = Field(default_factory=list)
    env: dict[str, str] = Field(default_factory=dict)
    url: str | None = None
    headers: dict[str, str] = Field(default_factory=dict)

    @model_validator(mode="after")
    def _check_transport_fields(self) -> "MCPServer":
        if self.transport == "stdio":
            if not self.command:
                msg = f"MCP server {self.name!r}: 'command' is required for stdio transport"
                raise ValueError(msg)
        elif not self.url:
            msg = f"MCP server {self.name!r}: 'url' is required for transport {self.transport!r}"
            raise ValueError(msg)
        return self


ExcelBackendName = Literal["mac_excel", "windows_vm", "libreoffice"]


class ExcelBackendConfig(BaseModel):
    """Workbook inspection service configuration."""

    enabled: bool = True
    backend: ExcelBackendName = DEFAULT_EXCEL_BACKEND
    windows_url: str | None = None
    auth_token: str | None = None
    timeout_seconds: int = Field(default=DEFAULT_EXCEL_TIMEOUT_SECONDS, gt=0)
    visible: bool = False
    allow_macros: bool = False
    max_cells_per_call: int = Field(default=DEFAULT_EXCEL_MAX_CELLS_PER_CALL, gt=0)
    max_format_cells_per_call: int = Field(default=DEFAULT_EXCEL_MAX_FORMAT_CELLS_PER_CALL, gt=0)

    @field_validator("windows_url")
    @classmethod
    def _normalise_windows_url(cls, value: str | None) -> str | None:
        if value is None:
            return None
        value = value.strip().rstrip("/")
        if not value:
            msg = "windows_url cannot be empty"
            raise ValueError(msg)
        return value

    @field_validator("auth_token")
    @classmethod
    def _normalise_auth_token(cls, value: str | None) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value:
            msg = "auth_token cannot be empty"
            raise ValueError(msg)
        return value

    @model_validator(mode="after")
    def _check_windows_url(self) -> "ExcelBackendConfig":
        if self.backend == "windows_vm" and self.windows_url is None:
            msg = "windows_url is required when excel_backend.backend is 'windows_vm'"
            raise ValueError(msg)
        return self


class WorkbookRepairCheckConfig(BaseModel):
    """Pre-scoring check for Excel repair-required workbooks."""

    enabled: bool = True
    workbook_path: str | None = None
    penalty_multiplier: float = Field(default=0.5, ge=0.0, le=1.0)

    @field_validator("workbook_path")
    @classmethod
    def _normalise_optional_path(cls, value: str | None) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value:
            msg = "workbook_path cannot be empty"
            raise ValueError(msg)
        return value


class WorkbookDigestConfig(BaseModel):
    """Optional precomputed workbook digest injected into the judge prompt.

    Off by default: enabling it changes the judge prompt, so validate grading
    agreement before relying on it broadly.
    """

    enabled: bool = False
    workbook_path: str | None = None
    max_sheets: int = Field(default=30, gt=0)
    max_formula_samples: int = Field(default=15, ge=0)
    max_preview_columns: int = Field(default=20, ge=0)
    max_named_ranges: int = Field(default=50, ge=0)

    @field_validator("workbook_path")
    @classmethod
    def _normalise_optional_path(cls, value: str | None) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value:
            msg = "workbook_path cannot be empty"
            raise ValueError(msg)
        return value


class RubricItem(BaseModel):
    """A single rubric item with an evaluation criterion and weight.

    Weight can be negative to penalise undesired outcomes.  The sign of the
    weight carries the semantics: positive means "reward when met", negative
    means "penalise when met".  Items may optionally belong to a rubric
    section, and may be marked as gates for that section.
    """

    criterion: str
    weight: float
    section: str | None = None
    gate: bool = False
    section_gate: str | None = None
    section_gates: list[str] = Field(default_factory=list)
    section_tolerance_pct: float | None = None

    @field_validator("section", "section_gate")
    @classmethod
    def _normalise_optional_string(cls, value: str | None, info: ValidationInfo) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value:
            msg = f"{info.field_name} cannot be empty"
            raise ValueError(msg)
        return value

    @field_validator("section_gates")
    @classmethod
    def _normalise_string_list(cls, values: list[str], info: ValidationInfo) -> list[str]:
        normalised: list[str] = []
        for value in values:
            value = value.strip()
            if not value:
                msg = f"{info.field_name} cannot contain empty strings"
                raise ValueError(msg)
            normalised.append(value)
        return normalised

    @field_validator("section_tolerance_pct")
    @classmethod
    def _validate_section_tolerance_pct(cls, value: float | None, info: ValidationInfo) -> float | None:
        return _validate_tolerance_pct(value, info.field_name or "section_tolerance_pct")


class RubricSection(BaseModel):
    """A named subsection in a nested rubric document."""

    name: str
    gate: str | None = None
    gates: list[str] = Field(default_factory=list)
    tolerance_pct: float | None = None
    criteria: list[RubricItem] = Field(default_factory=list)

    @field_validator("name", "gate")
    @classmethod
    def _normalise_string(cls, value: str | None, info: ValidationInfo) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value:
            msg = f"{info.field_name} cannot be empty"
            raise ValueError(msg)
        return value

    @field_validator("gates")
    @classmethod
    def _normalise_string_list(cls, values: list[str], info: ValidationInfo) -> list[str]:
        normalised: list[str] = []
        for value in values:
            value = value.strip()
            if not value:
                msg = f"{info.field_name} cannot contain empty strings"
                raise ValueError(msg)
            normalised.append(value)
        return normalised

    @field_validator("tolerance_pct")
    @classmethod
    def _validate_tolerance_pct(cls, value: float | None, info: ValidationInfo) -> float | None:
        return _validate_tolerance_pct(value, info.field_name or "tolerance_pct")


class RubricDocument(BaseModel):
    """Nested rubric JSON shape with top-level criteria and named sections."""

    criteria: list[RubricItem] = Field(default_factory=list)
    sections: list[RubricSection] = Field(default_factory=list)

    @model_validator(mode="after")
    def _check_unique_sections(self) -> "RubricDocument":
        seen: set[str] = set()
        for section in self.sections:
            if section.name in seen:
                msg = f"Duplicate rubric section {section.name!r}"
                raise ValueError(msg)
            seen.add(section.name)
        return self


RubricInput = list[RubricItem] | RubricDocument


def normalise_rubric(rubric: RubricInput) -> list[RubricItem]:
    """Normalise supported rubric shapes to a flat list of RubricItem objects."""
    section_gates: dict[str, list[str]] = {}
    section_tolerances: dict[str, float] = {}

    if isinstance(rubric, RubricDocument):
        items = [item.model_copy() for item in rubric.criteria]
        for section in rubric.sections:
            if section.tolerance_pct is not None:
                section_tolerances[section.name] = section.tolerance_pct
            section_gate_values: list[str] = []
            if section.gate is not None:
                section_gate_values.append(section.gate)
            section_gate_values.extend(section.gates)
            _append_unique_strings(section_gates.setdefault(section.name, []), section_gate_values)
            for item in section.criteria:
                if item.section is not None and item.section != section.name:
                    msg = (
                        f"Rubric item section {item.section!r} conflicts with "
                        f"nested section {section.name!r}"
                    )
                    raise ValueError(msg)
                items.append(item.model_copy(update={"section": section.name}))
    else:
        items = [item.model_copy() for item in rubric]

    for index, item in enumerate(items):
        if item.gate and item.section is None:
            msg = f"Rubric item {index} is marked as a gate but has no section"
            raise ValueError(msg)
        item_section_gate_values: list[str] = []
        if item.section_gate is not None:
            item_section_gate_values.append(item.section_gate)
        item_section_gate_values.extend(item.section_gates)
        if item_section_gate_values and item.section is None:
            msg = f"Rubric item {index} defines section_gate/section_gates but has no section"
            raise ValueError(msg)
        if item.section_tolerance_pct is not None:
            if item.section is None:
                msg = f"Rubric item {index} defines section_tolerance_pct but has no section"
                raise ValueError(msg)
            section_tolerances.setdefault(item.section, item.section_tolerance_pct)
        if item.section is not None:
            _append_unique_strings(section_gates.setdefault(item.section, []), item_section_gate_values)

    normalised_items: list[RubricItem] = []
    for item in items:
        gates = section_gates.get(item.section, []) if item.section is not None else []
        tolerance_pct = (
            section_tolerances.get(item.section)
            if item.section is not None and not item.gate
            else None
        )
        normalised_items.append(
            item.model_copy(
                update={
                    "section_gate": gates[0] if len(gates) == 1 else None,
                    "section_gates": gates,
                    "section_tolerance_pct": tolerance_pct,
                }
            )
        )

    return normalised_items


DEFAULT_JUDGE_MODEL = "anthropic/claude-opus-4-7"
ReasoningEffort = Literal["low", "medium", "high", "xhigh", "none"]


class GraderConfig(BaseModel):
    """Top-level grader configuration loaded from a TOML file.

    mode controls how rubric criteria are evaluated:
      - "individual": each criterion is evaluated in its own agent
        session (one invocation of gandalf-the-grader-judge per criterion).
      - "batch" (default): all criteria are sent to a single agent session,
        which writes a JSON array of verdicts in one go.

    batch_splits controls explicit positional chunks in batch mode. When it is
    unset and auto_parallel is true, batch mode computes chunks from
    auto_parallel_target_chunk_size.  Ignored in individual mode.

    max_concurrency is a hard cap on parallel judge sessions.  When omitted,
    batch auto-parallelism uses rate_limit_max_concurrency as its adaptive cap;
    individual mode remains serial unless max_concurrency is set.

    judge_timeout is the per-criterion budget in seconds, regardless of mode.
    In batch mode the effective timeout per session is
    ``judge_timeout * n_criteria_in_session``, optionally capped by
    batch_timeout.
    """

    model: str = DEFAULT_JUDGE_MODEL
    reasoning_effort: ReasoningEffort | None = None
    instructions: str | None = None
    instructions_path: str | None = None
    rubric: RubricInput | None = None
    rubric_path: str | None = None
    workdir: str
    trajectory_path: str
    sandbox_user: str | None = None
    mcp_servers: list[MCPServer] = Field(default_factory=list)
    excel_backend: ExcelBackendConfig = Field(default_factory=ExcelBackendConfig)
    workbook_repair_check: WorkbookRepairCheckConfig = Field(default_factory=WorkbookRepairCheckConfig)
    workbook_digest: WorkbookDigestConfig = Field(default_factory=WorkbookDigestConfig)
    output_dir: str
    judge_timeout: int = 300
    judge_guidance: str | None = None
    judge_guidance_path: str | None = None
    judge_prompt: str | None = None
    judge_prompt_path: str | None = None
    batch_timeout: int | None = None
    mode: Literal["individual", "batch"] = "batch"
    batch_splits: int | None = Field(default=None, ge=2)
    max_concurrency: int | None = Field(default=None, ge=1)
    auto_parallel: bool = True
    auto_parallel_target_chunk_size: int = Field(default=25, ge=1)
    rate_limit_discovery: bool = True
    rate_limit_initial_concurrency: int = Field(default=2, ge=1)
    rate_limit_max_concurrency: int = Field(default=32, ge=1)
    rate_limit_backoff_seconds: int = Field(default=15, ge=1)
    rate_limit_max_requeues_per_chunk: int = Field(default=3, ge=0)
    batch_recovery_enabled: bool = True
    batch_recovery_target_chunk_size: int = Field(default=8, ge=1)
    batch_recovery_max_rounds: int = Field(default=2, ge=0)
    judge_retries: int = 1
    golden_check: bool = False

    @model_validator(mode="before")
    @classmethod
    def _reject_cli_only_section_filters(cls, data: Any) -> Any:
        if not isinstance(data, dict):
            return data
        forbidden = sorted(
            set(data)
            & {
                "include_section",
                "include_sections",
                "exclude_section",
                "exclude_sections",
            }
        )
        if forbidden:
            msg = (
                "Section filters are CLI-only; use --include-section and/or "
                f"--exclude-section instead of config field(s): {', '.join(forbidden)}"
            )
            raise ValueError(msg)
        return data

    @model_validator(mode="after")
    def _check_no_inline_and_path(self) -> "GraderConfig":
        if self.instructions is not None and self.instructions_path is not None:
            msg = "Cannot set both 'instructions' and 'instructions_path'"
            raise ValueError(msg)
        if self.rubric is not None and self.rubric_path is not None:
            msg = "Cannot set both 'rubric' and 'rubric_path'"
            raise ValueError(msg)
        if self.rubric is None and self.rubric_path is None:
            msg = "Must set either 'rubric' or 'rubric_path'"
            raise ValueError(msg)
        if self.rubric is not None:
            self.rubric = normalise_rubric(self.rubric)
        if self.judge_guidance is not None and self.judge_guidance_path is not None:
            msg = "Cannot set both 'judge_guidance' and 'judge_guidance_path'"
            raise ValueError(msg)
        if self.judge_prompt is not None and self.judge_prompt_path is not None:
            msg = "Cannot set both 'judge_prompt' and 'judge_prompt_path'"
            raise ValueError(msg)
        if self.batch_splits is not None and self.mode != "batch":
            msg = "'batch_splits' can only be used with mode='batch'"
            raise ValueError(msg)
        if self.excel_backend.enabled and any(srv.name == "excel" for srv in self.mcp_servers):
            msg = (
                "excel_backend.enabled automatically attaches an MCP server named 'excel'; "
                "remove or rename the configured MCP server named 'excel'"
            )
            raise ValueError(msg)
        return self


class _BaseJudgeInput(BaseModel):
    """Shared fields for all judge input types."""

    model: str
    reasoning_effort: ReasoningEffort | None = None
    instructions: str
    final_output: str
    workdir: str
    mcp_servers: list[MCPServer] = Field(default_factory=list)
    excel_backend: ExcelBackendConfig = Field(default_factory=ExcelBackendConfig)
    judge_guidance: str = ""
    judge_prompt: str | None = None


class JudgeInput(_BaseJudgeInput):
    """Input passed to the inner judge for a single criterion evaluation."""

    criterion: str


class BatchJudgeInput(_BaseJudgeInput):
    """Input passed to the inner judge for batch (all-criteria) evaluation.

    Weights are intentionally omitted from criteria so the judge evaluates
    each criterion on its own merits.  Indices are derived from position.
    """

    criteria: list[str]


class LLMUsage(BaseModel):
    """Aggregate LLM token and cost metrics from judge sessions."""

    cost_usd: float = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    cache_read_tokens: int = 0

    def __add__(self, other: "LLMUsage") -> "LLMUsage":
        """Sum two LLMUsage instances field-by-field."""
        return LLMUsage(
            cost_usd=self.cost_usd + other.cost_usd,
            prompt_tokens=self.prompt_tokens + other.prompt_tokens,
            completion_tokens=self.completion_tokens + other.completion_tokens,
            cache_read_tokens=self.cache_read_tokens + other.cache_read_tokens,
        )


class Verdict(BaseModel):
    """Verdict returned by the inner judge."""

    met: bool | None
    reasoning: str
    evidence: list[str] = Field(default_factory=list)

    @classmethod
    def from_raw(cls, data: dict[str, Any]) -> "Verdict":
        """Create a Verdict from a raw JSON-parsed dict, normalizing types."""
        raw_met = data.get("met")
        return cls(
            met=bool(raw_met) if raw_met is not None else None,
            reasoning=str(data.get("reasoning", "No reasoning provided.")),
            evidence=list(data.get("evidence", [])),
        )

    @classmethod
    def errors(cls, n: int, reason: str) -> list["Verdict"]:
        """Return *n* error verdicts (met=None) that all share the same reason."""
        return [cls(met=None, reasoning=reason) for _ in range(n)]


class CriterionResult(BaseModel):
    """Result for a single criterion evaluation."""

    criterion: str
    weight: float
    section: str | None = None
    gate: bool = False
    met: bool | None
    reasoning: str
    evidence: list[str] = Field(default_factory=list)
    score_contribution: float = 0.0
    section_tolerance_pct: float | None = None
    skipped: bool = False
    skip_reason: str | None = None

    @field_validator("section", "skip_reason")
    @classmethod
    def _normalise_optional_string(cls, value: str | None, info: ValidationInfo) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value:
            msg = f"{info.field_name} cannot be empty"
            raise ValueError(msg)
        return value

    @model_validator(mode="after")
    def _check_gate_has_section(self) -> "CriterionResult":
        if self.gate and self.section is None:
            msg = "gate criterion results must belong to a section"
            raise ValueError(msg)
        return self


class SectionGateResult(BaseModel):
    """Verdict for a section-level gate that is not a rubric criterion."""

    section: str
    index: int = 0
    criterion: str
    met: bool | None
    reasoning: str
    evidence: list[str] = Field(default_factory=list)
    skipped: bool = False
    skip_reason: str | None = None

    @field_validator("skip_reason")
    @classmethod
    def _normalise_optional_string(cls, value: str | None, info: ValidationInfo) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value:
            msg = f"{info.field_name} cannot be empty"
            raise ValueError(msg)
        return value


class SectionResult(BaseModel):
    """Scoring summary for one rubric section.

    Gate index fields are 0-based indices into EvaluationInfo.criterion_results,
    not positions within this section only.
    """

    section: str
    score: float
    minimum_score: float = 0.0
    maximum_score: float = 0.0
    gate_count: int = 0
    passed_gate_indices: list[int] = Field(default_factory=list)
    failed_gate_indices: list[int] = Field(default_factory=list)
    section_gate: str | None = None
    section_gate_met: bool | None = None
    section_gate_reasoning: str | None = None
    section_gate_evidence: list[str] = Field(default_factory=list)
    section_gate_count: int = 0
    section_gate_results: list[SectionGateResult] = Field(default_factory=list)
    passed_section_gate_indices: list[int] = Field(default_factory=list)
    failed_section_gate_indices: list[int] = Field(default_factory=list)
    errored_section_gate_indices: list[int] = Field(default_factory=list)
    skipped_section_gate_indices: list[int] = Field(default_factory=list)
    section_gates_met: bool | None = None
    section_tolerance_pct: float | None = None
    gated_out: bool = False


class WorkbookRepairCheckResult(BaseModel):
    """Result of the deterministic workbook repair dialogue check."""

    enabled: bool = True
    checked: bool = False
    workbook_path: str | None = None
    backend: str | None = None
    opened: bool | None = None
    repair_dialog_detected: bool | None = None
    skipped_reason: str | None = None
    error: str | None = None
    details: dict[str, Any] = Field(default_factory=dict)


class EvaluationInfo(BaseModel):
    """Full evaluation output with reward/raw score, per-criterion results, and LLM usage."""

    reward: float
    raw_score: float
    unpenalized_reward: float | None = None
    repair_penalty_applied: bool = False
    repair_penalty_multiplier: float | None = None
    workbook_repair_check: WorkbookRepairCheckResult | None = None
    minimum_score: float = 0.0
    maximum_score: float = 0.0
    criterion_results: list[CriterionResult]
    section_results: list[SectionResult] = Field(default_factory=list)
    llm_usage: LLMUsage = Field(default_factory=LLMUsage)
    excel_metrics: dict[str, Any] | None = None
    batch_recovery: dict[str, Any] | None = None
    errored_criterion_count: int = 0
    errored_section_gate_count: int = 0
    evaluated_criteria_pct: float = 100.0


def load_config(path: str) -> GraderConfig:
    """Load grader configuration from a TOML file."""
    with open(path, "rb") as f:
        data = tomllib.load(f)
    return GraderConfig.model_validate(data)


def load_rubric(path: str) -> list[RubricItem]:
    """Load rubric items from a JSON file."""
    raw = Path(path).read_bytes()
    rubric: RubricInput = TypeAdapter(RubricInput).validate_json(raw)
    return normalise_rubric(rubric)
