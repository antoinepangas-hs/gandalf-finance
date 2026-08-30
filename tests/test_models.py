"""Tests for gandalf.models."""

import math
import os
import pathlib
from typing import Any, ClassVar

import pytest
from pydantic import ValidationError

from gandalf.models import (
    BatchJudgeInput,
    CriterionResult,
    DEFAULT_JUDGE_MODEL,
    ExcelBackendConfig,
    EvaluationInfo,
    GraderConfig,
    JudgeInput,
    MCPServer,
    RubricDocument,
    RubricItem,
    RubricSection,
    Verdict,
    WorkbookDigestConfig,
    WorkbookRepairCheckConfig,
    WorkbookRepairCheckResult,
    load_config,
    load_rubric,
)

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


class TestLoadConfig:
    def test_parses_all_fields(self) -> None:
        cfg = load_config(os.path.join(FIXTURES, "sample_grader.toml"))
        assert cfg.model == DEFAULT_JUDGE_MODEL
        assert cfg.sandbox_user == "sandbox"
        assert cfg.instructions == "Build a web app that displays hello world."
        assert cfg.rubric_path == "/tests/rubric.json"
        assert cfg.workdir == "/home/agent/workspace"
        assert cfg.trajectory_path == "/logs/agent/trajectory.json"
        assert cfg.output_dir == "/logs/grader"
        assert cfg.judge_timeout == 120

    def test_parses_mcp_servers(self) -> None:
        cfg = load_config(os.path.join(FIXTURES, "sample_grader.toml"))
        assert len(cfg.mcp_servers) == 1
        mcp = cfg.mcp_servers[0]
        assert mcp.name == "magic-server"
        assert mcp.transport == "stdio"
        assert mcp.command == "/usr/bin/mcp-server"
        assert mcp.args == ["--verbose"]

    def test_defaults_model(self, tmp_path: pathlib.Path) -> None:
        toml_content = """\
sandbox_user = "sandbox"
instructions = "Do something."
rubric_path = "/tests/rubric.json"
workdir = "/workspace"
trajectory_path = "/logs/trajectory.json"
output_dir = "/logs/grader"
"""
        p = tmp_path / "grader.toml"
        p.write_text(toml_content)
        cfg = load_config(str(p))
        assert cfg.model == DEFAULT_JUDGE_MODEL

    def test_parses_reasoning_effort(self, tmp_path: pathlib.Path) -> None:
        toml_content = """\
model = "openai/gpt-5.5"
reasoning_effort = "xhigh"
sandbox_user = "sandbox"
instructions = "Do something."
rubric_path = "/tests/rubric.json"
workdir = "/workspace"
trajectory_path = "/logs/trajectory.json"
output_dir = "/logs/grader"
"""
        p = tmp_path / "grader.toml"
        p.write_text(toml_content)
        cfg = load_config(str(p))
        assert cfg.reasoning_effort == "xhigh"

    def test_defaults_timeout(self, tmp_path: pathlib.Path) -> None:
        toml_content = """\
model = "openai/gpt-4o"
sandbox_user = "sandbox"
instructions = "Do something."
rubric_path = "/tests/rubric.json"
workdir = "/workspace"
trajectory_path = "/logs/trajectory.json"
output_dir = "/logs/grader"
"""
        p = tmp_path / "grader.toml"
        p.write_text(toml_content)
        cfg = load_config(str(p))
        assert cfg.judge_timeout == 300

    def test_parses_enabled_excel_backend_with_mac_default(self, tmp_path: pathlib.Path) -> None:
        toml_content = """\
instructions = "Do something."
rubric_path = "/tests/rubric.json"
workdir = "/workspace"
trajectory_path = "/logs/trajectory.json"
output_dir = "/logs/grader"

[excel_backend]
enabled = true
"""
        p = tmp_path / "grader.toml"
        p.write_text(toml_content)
        cfg = load_config(str(p))

        assert cfg.excel_backend.enabled is True
        assert cfg.excel_backend.backend == "mac_excel"
        assert cfg.excel_backend.timeout_seconds == 120
        assert cfg.excel_backend.allow_macros is False
        assert cfg.excel_backend.max_cells_per_call == 10_000
        assert cfg.excel_backend.auth_token is None

    def test_parses_workbook_repair_check(self, tmp_path: pathlib.Path) -> None:
        toml_content = """\
instructions = "Do something."
rubric_path = "/tests/rubric.json"
workdir = "/workspace"
trajectory_path = "/logs/trajectory.json"
output_dir = "/logs/grader"

[workbook_repair_check]
enabled = true
workbook_path = "model.xlsx"
penalty_multiplier = 0.25
"""
        p = tmp_path / "grader.toml"
        p.write_text(toml_content)
        cfg = load_config(str(p))

        assert cfg.workbook_repair_check.enabled is True
        assert cfg.workbook_repair_check.workbook_path == "model.xlsx"
        assert cfg.workbook_repair_check.penalty_multiplier == 0.25

    def test_defaults_excel_backend_to_local_mac(self, tmp_path: pathlib.Path) -> None:
        toml_content = """\
instructions = "Do something."
rubric_path = "/tests/rubric.json"
workdir = "/workspace"
trajectory_path = "/logs/trajectory.json"
output_dir = "/logs/grader"
"""
        p = tmp_path / "grader.toml"
        p.write_text(toml_content)
        cfg = load_config(str(p))

        assert cfg.excel_backend.enabled is True
        assert cfg.excel_backend.backend == "mac_excel"

    def test_missing_file_raises(self) -> None:
        with pytest.raises(FileNotFoundError):
            load_config("/nonexistent/grader.toml")

    def test_missing_required_field_raises(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "bad.toml"
        p.write_text('model = "x"\n')
        with pytest.raises(ValidationError):
            load_config(str(p))


class TestLoadRubric:
    def test_parses_items(self) -> None:
        rubric = load_rubric(os.path.join(FIXTURES, "sample_rubric.json"))
        assert len(rubric) == 3
        assert rubric[0].criterion == "The file index.html exists in the workspace"
        assert rubric[0].weight == 1.0
        assert rubric[1].weight == 2.0

    def test_empty_rubric(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "empty.json"
        p.write_text("[]")
        rubric = load_rubric(str(p))
        assert rubric == []

    def test_missing_file_raises(self) -> None:
        with pytest.raises(FileNotFoundError):
            load_rubric("/nonexistent/rubric.json")

    def test_parses_negative_weight_items(self) -> None:
        rubric = load_rubric(os.path.join(FIXTURES, "sample_rubric_with_negatives.json"))
        assert len(rubric) == 3
        assert rubric[0].weight == 2.0
        assert rubric[1].weight == 3.0
        assert rubric[2].weight == -1.0

    def test_parses_flat_section_and_gate_metadata(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {"criterion": "Workbook exists", "weight": 0, "section": "Model Build", "gate": true},
  {"criterion": "Forecast formulas are correct", "weight": 5, "section": "Model Build"}
]
"""
        )

        rubric = load_rubric(str(p))

        assert len(rubric) == 2
        assert rubric[0].section == "Model Build"
        assert rubric[0].gate is True
        assert rubric[1].section == "Model Build"
        assert rubric[1].gate is False

    def test_parses_nested_sections(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
{
  "criteria": [
    {"criterion": "Top-level deliverable exists", "weight": 1}
  ],
  "sections": [
    {
      "name": "Model Build",
      "criteria": [
        {"criterion": "Workbook exists", "weight": 0, "gate": true},
        {"criterion": "Forecast formulas are correct", "weight": 5}
      ]
    },
    {
      "name": "Formatting",
      "criteria": [
        {"criterion": "Uses clear number formats", "weight": 2}
      ]
    }
  ]
}
"""
        )

        rubric = load_rubric(str(p))

        assert [item.criterion for item in rubric] == [
            "Top-level deliverable exists",
            "Workbook exists",
            "Forecast formulas are correct",
            "Uses clear number formats",
        ]
        assert rubric[0].section is None
        assert rubric[1].section == "Model Build"
        assert rubric[1].gate is True
        assert rubric[3].section == "Formatting"

    def test_parses_nested_section_gate(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
{
  "sections": [
    {
      "name": "Model Build",
      "gate": "The submitted workbook exists and can be opened.",
      "criteria": [
        {"criterion": "Forecast formulas are correct", "weight": 5},
        {"criterion": "Revenue model is linked", "weight": 2}
      ]
    }
  ]
}
"""
        )

        rubric = load_rubric(str(p))

        assert len(rubric) == 2
        assert all(item.section == "Model Build" for item in rubric)
        assert all(item.section_gate == "The submitted workbook exists and can be opened." for item in rubric)
        assert all(item.section_gates == ["The submitted workbook exists and can be opened."] for item in rubric)

    def test_parses_nested_section_gates(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
{
  "sections": [
    {
      "name": "Model Build",
      "gate": "The submitted workbook exists.",
      "gates": [
        "The submitted workbook opens.",
        "The submitted workbook contains the required tabs."
      ],
      "criteria": [
        {"criterion": "Forecast formulas are correct", "weight": 5}
      ]
    }
  ]
}
"""
        )

        rubric = load_rubric(str(p))

        assert rubric[0].section_gate is None
        assert rubric[0].section_gates == [
            "The submitted workbook exists.",
            "The submitted workbook opens.",
            "The submitted workbook contains the required tabs.",
        ]

    def test_flat_section_gate_is_propagated_to_section_items(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Forecast formulas are correct",
    "weight": 5,
    "section": "Model Build",
    "section_gate": "The submitted workbook exists and can be opened."
  },
  {"criterion": "Revenue model is linked", "weight": 2, "section": "Model Build"}
]
"""
        )

        rubric = load_rubric(str(p))

        assert all(item.section_gate == "The submitted workbook exists and can be opened." for item in rubric)
        assert all(item.section_gates == ["The submitted workbook exists and can be opened."] for item in rubric)

    def test_flat_section_gates_are_combined_and_deduplicated(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Forecast formulas are correct",
    "weight": 5,
    "section": "Model Build",
    "section_gate": "The submitted workbook exists.",
    "section_gates": ["The submitted workbook opens."]
  },
  {
    "criterion": "Revenue model is linked",
    "weight": 2,
    "section": "Model Build",
    "section_gates": [
      "The submitted workbook opens.",
      "The submitted workbook contains the required tabs."
    ]
  }
]
"""
        )

        rubric = load_rubric(str(p))

        assert all(item.section_gate is None for item in rubric)
        assert all(
            item.section_gates
            == [
                "The submitted workbook exists.",
                "The submitted workbook opens.",
                "The submitted workbook contains the required tabs.",
            ]
            for item in rubric
        )

    def test_rejects_gate_without_section(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text('[{"criterion": "Workbook exists", "weight": 0, "gate": true}]')

        with pytest.raises(ValueError, match="gate.*section"):
            load_rubric(str(p))

    def test_rejects_section_gate_without_section(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Forecast formulas are correct",
    "weight": 5,
    "section_gate": "The submitted workbook exists."
  }
]
"""
        )

        with pytest.raises(ValueError, match="section_gate.*section"):
            load_rubric(str(p))

    def test_rejects_section_gates_without_section(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Forecast formulas are correct",
    "weight": 5,
    "section_gates": ["The submitted workbook exists."]
  }
]
"""
        )

        with pytest.raises(ValueError, match="section_gate.*section"):
            load_rubric(str(p))

    def test_rejects_empty_section_gate_string(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Forecast formulas are correct",
    "weight": 5,
    "section": "Model Build",
    "section_gates": ["   "]
  }
]
"""
        )

        with pytest.raises(ValidationError, match="section_gates"):
            load_rubric(str(p))

    def test_rejects_duplicate_nested_section_names(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
{
  "sections": [
    {"name": "Model Build", "criteria": []},
    {"name": "Model Build", "criteria": []}
  ]
}
"""
        )

        with pytest.raises(ValidationError, match="Duplicate rubric section"):
            load_rubric(str(p))

    def test_rejects_conflicting_nested_item_section(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
{
  "sections": [
    {
      "name": "Model Build",
      "criteria": [
        {"criterion": "Workbook exists", "weight": 1, "section": "Formatting"}
      ]
    }
  ]
}
"""
        )

        with pytest.raises(ValueError, match="conflicts"):
            load_rubric(str(p))

    def test_multiple_flat_section_gate_values_are_supported(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Forecast formulas are correct",
    "weight": 5,
    "section": "Model Build",
    "section_gate": "The submitted workbook exists."
  },
  {
    "criterion": "Revenue model is linked",
    "weight": 2,
    "section": "Model Build",
    "section_gate": "The submitted workbook is final."
  }
]
"""
        )

        rubric = load_rubric(str(p))

        assert all(
            item.section_gates == ["The submitted workbook exists.", "The submitted workbook is final."]
            for item in rubric
        )

    def test_parses_nested_section_tolerance(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
{
  "sections": [
    {
      "name": "Model Build",
      "tolerance_pct": 1.0,
      "criteria": [
        {"criterion": "Workbook exists", "weight": 0, "gate": true},
        {"criterion": "Revenue is $1,000", "weight": 5},
        {"criterion": "Used hardcoded values", "weight": -1}
      ]
    }
  ]
}
"""
        )

        rubric = load_rubric(str(p))

        assert rubric[0].section_tolerance_pct is None
        assert rubric[1].section_tolerance_pct == 1.0
        assert rubric[2].section_tolerance_pct == 1.0

    def test_flat_section_tolerance_first_value_wins_and_skips_gates(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Workbook exists",
    "weight": 0,
    "section": "Model Build",
    "gate": true,
    "section_tolerance_pct": 1.5
  },
  {
    "criterion": "Revenue is $1,000",
    "weight": 5,
    "section": "Model Build",
    "section_tolerance_pct": 2.0
  },
  {"criterion": "EBITDA is $200", "weight": 2, "section": "Model Build"}
]
"""
        )

        rubric = load_rubric(str(p))

        assert rubric[0].section_tolerance_pct is None
        assert rubric[1].section_tolerance_pct == 1.5
        assert rubric[2].section_tolerance_pct == 1.5

    def test_rejects_section_tolerance_without_section(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text('[{"criterion": "Revenue is $1,000", "weight": 1, "section_tolerance_pct": 1.0}]')

        with pytest.raises(ValueError, match="section_tolerance_pct.*section"):
            load_rubric(str(p))

    def test_rejects_negative_section_tolerance(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
{
  "sections": [
    {"name": "Model Build", "tolerance_pct": -1.0, "criteria": []}
  ]
}
"""
        )

        with pytest.raises(ValidationError, match="tolerance_pct"):
            load_rubric(str(p))

    def test_rejects_non_finite_section_tolerance(self) -> None:
        with pytest.raises(ValidationError, match="section_tolerance_pct"):
            RubricItem(
                criterion="Revenue is $1,000",
                weight=1.0,
                section="Model Build",
                section_tolerance_pct=math.inf,
            )

    def test_rejects_invalid_section_tolerance_type(self, tmp_path: pathlib.Path) -> None:
        p = tmp_path / "rubric.json"
        p.write_text(
            """\
[
  {
    "criterion": "Revenue is $1,000",
    "weight": 1,
    "section": "Model Build",
    "section_tolerance_pct": "loose"
  }
]
"""
        )

        with pytest.raises(ValidationError, match="section_tolerance_pct"):
            load_rubric(str(p))


class TestPydanticModels:
    def test_mcp_server_defaults(self) -> None:
        srv = MCPServer(name="test", command="/bin/test")
        assert srv.transport == "stdio"
        assert srv.args == []
        assert srv.env == {}
        assert srv.url is None
        assert srv.headers == {}

    def test_mcp_server_stdio_requires_command(self) -> None:
        with pytest.raises(ValidationError, match="'command' is required"):
            MCPServer(name="test")

    def test_mcp_server_remote_streamable_http(self) -> None:
        srv = MCPServer(
            name="remote",
            transport="streamable-http",
            url="http://localhost:8000/mcp",
        )
        assert srv.transport == "streamable-http"
        assert srv.url == "http://localhost:8000/mcp"
        assert srv.command is None

    def test_mcp_server_remote_http_with_headers(self) -> None:
        srv = MCPServer(
            name="remote",
            transport="http",
            url="https://api.example.com/mcp",
            headers={"Authorization": "Bearer token"},
        )
        assert srv.transport == "http"
        assert srv.headers == {"Authorization": "Bearer token"}

    def test_mcp_server_remote_sse(self) -> None:
        srv = MCPServer(
            name="remote",
            transport="sse",
            url="http://localhost:8000/sse",
        )
        assert srv.transport == "sse"

    def test_mcp_server_remote_requires_url(self) -> None:
        with pytest.raises(ValidationError, match="'url' is required"):
            MCPServer(name="remote", transport="streamable-http")

    def test_mcp_server_rejects_unknown_transport(self) -> None:
        with pytest.raises(ValidationError):
            MCPServer(name="test", command="/bin/test", transport="grpc")  # type: ignore[arg-type]

    def test_excel_backend_defaults_to_local_mac_enabled(self) -> None:
        cfg = ExcelBackendConfig()
        assert cfg.enabled is True
        assert cfg.backend == "mac_excel"
        assert cfg.windows_url is None
        assert cfg.auth_token is None

    def test_excel_backend_can_be_disabled(self) -> None:
        cfg = ExcelBackendConfig(enabled=False)
        assert cfg.enabled is False
        assert cfg.backend == "mac_excel"

    def test_workbook_repair_check_defaults_to_half_penalty(self) -> None:
        cfg = WorkbookRepairCheckConfig()
        assert cfg.enabled is True
        assert cfg.workbook_path is None
        assert cfg.penalty_multiplier == 0.5

    def test_workbook_repair_check_rejects_invalid_values(self) -> None:
        with pytest.raises(ValidationError, match="workbook_path"):
            WorkbookRepairCheckConfig(workbook_path=" ")
        with pytest.raises(ValidationError, match="penalty_multiplier"):
            WorkbookRepairCheckConfig(penalty_multiplier=-0.1)
        with pytest.raises(ValidationError, match="penalty_multiplier"):
            WorkbookRepairCheckConfig(penalty_multiplier=1.1)

    def test_workbook_digest_defaults_off(self) -> None:
        cfg = WorkbookDigestConfig()
        assert cfg.enabled is False
        assert cfg.workbook_path is None
        assert cfg.max_sheets == 30
        assert cfg.max_formula_samples == 15

    def test_workbook_digest_rejects_invalid_values(self) -> None:
        with pytest.raises(ValidationError, match="workbook_path"):
            WorkbookDigestConfig(workbook_path="  ")
        with pytest.raises(ValidationError, match="max_sheets"):
            WorkbookDigestConfig(max_sheets=0)

    def test_excel_backend_windows_vm_requires_url(self) -> None:
        with pytest.raises(ValidationError, match="windows_url"):
            ExcelBackendConfig(enabled=True, backend="windows_vm")

    def test_excel_backend_accepts_libreoffice_without_windows_url(self) -> None:
        cfg = ExcelBackendConfig(enabled=True, backend="libreoffice")

        assert cfg.enabled is True
        assert cfg.backend == "libreoffice"
        assert cfg.windows_url is None

    def test_excel_backend_rejects_invalid_backend(self) -> None:
        with pytest.raises(ValidationError):
            ExcelBackendConfig(backend="numbers")  # type: ignore[arg-type]

    def test_excel_backend_rejects_invalid_limits(self) -> None:
        with pytest.raises(ValidationError, match="timeout_seconds"):
            ExcelBackendConfig(timeout_seconds=0)
        with pytest.raises(ValidationError, match="max_cells_per_call"):
            ExcelBackendConfig(max_cells_per_call=0)
        with pytest.raises(ValidationError, match="max_format_cells_per_call"):
            ExcelBackendConfig(max_format_cells_per_call=0)

    def test_excel_backend_normalises_auth_token(self) -> None:
        cfg = ExcelBackendConfig(auth_token=" secret/ ")
        assert cfg.auth_token == "secret/"

    def test_excel_backend_rejects_empty_auth_token(self) -> None:
        with pytest.raises(ValidationError, match="auth_token"):
            ExcelBackendConfig(auth_token=" ")

    def test_grader_config_rejects_excel_mcp_name_conflict(self) -> None:
        with pytest.raises(ValidationError, match="MCP server named 'excel'"):
            GraderConfig(
                instructions="test",
                rubric_path="/rubric.json",
                workdir="/workspace",
                trajectory_path="/logs/trajectory.json",
                mcp_servers=[MCPServer(name="excel", command="/bin/excel-mcp")],
                excel_backend=ExcelBackendConfig(enabled=True),
                output_dir="/logs/grader",
            )

    def test_grader_config_allows_custom_excel_mcp_when_backend_disabled(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            mcp_servers=[MCPServer(name="excel", command="/bin/excel-mcp")],
            excel_backend=ExcelBackendConfig(enabled=False),
            output_dir="/logs/grader",
        )
        assert cfg.excel_backend.enabled is False

    def test_grader_config_has_trajectory_path(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            sandbox_user="sandbox",
            output_dir="/logs/grader",
        )
        assert cfg.trajectory_path == "/logs/trajectory.json"
        assert cfg.model == DEFAULT_JUDGE_MODEL

    def test_grader_config_judge_guidance_path_defaults_none(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            sandbox_user="sandbox",
            output_dir="/logs/grader",
        )
        assert cfg.judge_guidance_path is None

    def test_grader_config_judge_guidance_path_set(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            sandbox_user="sandbox",
            output_dir="/logs/grader",
            judge_guidance_path="/opt/grader/judge-guidance.md",
        )
        assert cfg.judge_guidance_path == "/opt/grader/judge-guidance.md"

    def test_judge_input_includes_final_output(self) -> None:
        ji = JudgeInput(
            model="test-model",
            instructions="test",
            final_output="agent said done",
            criterion="check something",
            workdir="/workspace",
        )
        assert ji.final_output == "agent said done"

    def test_judge_input_guidance_defaults_empty(self) -> None:
        ji = JudgeInput(
            model="test-model",
            instructions="test",
            final_output="done",
            criterion="check",
            workdir="/workspace",
        )
        assert ji.judge_guidance == ""

    def test_judge_input_guidance_roundtrip(self) -> None:
        ji = JudgeInput(
            model="test-model",
            instructions="test",
            final_output="done",
            criterion="check",
            workdir="/workspace",
            judge_guidance="Use openpyxl for .xlsx files.",
        )
        raw = ji.model_dump_json()
        restored = JudgeInput.model_validate_json(raw)
        assert restored.judge_guidance == "Use openpyxl for .xlsx files."

    def test_verdict_defaults(self) -> None:
        v = Verdict(met=True, reasoning="ok")
        assert v.evidence == []

    def test_verdict_with_evidence(self) -> None:
        v = Verdict(met=False, reasoning="fail", evidence=["check1", "check2"])
        assert len(v.evidence) == 2

    def test_verdict_met_none(self) -> None:
        v = Verdict(met=None, reasoning="error")
        assert v.met is None
        data = v.model_dump()
        assert data["met"] is None

    def test_verdict_none_serialization_roundtrip(self) -> None:
        v = Verdict(met=None, reasoning="error")
        raw = v.model_dump_json()
        restored = Verdict.model_validate_json(raw)
        assert restored.met is None

    def test_criterion_result(self) -> None:
        r = CriterionResult(
            criterion="test",
            weight=1.0,
            met=True,
            reasoning="ok",
        )
        assert r.evidence == []

    def test_criterion_result_negative_weight(self) -> None:
        r = CriterionResult(
            criterion="used hardcoded values",
            weight=-1.0,
            met=True,
            reasoning="found hardcoded values",
        )
        assert r.weight == -1.0

    def test_criterion_result_met_none(self) -> None:
        r = CriterionResult(criterion="test", weight=1.0, met=None, reasoning="error")
        assert r.met is None
        data = r.model_dump()
        assert data["met"] is None

    def test_evaluation_info(self) -> None:
        info = EvaluationInfo(
            reward=0.5,
            raw_score=3.0,
            minimum_score=-1.0,
            maximum_score=6.0,
            criterion_results=[
                CriterionResult(criterion="c1", weight=3.0, met=True, reasoning="ok"),
                CriterionResult(criterion="c2", weight=3.0, met=False, reasoning="fail"),
                CriterionResult(criterion="c3", weight=-1.0, met=False, reasoning="avoided"),
            ],
        )
        assert info.reward == 0.5
        assert info.raw_score == 3.0
        assert info.minimum_score == -1.0
        assert info.maximum_score == 6.0
        assert len(info.criterion_results) == 3

    def test_evaluation_info_accepts_workbook_repair_check_result(self) -> None:
        info = EvaluationInfo(
            reward=0.5,
            raw_score=10.0,
            unpenalized_reward=1.0,
            repair_penalty_applied=True,
            repair_penalty_multiplier=0.5,
            workbook_repair_check=WorkbookRepairCheckResult(
                checked=True,
                workbook_path="submitted_output.xlsx",
                backend="mac_excel",
                opened=True,
                repair_dialog_detected=True,
            ),
            criterion_results=[],
        )

        assert info.reward == 0.5
        assert info.workbook_repair_check is not None
        assert info.workbook_repair_check.repair_dialog_detected is True

    def test_grader_config_sandbox_user_defaults_none(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            output_dir="/logs/grader",
        )
        assert cfg.sandbox_user is None

    def test_grader_config_sandbox_user_explicit(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            output_dir="/logs/grader",
            sandbox_user="sandbox",
        )
        assert cfg.sandbox_user == "sandbox"

    def test_grader_config_sandbox_user_omitted_from_toml(self, tmp_path: pathlib.Path) -> None:
        toml_content = """\
instructions = "Do something."
rubric_path = "/tests/rubric.json"
workdir = "/workspace"
trajectory_path = "/logs/trajectory.json"
output_dir = "/logs/grader"
"""
        p = tmp_path / "grader.toml"
        p.write_text(toml_content)
        cfg = load_config(str(p))
        assert cfg.sandbox_user is None

    def test_grader_config_judge_retries_default(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            sandbox_user="sandbox",
            output_dir="/logs/grader",
        )
        assert cfg.judge_retries == 1

    def test_grader_config_judge_retries_explicit(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            sandbox_user="sandbox",
            output_dir="/logs/grader",
            judge_retries=3,
        )
        assert cfg.judge_retries == 3

    def test_grader_config_golden_check_default(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            sandbox_user="sandbox",
            output_dir="/logs/grader",
        )
        assert cfg.golden_check is False

    def test_grader_config_golden_check_explicit(self) -> None:
        cfg = GraderConfig(
            instructions="test",
            rubric_path="/rubric.json",
            workdir="/workspace",
            trajectory_path="/logs/trajectory.json",
            sandbox_user="sandbox",
            output_dir="/logs/grader",
            golden_check=True,
        )
        assert cfg.golden_check is True

    def test_evaluation_info_errored_fields(self) -> None:
        info = EvaluationInfo(
            reward=0.5,
            raw_score=1.0,
            criterion_results=[
                CriterionResult(criterion="c1", weight=1.0, met=True, reasoning="ok"),
                CriterionResult(criterion="c2", weight=1.0, met=None, reasoning="error"),
            ],
            errored_criterion_count=1,
            evaluated_criteria_pct=50.0,
        )
        assert info.errored_criterion_count == 1
        assert info.evaluated_criteria_pct == 50.0

    def test_evaluation_info_errored_fields_default(self) -> None:
        info = EvaluationInfo(
            reward=1.0,
            raw_score=1.0,
            criterion_results=[
                CriterionResult(criterion="c1", weight=1.0, met=True, reasoning="ok"),
            ],
        )
        assert info.errored_criterion_count == 0
        assert info.evaluated_criteria_pct == 100.0

    def test_judge_input_model_copy(self) -> None:
        ji = JudgeInput(
            model="test-model",
            instructions="test",
            final_output="agent said done",
            criterion="check something",
            workdir="/workspace",
        )
        cloned = ji.model_copy(update={"workdir": "/new-workspace"})
        assert cloned.workdir == "/new-workspace"
        assert ji.workdir == "/workspace"

    def test_judge_input_serialization(self) -> None:
        ji = JudgeInput(
            model="test-model",
            instructions="test",
            final_output="agent said done",
            criterion="check something",
            workdir="/workspace",
            mcp_servers=[MCPServer(name="srv", command="/bin/srv")],
        )
        raw = ji.model_dump_json()
        restored = JudgeInput.model_validate_json(raw)
        assert restored.model == ji.model
        assert restored.final_output == ji.final_output
        assert len(restored.mcp_servers) == 1


class TestMutualExclusivity:
    """Verify that inline and path variants cannot both be set."""

    def base_kwargs(self) -> dict[str, Any]:
        return {
            "instructions": "test",
            "rubric_path": "/rubric.json",
            "workdir": "/workspace",
            "trajectory_path": "/logs/trajectory.json",
            "sandbox_user": "sandbox",
            "output_dir": "/logs/grader",
        }

    def test_instructions_inline_only(self) -> None:
        cfg = GraderConfig(**self.base_kwargs())
        assert cfg.instructions == "test"
        assert cfg.instructions_path is None

    def test_instructions_path_only(self) -> None:
        kw = self.base_kwargs()
        del kw["instructions"]
        cfg = GraderConfig(**kw, instructions_path="/some/instructions.md")
        assert cfg.instructions_path == "/some/instructions.md"
        assert cfg.instructions is None

    def test_instructions_both_raises(self) -> None:
        with pytest.raises(ValidationError, match="instructions"):
            GraderConfig(
                **self.base_kwargs(),
                instructions_path="/some/instructions.md",
            )

    def test_instructions_neither_is_valid(self) -> None:
        kw = self.base_kwargs()
        del kw["instructions"]
        cfg = GraderConfig(**kw)
        assert cfg.instructions is None
        assert cfg.instructions_path is None

    def test_rubric_inline_only(self) -> None:
        kw = self.base_kwargs()
        del kw["rubric_path"]
        cfg = GraderConfig(**kw, rubric=[RubricItem(criterion="c", weight=1.0)])
        assert cfg.rubric is not None
        assert cfg.rubric_path is None

    def test_rubric_inline_nested_normalised(self) -> None:
        kw = self.base_kwargs()
        del kw["rubric_path"]
        cfg = GraderConfig(
            **kw,
            rubric=RubricDocument(
                sections=[
                    RubricSection(
                        name="Model Build",
                        criteria=[RubricItem(criterion="Workbook exists", weight=0.0, gate=True)],
                    )
                ]
            ),
        )
        assert cfg.rubric is not None
        assert isinstance(cfg.rubric, list)
        assert cfg.rubric[0].section == "Model Build"
        assert cfg.rubric[0].gate is True

    def test_rubric_path_only(self) -> None:
        cfg = GraderConfig(**self.base_kwargs())
        assert cfg.rubric_path == "/rubric.json"
        assert cfg.rubric is None

    def test_rubric_both_raises(self) -> None:
        with pytest.raises(ValidationError, match="rubric"):
            GraderConfig(
                **self.base_kwargs(),
                rubric=[RubricItem(criterion="c", weight=1.0)],
            )

    def test_rubric_neither_raises(self) -> None:
        kw = self.base_kwargs()
        del kw["rubric_path"]
        with pytest.raises(ValidationError, match="rubric"):
            GraderConfig(**kw)

    def test_judge_guidance_inline_only(self) -> None:
        cfg = GraderConfig(**self.base_kwargs(), judge_guidance="inline text")
        assert cfg.judge_guidance == "inline text"
        assert cfg.judge_guidance_path is None

    def test_judge_guidance_path_only(self) -> None:
        cfg = GraderConfig(**self.base_kwargs(), judge_guidance_path="/some/file.md")
        assert cfg.judge_guidance_path == "/some/file.md"
        assert cfg.judge_guidance is None

    def test_judge_guidance_both_raises(self) -> None:
        with pytest.raises(ValidationError, match="judge_guidance"):
            GraderConfig(
                **self.base_kwargs(),
                judge_guidance="inline",
                judge_guidance_path="/some/file.md",
            )

    def test_judge_prompt_inline_only(self) -> None:
        cfg = GraderConfig(**self.base_kwargs(), judge_prompt="template text")
        assert cfg.judge_prompt == "template text"
        assert cfg.judge_prompt_path is None

    def test_judge_prompt_path_only(self) -> None:
        cfg = GraderConfig(**self.base_kwargs(), judge_prompt_path="/some/template.j2")
        assert cfg.judge_prompt_path == "/some/template.j2"
        assert cfg.judge_prompt is None

    def test_judge_prompt_both_raises(self) -> None:
        with pytest.raises(ValidationError, match="judge_prompt"):
            GraderConfig(
                **self.base_kwargs(),
                judge_prompt="inline",
                judge_prompt_path="/some/template.j2",
            )

    def test_neither_set_is_valid(self) -> None:
        cfg = GraderConfig(**self.base_kwargs())
        assert cfg.judge_guidance is None
        assert cfg.judge_guidance_path is None
        assert cfg.judge_prompt is None
        assert cfg.judge_prompt_path is None


class TestJudgePrompt:
    """Verify judge_prompt field on JudgeInput / BatchJudgeInput."""

    def test_judge_input_defaults_none(self) -> None:
        ji = JudgeInput(
            model="m",
            instructions="i",
            final_output="o",
            criterion="c",
            workdir="/w",
        )
        assert ji.judge_prompt is None

    def test_judge_input_roundtrip(self) -> None:
        ji = JudgeInput(
            model="m",
            instructions="i",
            final_output="o",
            criterion="c",
            workdir="/w",
            judge_prompt="Hello {{ instructions }}",
        )
        raw = ji.model_dump_json()
        restored = JudgeInput.model_validate_json(raw)
        assert restored.judge_prompt == "Hello {{ instructions }}"

    def test_batch_judge_input_defaults_none(self) -> None:
        bji = BatchJudgeInput(
            model="m",
            instructions="i",
            final_output="o",
            criteria=[],
            workdir="/w",
        )
        assert bji.judge_prompt is None

    def test_batch_judge_input_roundtrip(self) -> None:
        bji = BatchJudgeInput(
            model="m",
            instructions="i",
            final_output="o",
            criteria=[],
            workdir="/w",
            judge_prompt="Batch {{ n_max }}",
        )
        raw = bji.model_dump_json()
        restored = BatchJudgeInput.model_validate_json(raw)
        assert restored.judge_prompt == "Batch {{ n_max }}"


class TestGraderConfigMode:
    """Validate mode, batch_splits, and max_concurrency field constraints."""

    _DEFAULTS: ClassVar[dict[str, Any]] = {
        "instructions": "test",
        "rubric_path": "/rubric.json",
        "workdir": "/workspace",
        "trajectory_path": "/logs/trajectory.json",
        "output_dir": "/logs/grader",
    }

    def _cfg(self, **overrides: Any) -> GraderConfig:
        return GraderConfig(**{**self._DEFAULTS, **overrides})

    def test_mode_defaults_to_batch(self) -> None:
        assert self._cfg().mode == "batch"

    def test_mode_individual_accepted(self) -> None:
        assert self._cfg(mode="individual").mode == "individual"

    def test_mode_sequential_rejected(self) -> None:
        with pytest.raises(ValidationError):
            self._cfg(mode="sequential")

    def test_mode_invalid_value_rejected(self) -> None:
        with pytest.raises(ValidationError):
            self._cfg(mode="parallel")

    def test_section_filter_fields_rejected_in_config(self) -> None:
        with pytest.raises(ValidationError, match="Section filters are CLI-only"):
            self._cfg(include_sections=["Section A"])

        with pytest.raises(ValidationError, match="Section filters are CLI-only"):
            self._cfg(exclude_sections=["Section A"])

    # -- batch_splits --

    def test_batch_splits_defaults_none(self) -> None:
        assert self._cfg().batch_splits is None

    def test_batch_splits_2_accepted(self) -> None:
        assert self._cfg(mode="batch", batch_splits=2).batch_splits == 2

    def test_batch_splits_10_accepted(self) -> None:
        assert self._cfg(mode="batch", batch_splits=10).batch_splits == 10

    def test_batch_splits_1_rejected(self) -> None:
        """batch_splits must be >= 2 (splitting into 1 chunk is meaningless)."""
        with pytest.raises(ValidationError):
            self._cfg(mode="batch", batch_splits=1)

    def test_batch_splits_0_rejected(self) -> None:
        with pytest.raises(ValidationError):
            self._cfg(mode="batch", batch_splits=0)

    def test_batch_splits_negative_rejected(self) -> None:
        with pytest.raises(ValidationError):
            self._cfg(mode="batch", batch_splits=-1)

    def test_batch_splits_with_individual_mode_rejected(self) -> None:
        with pytest.raises(ValidationError, match=r"batch_splits.*batch"):
            self._cfg(mode="individual", batch_splits=3)

    # -- max_concurrency --

    def test_max_concurrency_defaults_none(self) -> None:
        assert self._cfg().max_concurrency is None

    def test_max_concurrency_1_accepted(self) -> None:
        assert self._cfg(max_concurrency=1).max_concurrency == 1

    def test_max_concurrency_0_rejected(self) -> None:
        with pytest.raises(ValidationError):
            self._cfg(max_concurrency=0)

    def test_max_concurrency_negative_rejected(self) -> None:
        with pytest.raises(ValidationError):
            self._cfg(max_concurrency=-1)

    # -- combined --

    def test_batch_splits_and_max_concurrency_independent(self) -> None:
        """batch_splits=10 with max_concurrency=2 is valid — 10 chunks, 2 parallel."""
        cfg = self._cfg(mode="batch", batch_splits=10, max_concurrency=2)
        assert cfg.batch_splits == 10
        assert cfg.max_concurrency == 2

    def test_individual_with_max_concurrency(self) -> None:
        cfg = self._cfg(mode="individual", max_concurrency=4)
        assert cfg.max_concurrency == 4
        assert cfg.batch_splits is None

    def test_adaptive_parallel_defaults(self) -> None:
        cfg = self._cfg()
        assert cfg.auto_parallel is True
        assert cfg.auto_parallel_target_chunk_size == 25
        assert cfg.rate_limit_discovery is True
        assert cfg.rate_limit_initial_concurrency == 2
        assert cfg.rate_limit_max_concurrency == 32
        assert cfg.rate_limit_backoff_seconds == 15
        assert cfg.rate_limit_max_requeues_per_chunk == 3

    def test_adaptive_parallel_explicit_values(self) -> None:
        cfg = self._cfg(
            auto_parallel=False,
            auto_parallel_target_chunk_size=10,
            rate_limit_discovery=False,
            rate_limit_initial_concurrency=4,
            rate_limit_max_concurrency=16,
            rate_limit_backoff_seconds=2,
            rate_limit_max_requeues_per_chunk=0,
        )
        assert cfg.auto_parallel is False
        assert cfg.auto_parallel_target_chunk_size == 10
        assert cfg.rate_limit_discovery is False
        assert cfg.rate_limit_initial_concurrency == 4
        assert cfg.rate_limit_max_concurrency == 16
        assert cfg.rate_limit_backoff_seconds == 2
        assert cfg.rate_limit_max_requeues_per_chunk == 0

    def test_batch_recovery_defaults(self) -> None:
        cfg = self._cfg()
        assert cfg.batch_recovery_enabled is True
        assert cfg.batch_recovery_target_chunk_size == 8
        assert cfg.batch_recovery_max_rounds == 2

    def test_batch_recovery_explicit_values(self) -> None:
        cfg = self._cfg(
            batch_recovery_enabled=False,
            batch_recovery_target_chunk_size=4,
            batch_recovery_max_rounds=0,
        )
        assert cfg.batch_recovery_enabled is False
        assert cfg.batch_recovery_target_chunk_size == 4
        assert cfg.batch_recovery_max_rounds == 0

    @pytest.mark.parametrize(
        ("field", "value"),
        [
            ("auto_parallel_target_chunk_size", 0),
            ("rate_limit_initial_concurrency", 0),
            ("rate_limit_max_concurrency", 0),
            ("rate_limit_backoff_seconds", 0),
            ("rate_limit_max_requeues_per_chunk", -1),
            ("batch_recovery_target_chunk_size", 0),
            ("batch_recovery_max_rounds", -1),
        ],
    )
    def test_adaptive_parallel_rejects_invalid_limits(self, field: str, value: int) -> None:
        with pytest.raises(ValidationError):
            self._cfg(**{field: value})
