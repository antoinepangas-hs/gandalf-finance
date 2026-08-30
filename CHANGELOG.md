# Changelog

## gandalf-finance 1.0.0

Fork of gandalf-the-grader for grading financial-modelling submissions.
Distribution renamed to `gandalf-finance`; the console scripts keep their
upstream names (`gandalf-the-grader`, `gandalf-the-grader-judge`) so existing
verifier images work unchanged.

Added over upstream: `excel_mcp.py` (Excel MCP server), `workbook_digest.py`
(deterministic workbook inspection), nested rubric schema with `section_gates`
and `section_tolerance_pct` on `RubricItem`, `judge_guidance_path` as a prompt
block separate from the task instructions, `--include-section`, and
workbook-formatting guidance in the judge templates.

Not published to PyPI; install from source or a vendored sdist.

---

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [1.0.0]

### Added

- Initial open-source release of Gandalf the Grader.

[1.0.0]: https://github.com/Handshake-AI-Research/gandalf-the-grader/releases/tag/v1.0.0
