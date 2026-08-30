# gandalf-finance

A fork of [gandalf-the-grader](https://github.com/Handshake-AI-Research/gandalf-the-grader)
for grading financial-modelling submissions.

## What this fork adds

| module | purpose |
|---|---|
| `excel_mcp.py` | Excel MCP server (LibreOffice / Mac Excel / Windows-VM backends) |
| `workbook_digest.py` | deterministic workbook inspection |
| `paths.py` | work-root and protected-path helpers |
| `models.py` | nested rubric schema (`RubricDocument`), section gates and tolerances |
| `orchestrator.py` | section include/exclude, gates-only runs, Excel backend wiring |
| `judge_batch.j2`, `judge_single.j2` | workbook-formatting guidance for the judge |

## Install

```bash
# Not published to PyPI. Install from source:
uv pip install git+https://github.com/antoinepangas-hs/gandalf-finance
# or from a vendored sdist:
uv pip install ./gandalf_finance-1.0.0.tar.gz

gandalf-the-grader --help
```

The console scripts keep their upstream names (`gandalf-the-grader`,
`gandalf-the-grader-judge`) so existing verifier images work unchanged.

## Licence

Apache-2.0, as upstream. See `LICENSE.txt`.
