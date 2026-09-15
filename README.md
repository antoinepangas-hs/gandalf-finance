# gandalf-finance

This fork of [gandalf-the-grader](https://github.com/Handshake-AI-Research/gandalf-the-grader)
grades agent submissions for **ATLAS Finance**, a benchmark of multi-step
financial-modelling tasks whose deliverable is an Excel workbook. On top of
upstream it adds an Excel MCP server and deterministic workbook inspection, a
nested rubric schema with section gates, tolerance bands and negative-weight
penalty criteria, and orchestration options for grading a subset of sections or
running gates only.

The console scripts keep their upstream names (`gandalf-the-grader`,
`gandalf-the-grader-judge`), so existing verifier images work unchanged.

## What this fork adds

| module | purpose |
|---|---|
| `excel_mcp.py` | Excel MCP server (LibreOffice / Mac Excel / Windows-VM backends) |
| `workbook_digest.py` | deterministic workbook inspection |
| `paths.py` | work-root and protected-path helpers |
| `models.py` | nested rubric schema (`RubricDocument`), section gates, tolerance bands, negative-weight penalties |
| `orchestrator.py` | section include/exclude, gates-only runs, Excel backend wiring |
| `judge_batch.j2`, `judge_single.j2` | workbook-formatting guidance for the judge |

### Rubric schema

Upstream grades a flat list of binary criteria. This fork nests criteria under
sections, and a section can carry:

* **gates** — criteria that zero the whole section when unmet, so a model that
  fails a structural requirement cannot collect partial credit underneath it;
* **tolerance bands** — a relative percentage a numeric answer may differ by
  before it counts as wrong;
* **penalties** — negative weights, where the sign carries the semantics:
  positive means "reward when met", negative means "penalise when met".

## Install

```bash
# Not published to PyPI. Install from source, pinned to a tag:
uv pip install "gandalf-finance[pinned] @ git+https://github.com/antoinepangas-hs/gandalf-finance@v1.1.0"

gandalf-the-grader --help
```

`[pinned]` resolves the full transitive tree to exact `==` pins, which is what
makes a verifier image build reproducible; without it the judge's model-calling
dependencies float. It requires `HATCH_PINNED_EXTRA_ENABLE=1` at build time,
because the extra is emitted only when that is set.

## Licence

Apache-2.0, as upstream. See `LICENSE.txt`.
