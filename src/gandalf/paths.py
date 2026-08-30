"""Filesystem location helpers for avoiding macOS Excel sandbox prompts.

Microsoft Excel for Mac runs in the App Sandbox. When it is driven via
automation (xlwings) to open a workbook in a TCC/sandbox-protected location, it
shows a "Grant Access" powerbox dialog. Keeping run artefacts out of those
locations avoids the prompt entirely. These helpers are macOS-aware and are
no-ops on other platforms.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

DEFAULT_WORK_ROOT_NAME = "gandalf_runs"

# Locations macOS gates behind the App Sandbox powerbox / TCC. Excel prompts to
# "Grant Access" for files opened here via automation. Third-party cloud-sync
# roots are included because they exhibit the same behaviour.
_PROTECTED_HOME_SUBDIRS = (
    "Documents",
    "Desktop",
    "Downloads",
    "Library/Mobile Documents",  # iCloud Drive
    "Library/CloudStorage",  # modern OneDrive / Google Drive / Dropbox
    "Dropbox",
    "Google Drive",
    "OneDrive",
)


def default_work_root() -> Path:
    """Return the default non-protected root for run artefacts (``~/gandalf_runs``)."""
    return Path.home() / DEFAULT_WORK_ROOT_NAME


def excel_access_root() -> Path:
    """Return the stable root that Microsoft Excel opens workbooks/clones from.

    All Mac-Excel reads (repair check, preflight, judge workspace clones) are
    routed through per-run subdirectories of this single folder. macOS makes a
    powerbox "Grant Access" grant persistent and folder-wide, so granting this
    one folder once stops Excel prompting on subsequent runs. Using a *stable*
    parent is what makes the grant carry over — random system temp dirs do not.
    """
    return default_work_root() / "excel_access"


def ensure_work_root(root: Path) -> Path:
    """Create *root* if needed and return it, raising if it is not usable."""
    try:
        root.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        msg = f"Work root {str(root)!r} could not be created: {e}"
        raise RuntimeError(msg) from e
    if not os.access(root, os.W_OK):
        msg = f"Work root {str(root)!r} is not writable"
        raise RuntimeError(msg)
    return root


def is_protected_path(path: Path | str) -> bool:
    """Return whether *path* is in a macOS sandbox/TCC-protected location.

    Only meaningful on macOS; returns ``False`` on other platforms so Linux/CI
    runs are unaffected. Symlinks are resolved before matching.
    """
    if sys.platform != "darwin":
        return False
    resolved = Path(path).expanduser().resolve()
    home = Path.home().resolve()
    bases = [(home / sub).resolve() for sub in _PROTECTED_HOME_SUBDIRS]
    bases.append(Path("/Volumes"))  # external / network volumes
    return any(resolved.is_relative_to(base) for base in bases)
