"""Tests for macOS-aware filesystem location helpers."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

from gandalf import paths


def _force_darwin(monkeypatch: pytest.MonkeyPatch, home: Path) -> None:
    monkeypatch.setattr(sys, "platform", "darwin")
    monkeypatch.setenv("HOME", str(home))


def test_default_work_root_is_under_home(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("HOME", str(tmp_path))
    assert paths.default_work_root() == tmp_path / "gandalf_runs"


def test_ensure_work_root_creates_directory(tmp_path: Path) -> None:
    root = tmp_path / "nested" / "gandalf_runs"
    assert paths.ensure_work_root(root) == root
    assert root.is_dir()


def test_ensure_work_root_rejects_unwritable(tmp_path: Path) -> None:
    root = tmp_path / "ro_root"
    root.mkdir()
    root.chmod(0o500)
    try:
        with pytest.raises(RuntimeError, match="not writable"):
            paths.ensure_work_root(root)
    finally:
        root.chmod(0o700)


@pytest.mark.parametrize(
    "subdir",
    ["Documents", "Desktop", "Downloads", "Library/Mobile Documents", "Library/CloudStorage", "Dropbox"],
)
def test_is_protected_path_flags_protected_dirs(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, subdir: str
) -> None:
    _force_darwin(monkeypatch, tmp_path)
    target = tmp_path / subdir / "BTBv2" / "workdir"
    target.mkdir(parents=True)
    assert paths.is_protected_path(target) is True


def test_is_protected_path_allows_neutral_dirs(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _force_darwin(monkeypatch, tmp_path)
    neutral = tmp_path / "gandalf_runs" / "run1" / "workdir"
    neutral.mkdir(parents=True)
    assert paths.is_protected_path(neutral) is False


def test_is_protected_path_flags_volumes(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _force_darwin(monkeypatch, tmp_path)
    assert paths.is_protected_path("/Volumes/ExternalDrive/run/workdir") is True


def test_is_protected_path_noop_off_darwin(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(sys, "platform", "linux")
    monkeypatch.setenv("HOME", str(tmp_path))
    docs = tmp_path / "Documents" / "workdir"
    docs.mkdir(parents=True)
    assert paths.is_protected_path(docs) is False
