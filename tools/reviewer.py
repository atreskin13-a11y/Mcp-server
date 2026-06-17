from __future__ import annotations

from pathlib import Path
from typing import Iterable, Set, Optional


def _iter_source_files(
    root_dir: Path,
    include_exts: Set[str],
    exclude_dirs: Set[str],
) -> Iterable[Path]:
    for path in root_dir.rglob("*"):
        if not path.is_file():
            continue
        if any(part in exclude_dirs for part in path.parts):
            continue
        if path.suffix.lower() in include_exts:
            yield path


def collect_project_code(
    root_dir: str,
    output_file: str,
    include_exts: Optional[Set[str]] = None,
    exclude_dirs: Optional[Set[str]] = None,
    add_line_numbers: bool = False,
) -> None:
    """
    Собирает код проекта в один файл.

    root_dir — корневая директория проекта.
    output_file — путь к результирующему файлу.
    include_exts — набор расширений ('.py', '.md', '.txt', ...).
    exclude_dirs — директории, которые нужно пропустить.
    add_line_numbers — если True, добавляет номера строк.
    """
    include_exts = include_exts or {".py"}
    exclude_dirs = exclude_dirs or {"__pycache__", ".git", "venv", "env"}

    root = Path(root_dir).resolve()
    out_path = Path(output_file).resolve()
    if not out_path.parent.exists():
        out_path.parent.mkdir(parents=True, exist_ok=True)

    parts = []

    for file_path in _iter_source_files(root, include_exts, exclude_dirs):
        rel = file_path.relative_to(root)
        header = f"\n\n# ===== File: {rel} =====\n"
        parts.append(header)

        try:
            text = file_path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            text = file_path.read_text(errors="ignore")

        if add_line_numbers:
            numbered_lines = []
            for idx, line in enumerate(text.splitlines(), start=1):
                numbered_lines.append(f"{idx:4}: {line}")
            text = "\n".join(numbered_lines)

        parts.append(text)

    out_path.write_text("".join(parts), encoding="utf-8")