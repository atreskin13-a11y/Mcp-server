from __future__ import annotations

from pathlib import Path
from typing import List


def list_python_files_full_paths(directory: str | Path) -> List[str]:
    """
    Рекурсивно возвращает список абсолютных путей ко всем .py файлам
    в указанной директории.
    """
    root = Path(directory).resolve()
    if not root.exists():
        return []
    result: List[str] = []
    for path in root.rglob("*.py"):
        if path.is_file():
            result.append(str(path.resolve()))
    return result


def _ensure_parent_dir(path: Path) -> None:
    if not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)


def write_file(
    file_path: str,
    content: str | bytes,
    encoding: str = "utf-8",
    mode: str = "w",
    create_parents: bool = True,
) -> None:
    """
    Записывает указанный контент в файл.
    Если create_parents=True, создаёт все недостающие директории.
    """
    path = Path(file_path)
    if create_parents:
        _ensure_parent_dir(path)

    if isinstance(content, bytes):
        with path.open("wb") as f:
            f.write(content)
        return

    with path.open(mode, encoding=encoding) as f:
        f.write(str(content))


def read_file_txt(
    file_path: str,
    encoding: str = "utf-8",
    mode: str = "r",
) -> str:
    """
    Читает текстовый файл и возвращает его содержимое.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")
    with path.open(mode, encoding=encoding) as f:
        return f.read()