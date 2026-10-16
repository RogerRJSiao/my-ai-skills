"""Path whitelist. Every path that arrives from an MCP client goes through a
PathPolicy before it is opened.

Reads are allowed under data_root, the skills folder, the output folder (to
append to an earlier result file) and any extra read roots. Writes are
allowed only under the output folder. Relative paths resolve against
data_root. Symlinks and ".." are resolved before the check.
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


class PathNotAllowedError(PermissionError):
    """Raised when a requested path falls outside every allowed root."""


@dataclass(frozen=True)
class PathPolicy:
    base: Path
    read_roots: tuple[Path, ...]
    write_roots: tuple[Path, ...]

    @classmethod
    def build(cls, *, data_root: Path, output: Path, skills_dir: Path,
              extra_read_dirs: tuple[Path, ...] = ()) -> "PathPolicy":
        reads = []
        for r in (data_root, skills_dir, output, *extra_read_dirs):
            r = Path(r).resolve()
            if r not in reads:
                reads.append(r)
        return cls(Path(data_root).resolve(), tuple(reads), (Path(output).resolve(),))

    def _resolve(self, path: str | os.PathLike) -> Path:
        p = Path(path).expanduser()
        if not p.is_absolute():
            p = self.base / p
        return p.resolve()

    @staticmethod
    def _check(p: Path, roots: tuple[Path, ...], kind: str) -> None:
        if not any(p == r or p.is_relative_to(r) for r in roots):
            listed = ", ".join(str(r) for r in roots)
            raise PathNotAllowedError(f"{p} is outside the allowed {kind} directories ({listed}).")

    def read(self, path: str | os.PathLike, *, must_exist: bool = True) -> Path:
        p = self._resolve(path)
        self._check(p, self.read_roots, "read")
        if must_exist and not p.exists():
            raise FileNotFoundError(str(p))
        return p

    def write(self, path: str | os.PathLike) -> Path:
        p = self._resolve(path)
        self._check(p, self.write_roots, "write")
        return p
