"""Runtime settings shared by every server.

Precedence: CLI arguments (runtime.transport) > environment variables > defaults.

Environment variables:
    KD_MCP_SKILLS_DIR     skills/ folder (default: the repo's skills/, i.e. an
                             editable install; deployments set it explicitly)
    KD_MCP_DATA_ROOT      where student files are read from; relative tool
                             paths resolve against it (default: cwd)
    KD_MCP_OUTPUT_DIR     the only folder tools may write to
                             (default: <data_root>/outputs)
    KD_MCP_ALLOWED_DIRS   extra read-only roots, separated by os.pathsep
    KD_MCP_AUDIT_LOG      audit log file (default:
                             <output_dir>/_audit/<server>.jsonl; "" disables)
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


def _path(v: str | os.PathLike) -> Path:
    return Path(v).expanduser().resolve()


def default_skills_dir() -> Path:
    # src/kd_mcp/core/config.py -> repo root is four levels up
    return (Path(__file__).resolve().parents[4] / "skills").resolve()


@dataclass(frozen=True)
class Settings:
    server_name: str
    skills_dir: Path
    data_root: Path
    output_dir: Path
    extra_read_dirs: tuple[Path, ...] = field(default_factory=tuple)
    audit_log: Path | None = None

    @classmethod
    def from_env(
        cls,
        server_name: str,
        *,
        skills_dir: str | os.PathLike | None = None,
        data_root: str | os.PathLike | None = None,
        output_dir: str | os.PathLike | None = None,
    ) -> "Settings":
        env = os.environ
        skills = _path(skills_dir or env.get("KD_MCP_SKILLS_DIR") or default_skills_dir())
        data = _path(data_root or env.get("KD_MCP_DATA_ROOT") or Path.cwd())
        out = _path(output_dir or env.get("KD_MCP_OUTPUT_DIR") or data / "outputs")
        extra = tuple(
            _path(p.strip()) for p in env.get("KD_MCP_ALLOWED_DIRS", "").split(os.pathsep) if p.strip()
        )
        audit_env = env.get("KD_MCP_AUDIT_LOG")
        if audit_env == "":
            audit = None
        elif audit_env:
            audit = _path(audit_env)
        else:
            audit = out / "_audit" / f"{server_name}.jsonl"
        return cls(server_name, skills, data, out, extra, audit)
