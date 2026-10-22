"""Engine lookup from skill.toml's `[grading] engine` / `engine_path`.

Every engine lives in its skill folder: engine_path (e.g. "engine") is a
Python package whose __init__ exposes ENGINE. Skill engines may import
kd_mcp (e.g. kd_mcp.core.grading.models).

Engines are imported lazily: an engine whose optional dependencies are not
installed (e.g. openpyxl without the `office` extra) is reported as
unavailable instead of stopping the whole server.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

from .base import Engine

# skill engines are registered in sys.modules under this prefix + engine name
SKILL_ENGINE_PREFIX = "kd_mcp_skill_engines"


class EngineUnavailableError(ValueError):
    pass


def engine_module(name: str, path: Path) -> ModuleType:
    """The engine's package module, loaded from its skill folder."""
    pkg = f"{SKILL_ENGINE_PREFIX}.{name}"
    if pkg in sys.modules:
        return sys.modules[pkg]
    init = Path(path) / "__init__.py"
    if not init.is_file():
        raise EngineUnavailableError(f"評分 engine「{name}」找不到 {init}")
    sys.modules.setdefault(SKILL_ENGINE_PREFIX, ModuleType(SKILL_ENGINE_PREFIX))
    spec = importlib.util.spec_from_file_location(pkg, init, submodule_search_locations=[str(path)])
    mod = importlib.util.module_from_spec(spec)
    sys.modules[pkg] = mod          # before exec, so the package's relative imports resolve
    try:
        spec.loader.exec_module(mod)
    except BaseException:
        sys.modules.pop(pkg, None)
        raise
    return mod


def get_engine(name: str, path: Path) -> Engine:
    try:
        return engine_module(name, path).ENGINE
    except ImportError as e:
        raise EngineUnavailableError(
            f"評分 engine「{name}」無法載入（{e}）；請確認已安裝對應的 extras，例如 pip install kd-mcp[office]"
        ) from e


def engine_status(name: str, path: Path) -> tuple[bool, str | None]:
    try:
        get_engine(name, path)
        return True, None
    except EngineUnavailableError as e:
        return False, str(e)
