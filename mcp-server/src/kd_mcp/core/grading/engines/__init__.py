"""Engine lookup by the `[grading] engine` name in skill.toml.

Engines are imported lazily: an engine whose optional dependencies are not
installed (e.g. openpyxl without the `office` extra) is reported as
unavailable instead of stopping the whole server.
"""
from __future__ import annotations

import importlib

from .base import Engine

# engine name -> module that exposes ENGINE
ENGINE_MODULES: dict[str, str] = {
    "xlsx_s06": f"{__name__}.xlsx_s06",
}


class EngineUnavailableError(ValueError):
    pass


def known_engines() -> set[str]:
    return set(ENGINE_MODULES)


def get_engine(name: str) -> Engine:
    module = ENGINE_MODULES.get(name)
    if module is None:
        raise EngineUnavailableError(f"未知的評分 engine「{name}」")
    try:
        return importlib.import_module(module).ENGINE
    except ImportError as e:
        raise EngineUnavailableError(
            f"評分 engine「{name}」無法載入（{e}）；請確認已安裝對應的 extras，例如 pip install kd-mcp[office]"
        ) from e


def engine_status(name: str) -> tuple[bool, str | None]:
    try:
        get_engine(name)
        return True, None
    except EngineUnavailableError as e:
        return False, str(e)
