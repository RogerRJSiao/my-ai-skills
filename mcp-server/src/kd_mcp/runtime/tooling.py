"""Wrap tool functions: audit every call and turn anticipated failures into
ToolError so their message reaches the client (anything else is a crash and
the client only sees the SDK's generic 'Error executing tool')."""
from __future__ import annotations

import functools
import time
from collections.abc import Callable

from mcp.server.mcpserver.exceptions import ToolError

from ..core.logging import AuditLogger
from ..core.security import PathNotAllowedError

ANTICIPATED = (PathNotAllowedError, FileNotFoundError, FileExistsError,
               NotADirectoryError, ValueError, KeyError)


def current_user() -> str:
    # stdio = the local user; runtime/auth.py will map HTTP tokens to user IDs.
    return "local"


def audited(tool_name: str, audit: AuditLogger | None,
            user: Callable[[], str] = current_user):
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            t0 = time.perf_counter()

            def log(status: str, error: str | None = None):
                if audit:
                    audit.record(tool=tool_name, user=user(), status=status,
                                 duration_ms=int((time.perf_counter() - t0) * 1000),
                                 args=kwargs, error=error)

            try:
                result = fn(*args, **kwargs)
            except ANTICIPATED as e:
                log("rejected", f"{type(e).__name__}: {e}")
                raise ToolError(f"{type(e).__name__}: {e}") from e
            except Exception as e:
                log("crashed", f"{type(e).__name__}: {e}")
                raise
            log("ok")
            return result
        return wrapper
    return deco
