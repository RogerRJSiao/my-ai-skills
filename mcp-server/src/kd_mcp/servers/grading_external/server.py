"""kd-grading-external — 外訓評分 server (prefix "ge", domain "training_external")."""
from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from ... import __version__
from ...core.config import Settings
from ...core.logging import AuditLogger
from ...toolkits import grading

NAME = "kd-grading-external"
DOMAIN = "training_external"
PREFIX = "ge"
DEFAULT_PORT = 8102


def build_server(settings: Settings | None = None) -> MCPServer:
    settings = settings or Settings.from_env(NAME)
    mcp = MCPServer(NAME, instructions=grading.instructions(PREFIX, DOMAIN), version=__version__)
    grading.register(
        mcp,
        domain=DOMAIN,
        prefix=PREFIX,
        data_root=settings.data_root,
        output=settings.output_dir,
        skills_dir=settings.skills_dir,
        extra_read_dirs=settings.extra_read_dirs,
        audit=AuditLogger(settings.audit_log, NAME),
    )
    return mcp
