"""kd-training-internal — 內訓評分 server (prefix "ti", domain "training_internal")."""
from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from ... import __version__
from ...core.config import Settings
from ...core.logging import AuditLogger
from ...toolkits import training

NAME = "kd-training-internal"
DOMAIN = "training_internal"
PREFIX = "ti"
DEFAULT_PORT = 8101


def build_server(settings: Settings | None = None) -> MCPServer:
    settings = settings or Settings.from_env(NAME)
    mcp = MCPServer(NAME, instructions=training.instructions(PREFIX, DOMAIN), version=__version__)
    training.register(
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
