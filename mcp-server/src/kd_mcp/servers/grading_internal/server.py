"""kd-grading-internal — 內訓評分 server (prefix "gi", domain "training_internal")."""
from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from ... import __version__
from ...core.config import Settings
from ...core.logging import AuditLogger
from ...toolkits import grading

NAME = "kd-grading-internal"
DOMAIN = "training_internal"
PREFIX = "gi"
DEFAULT_PORT = 8101

#--組裝 server
def build_server(settings: Settings | None = None) -> MCPServer:
    settings = settings or Settings.from_env(NAME)
    #--建立一個空的 MCP server
    mcp = MCPServer(NAME, instructions=grading.instructions(PREFIX, DOMAIN), version=__version__)
    #--把 tools 掛到 server 上，對應 toolkits 的 grading.py
    #--register() 執行完，6 個 gi_* tool 就登記到 server 上了，接著 transport.py 才開始接收 Claude 的請求
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
