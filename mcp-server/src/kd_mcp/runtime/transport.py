"""Command line + transport selection shared by every server's __main__.

    python -m kd_mcp.servers.<name> [--transport stdio|streamable-http]
        [--host 127.0.0.1] [--port N] [--data-root DIR] [--output-dir DIR]
        [--skills-dir DIR]
"""
from __future__ import annotations

import argparse
from collections.abc import Callable

from mcp.server.mcpserver import MCPServer
from mcp.server.transport_security import TransportSecuritySettings

from ..core.config import Settings


def build_arg_parser(prog: str, default_port: int) -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(prog=prog)
    ap.add_argument("--transport", choices=["stdio", "streamable-http"], default="stdio")
    ap.add_argument("--host", default="127.0.0.1",
                    help="streamable-http only; keep 127.0.0.1 behind IIS until runtime/auth exists")
    ap.add_argument("--port", type=int, default=default_port, help="streamable-http only")
    ap.add_argument("--allowed-host", action="append", default=[], metavar="HOST",
                    help="extra Host header to accept, e.g. the IIS site name mcp.example.local "
                         "(repeatable; localhost is always allowed)")
    ap.add_argument("--data-root", help="overrides KD_MCP_DATA_ROOT")
    ap.add_argument("--output-dir", help="overrides KD_MCP_OUTPUT_DIR")
    ap.add_argument("--skills-dir", help="overrides KD_MCP_SKILLS_DIR")
    return ap


def serve(
    build_server: Callable[[Settings], MCPServer],
    server_name: str,
    default_port: int,
    argv: list[str] | None = None,
) -> None:
    args = build_arg_parser(server_name, default_port).parse_args(argv)
    settings = Settings.from_env(
        server_name, skills_dir=args.skills_dir, data_root=args.data_root, output_dir=args.output_dir,
    )
    server = build_server(settings)
    if args.transport == "stdio":
        server.run("stdio")
    else:
        server.run("streamable-http", host=args.host, port=args.port,
                   transport_security=transport_security(args.allowed_host))


def transport_security(extra_hosts: list[str]) -> TransportSecuritySettings:
    """mcp only accepts localhost Host headers by default when bound to
    127.0.0.1 (DNS-rebinding protection). Behind IIS ARR the original Host
    header is forwarded, so the public site name must be allowed explicitly."""
    hosts = ["127.0.0.1:*", "localhost:*", "[::1]:*"]
    origins = ["http://127.0.0.1:*", "http://localhost:*", "http://[::1]:*"]
    for h in extra_hosts:
        hosts += [h, f"{h}:*"]
        origins += [f"https://{h}", f"http://{h}"]
    return TransportSecuritySettings(
        enable_dns_rebinding_protection=True, allowed_hosts=hosts, allowed_origins=origins,
    )
