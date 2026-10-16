"""python -m kd_mcp.servers.training_external [--transport stdio|streamable-http] ..."""
from __future__ import annotations

from ...runtime.transport import serve
from .server import DEFAULT_PORT, NAME, build_server


def main(argv: list[str] | None = None) -> None:
    serve(build_server, NAME, DEFAULT_PORT, argv)


if __name__ == "__main__":
    main()
