# -*- coding: utf-8 -*-
"""
mcp_server/compat.py
MCP 1.x 및 MCP 2.x 완벽 호환성 레이어.
"""

from __future__ import annotations

try:
    from mcp.server.mcpserver import MCPServer
except ImportError:
    try:
        from mcp.server import MCPServer
    except ImportError:
        from mcp.server.fastmcp import FastMCP as MCPServer

FastMCP = MCPServer

__all__ = ["MCPServer", "FastMCP"]
