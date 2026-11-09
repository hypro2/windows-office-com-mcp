"""
Standard JSON Response Envelope (Inspired by GenOffice CLI & Astryx)

Ensures every tool and command produces consistent, machine-readable JSON envelopes
with actionable suggestions for AI Agents to enable self-healing workflows.
"""

from typing import Dict, Any, Optional, List
import json


def create_envelope(
    status: str,  # "ok" | "partial" | "error"
    command: str,
    summary: str,
    output_path: Optional[str] = None,
    suggestion: Optional[str] = None,
    warnings: Optional[List[str]] = None,
    detail: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Constructs a standardized JSON envelope conforming to GenOffice agent standards.
    """
    payload = {
        "status": status,
        "command": command,
        "summary": summary,
    }
    if output_path is not None:
        payload["output_path"] = output_path
    if suggestion is not None:
        payload["suggestion"] = suggestion
    if warnings:
        payload["warnings"] = warnings
    if detail is not None:
        payload["detail"] = detail

    return payload


def ok_response(
    command: str,
    summary: str,
    output_path: Optional[str] = None,
    detail: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    return create_envelope(
        status="ok",
        command=command,
        summary=summary,
        output_path=output_path,
        detail=detail or {}
    )


def error_response(
    command: str,
    summary: str,
    suggestion: str,
    code: str = "ERR_GENERIC",
    detail: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    d = detail or {}
    d["code"] = code
    return create_envelope(
        status="error",
        command=command,
        summary=summary,
        suggestion=suggestion,
        detail=d
    )
