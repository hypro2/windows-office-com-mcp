# -*- coding: utf-8 -*-
"""
mcp_server/tools/folder_tools.py
사용자 실무 문서 폴더 벌크 학습 및 기존 도구 미세조정 & 신규 도구 합성 MCP 도구 세트.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from mcp_server.compat import MCPServer

from .folder_inspector import FolderInspector

logger = logging.getLogger("mcp_server.folder_tools")


def register_folder_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def office_learn_folder(
        folder_path: str,
        profile_name: str = "my_office_profile",
        synthesize_new_tools: bool = True,
        recursive: bool = False,
    ) -> dict[str, Any]:
        """
        [사용자 실무 폴더 벌크 학습 & 기존 도구 미세조정]
        자주 사용하는 오피스 문서(PPTX, XLSX, DOCX, HWP)가 모인 폴더를 전수 스캔(COM 개방/확인/종료 반복)하여:
        1. 대표 브랜드 컬러, 폰트 위계, 카드 모서리 곡률, 여백 규격을 추출해 '기존 70개 공식 도구'의 서식을 맞춤 미세조정합니다.
        2. 사내 고유 서식이 발견되면 '신규 슬라이스/도구 레시피(custom_slice_<id>)'를 자동 합성하여 즉시 사용 가능하도록 등록합니다.

        - folder_path: 자주 쓰는 문서들이 들어 있는 폴더 경로 (예: "C:/docs/자주쓰는문서들")
        - profile_name: 저장할 프로필 이름 (예: "전략기획팀_표준", "여신심사_서식")
        - synthesize_new_tools: 고유 서식 발견 시 신규 도구 자동 합성 여부 (기본 True)
        - recursive: 하위 폴더까지 재귀적으로 스캔할지 여부 (기본 False)
        """
        try:
            fpath = Path(folder_path)
            profile = FolderInspector.inspect_and_learn_folder(
                folder_path=fpath,
                profile_name=profile_name,
                synthesize_new_tools=synthesize_new_tools,
                recursive=recursive,
            )

            synth_count = len(profile.get("synthesized_tools", []))
            return {
                "success": True,
                "profile_name": profile_name,
                "total_documents_analyzed": profile.get("total_documents_analyzed", 0),
                "palette": profile.get("palette", {}),
                "fonts": profile.get("fonts", {}),
                "synthesized_new_tools_count": synth_count,
                "synthesized_tools": profile.get("synthesized_tools", []),
                "message": (
                    f"폴더 '{fpath.name}'의 {profile.get('total_documents_analyzed', 0)}개 문서를 성공적으로 학습했습니다! "
                    f"기존 도구들이 프로필 '{profile_name}'의 폰트/컬러 규격으로 미세조정되었으며, "
                    f"{synth_count}개의 고유 서식이 신규 도구(레시피)로 자동 합성되었습니다."
                ),
            }
        except Exception as e:
            logger.error(f"폴더 벌크 학습 실패: {e}", exc_info=True)
            return {"success": False, "error": f"폴더 벌크 학습 실패: {e}"}

    @mcp.tool()
    def office_list_profiles() -> dict[str, Any]:
        """
        [학습된 서식 프로필 및 신규 합성 도구 목록 조회]
        현재 등록된 사용자 서식 프로필 목록과 활성(Active) 상태, 자동 합성된 신규 도구 목록을 반환합니다.
        """
        try:
            profiles = FolderInspector.list_all_profiles()
            active_info = FolderInspector.get_active_profile()
            return {
                "success": True,
                "active_profile": active_info.get("profile_name") if active_info else "none (기본 테마 적용 중)",
                "total_profiles": len(profiles),
                "profiles": profiles,
            }
        except Exception as e:
            return {"success": False, "error": f"프로필 목록 조회 실패: {e}"}

    @mcp.tool()
    def office_set_active_profile(profile_name: str) -> dict[str, Any]:
        """
        [활성 서식 프로필 전환 및 기존 도구 동기화]
        지정한 프로필을 모든 기존 도구의 기본 서식 스타일로 전환 적용합니다.
        """
        try:
            profile = FolderInspector.set_active_profile(profile_name)
            return {
                "success": True,
                "active_profile": profile_name,
                "palette": profile.get("palette", {}),
                "fonts": profile.get("fonts", {}),
                "message": f"활성 프로필이 '{profile_name}'으로 전환되었습니다. 이제 기존 도구들이 이 규격으로 작동합니다.",
            }
        except Exception as e:
            return {"success": False, "error": f"프로필 전환 실패: {e}"}

    @mcp.tool()
    def office_show_tuning_status() -> dict[str, Any]:
        """
        [기존 도구 미세조정 상태 확인]
        현재 기존 70개 공식 MCP 도구에 어떤 사용자 폰트, 브랜드 컬러 팔레트, 카드 여백 규격이
        미세조정되어 적용 중인지 실시간 상태를 확인합니다.
        """
        try:
            active_info = FolderInspector.get_active_profile()
            if not active_info:
                return {
                    "success": True,
                    "status": "기본 프리셋(business_navy) 적용 중",
                    "guide": "'office_learn_folder'를 호출해 자주 쓰는 폴더를 학습하면 나만의 서식 규격으로 자동 미세조정됩니다.",
                }

            return {
                "success": True,
                "active_profile": active_info.get("profile_name"),
                "applied_fonts": active_info.get("fonts", {}),
                "applied_palette": active_info.get("palette", {}),
                "applied_layout_tokens": active_info.get("layout_tokens", {}),
                "synthesized_tools_available": active_info.get("synthesized_tools", []),
            }
        except Exception as e:
            return {"success": False, "error": f"상태 확인 실패: {e}"}
