# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/archetype_mapper.py
4대 오피스(HWP, DOCX, PPT, Excel) 정보 아키텍처 및 시각적 컴포넌트(Archetype) 매퍼.
원시 텍스트/데이터를 도구별 최적의 시각 구조로 변환합니다.
"""

from __future__ import annotations

from typing import Any


class ArchetypeMapper:
    """4대 오피스 시각적 아키타입 매핑 엔진"""

    # =========================================================================
    # 1. PowerPoint Archetypes
    # =========================================================================
    @staticmethod
    def map_ppt_slide(
        archetype: str,
        action_title: str,
        items: list[dict[str, Any]],
        footnote: str = "",
        category: str = "전략 보고",
    ) -> dict[str, Any]:
        """
        PPT 슬라이드 아키타입 매핑:
        - kpi_grid_3: 3개 핵심 지표 카드 (대형 숫자 + 부연 설명)
        - process_horizontal_4: 4단계 가로형 프로세스 쉐브론
        - comparison_2col: 좌우 대조 비교표/카드
        - card_grid_2x2: 4대 전략/사분면 그리드
        """
        valid_archetypes = ("kpi_grid_3", "process_horizontal_4", "comparison_2col", "card_grid_2x2", "table_matrix")
        arch = archetype if archetype in valid_archetypes else "kpi_grid_3"

        return {
            "slide_type": arch,
            "category_badge": category,
            "action_title": action_title,
            "items": items,
            "footnote": footnote,
            "layout_config": ArchetypeMapper._get_ppt_layout_config(arch),
        }

    @staticmethod
    def _get_ppt_layout_config(archetype: str) -> dict[str, Any]:
        configs = {
            "kpi_grid_3": {
                "columns": 3,
                "card_height": 280,
                "value_font_size": 40,
                "title_font_size": 18,
                "body_font_size": 13,
                "max_body_chars": 45,
            },
            "process_horizontal_4": {
                "columns": 4,
                "card_height": 300,
                "step_font_size": 14,
                "title_font_size": 16,
                "body_font_size": 12,
                "max_body_chars": 50,
            },
            "comparison_2col": {
                "columns": 2,
                "card_height": 350,
                "title_font_size": 20,
                "body_font_size": 14,
                "max_body_chars": 120,
            },
            "card_grid_2x2": {
                "columns": 2,
                "rows": 2,
                "card_height": 180,
                "title_font_size": 16,
                "body_font_size": 12,
                "max_body_chars": 60,
            },
            "table_matrix": {
                "header_bg": "002060",
                "alt_bg": "F2F2F2",
                "body_font_size": 11,
            },
        }
        return configs.get(archetype, {})

    # =========================================================================
    # 2. Excel Topology Archetypes
    # =========================================================================
    @staticmethod
    def map_excel_topology(
        title: str,
        kpi_tiles: list[dict[str, str]],
        headers: list[str],
        data_rows: list[list[Any]],
        summary_formulas: dict[str, str] | None = None,
        unit: str = "단위: 백만원, %",
    ) -> dict[str, Any]:
        """
        Excel 대시보드 2D 토폴로지 매핑:
        - Row 1: 여백/문서 타이틀
        - Row 2~4: 상단 3개 KPI 요약 타일 (KPI 카드)
        - Row 6: 섹션 타이틀 및 단위
        - Row 7: 테이블 헤더 (진한 테마색)
        - Row 8~N: 데이터 그리드
        - Row N+1: 합계/평균 수식 행
        """
        return {
            "topology_type": "executive_dashboard",
            "title": title,
            "unit": unit,
            "kpi_tiles": kpi_tiles[:4],  # 최대 4개
            "grid": {
                "start_row": 7,
                "start_col": 2,  # B열부터 시작하여 좌측 여백(A열 4px) 확보
                "headers": headers,
                "data_rows": data_rows,
                "summary_formulas": summary_formulas or {},
            },
            "view_settings": {
                "freeze_panes_row": 8,
                "freeze_panes_col": 2,
                "fit_to_page_wide": 1,
                "show_gridlines": True,
            },
        }

    # =========================================================================
    # 3. HWP One-Page Archetypes
    # =========================================================================
    @staticmethod
    def map_hwp_one_page(
        title: str,
        executive_summary: list[str],
        sections: list[dict[str, Any]],
        action_table: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """
        한글(HWP) A4 1장 보고서 아키타입:
        - 상단 1행 1열 요약 상자(Callout Table, 배경 연회색)
        - 본문 개조식 계층:
            □ (18pt Bold, 상하 6pt 여백)
            ○ (15pt Bold, 들여쓰기 10pt)
            - (13pt Regular, 내어쓰기 15pt 필수)
            * (11pt Regular, 참고사항)
        - 하단 세부 추진계획 표
        """
        return {
            "doc_archetype": "hwp_one_page_proposal",
            "title": title,
            "executive_summary_box": executive_summary[:3],  # 최대 3줄 요약
            "sections": sections,
            "action_table": action_table,
            "typography_budget": {
                "max_lines": 28,
                "body_font": "함초롬바탕",
                "line_spacing": 160,
                "hanging_indent_pt": 15,
                "keep_words": True,  # 어절 단위 줄바꿈
            },
        }

    # =========================================================================
    # 4. Word (DOCX) 4-Block Proposal
    # =========================================================================
    @staticmethod
    def map_docx_4block_proposal(
        title: str,
        author_info: dict[str, str],
        block1_summary: list[str],
        block2_comparison: dict[str, Any],
        block3_plan: dict[str, Any],
        block4_effects: list[str],
    ) -> dict[str, Any]:
        """
        워드(DOCX) 품의·기획 4대 표준 블록 매핑:
        - Block 1: Executive Summary (음영 박스 3줄 결론)
        - Block 2: 현행 대비 개선방안 대조표 (3:7 황금비 열 너비)
        - Block 3: 세부 실행 계획 및 일정표 (WBS)
        - Block 4: 기대효과 및 정량 지표
        """
        return {
            "doc_archetype": "docx_4block_proposal",
            "title": title,
            "author_info": author_info,
            "blocks": {
                "block1_summary": block1_summary,
                "block2_comparison": block2_comparison,
                "block3_plan": block3_plan,
                "block4_effects": block4_effects,
            },
            "styles": {
                "heading1_pt": 18,
                "heading2_pt": 14,
                "body_pt": 10.5,
                "callout_bg": "F4F6F9",
                "border_color": "002060",
            },
        }
