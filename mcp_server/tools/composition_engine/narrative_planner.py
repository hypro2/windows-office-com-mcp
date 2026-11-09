# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/narrative_planner.py
자연어 요청 기반 문서 서사(Narrative) 및 정보 아키텍처 자동 설계기.
사용자의 주제와 거친 데이터를 실무 규격에 맞는 정형화된 Blueprint로 기획합니다.
"""

from __future__ import annotations

import logging
from typing import Any

from .archetype_mapper import ArchetypeMapper
from .spatial_budgeter import SpatialBudgeter
from .composition_linter import CompositionLinter
from .auto_composer import AutoComposer

logger = logging.getLogger("mcp_server.composition_engine")


class NarrativePlanner:
    """4대 오피스 문서 서사 및 정보 아키텍처 기획기"""

    @classmethod
    def auto_compose_from_text(
        cls,
        raw_text: str,
        presentation_title: str | None = None,
        theme: str | None = None,
    ) -> dict[str, Any]:
        """원시 텍스트 덤프 기반 PPT 자동 기획·구성"""
        return AutoComposer.compose(raw_text, presentation_title=presentation_title, theme=theme)

    @classmethod
    def plan_ppt(
        cls,
        topic: str,
        slides_input: list[dict[str, Any]],
        theme: str = "navy_clean",
    ) -> dict[str, Any]:
        """
        PPT 전체 덱 스토리라인 및 슬라이드별 아키타입 기획
        """
        planned_slides = []
        all_warnings = []

        for idx, s in enumerate(slides_input, start=1):
            arch = s.get("archetype", "kpi_grid_3")
            action_title = s.get("action_title", f"{topic}의 핵심 성과 분석")
            items = s.get("items", [])
            footnote = s.get("footnote", "")
            category = s.get("category", "임원 보고")

            slide_spec = ArchetypeMapper.map_ppt_slide(
                archetype=arch,
                action_title=action_title,
                items=items,
                footnote=footnote,
                category=category,
            )

            # 1. 공간 예산 검사
            budget_res = SpatialBudgeter.audit_ppt_slide(slide_spec)
            if not budget_res["passed"]:
                all_warnings.extend([f"슬라이드 #{idx}: {w}" for w in budget_res["warnings"]])

            # 2. 린터 검사
            lint_res = CompositionLinter.lint_ppt_slide(slide_spec)
            if not lint_res["valid"]:
                all_warnings.extend([f"슬라이드 #{idx}: {i}" for i in lint_res["issues"]])

            planned_slides.append(slide_spec)

        return {
            "doc_type": "ppt",
            "topic": topic,
            "theme": theme,
            "slide_count": len(planned_slides),
            "slides": planned_slides,
            "audit_warnings": all_warnings,
            "quality_status": "EXCELLENT" if not all_warnings else "NEEDS_ATTENTION",
        }

    @classmethod
    def plan_excel(
        cls,
        title: str,
        kpi_tiles: list[dict[str, str]],
        headers: list[str],
        data_rows: list[list[Any]],
        summary_formulas: dict[str, str] | None = None,
        unit: str = "단위: 백만원, %",
    ) -> dict[str, Any]:
        """
        Excel 대시보드 2D 토폴로지 기획
        """
        spec = ArchetypeMapper.map_excel_topology(
            title=title,
            kpi_tiles=kpi_tiles,
            headers=headers,
            data_rows=data_rows,
            summary_formulas=summary_formulas,
            unit=unit,
        )

        budget_res = SpatialBudgeter.audit_excel_topology(spec)
        lint_res = CompositionLinter.lint_excel_topology(spec)

        warnings = list(budget_res.get("warnings", [])) + list(lint_res.get("issues", []))

        return {
            "doc_type": "excel",
            "spec": spec,
            "audit_warnings": warnings,
            "quality_status": "EXCELLENT" if not warnings else "NEEDS_ATTENTION",
        }

    @classmethod
    def plan_hwp_one_page(
        cls,
        title: str,
        executive_summary: list[str],
        sections: list[dict[str, Any]],
        action_table: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """
        한글 A4 1장 보고서 기획 및 예산 점검
        """
        spec = ArchetypeMapper.map_hwp_one_page(
            title=title,
            executive_summary=executive_summary,
            sections=sections,
            action_table=action_table,
        )

        budget_res = SpatialBudgeter.audit_hwp_one_page(spec)
        lint_res = CompositionLinter.lint_hwp_one_page(spec)

        warnings = []
        if not budget_res["passed"]:
            warnings.append(budget_res["recommendation"])
        warnings.extend(lint_res.get("issues", []))

        return {
            "doc_type": "hwp",
            "spec": spec,
            "budget_details": budget_res,
            "audit_warnings": warnings,
            "quality_status": "EXCELLENT" if not warnings else "NEEDS_ATTENTION",
        }

    @classmethod
    def plan_docx_proposal(
        cls,
        title: str,
        author_info: dict[str, str],
        block1_summary: list[str],
        block2_comparison: dict[str, Any],
        block3_plan: dict[str, Any],
        block4_effects: list[str],
    ) -> dict[str, Any]:
        """
        워드 4대 표준 블록 품의·기획서 기획
        """
        spec = ArchetypeMapper.map_docx_4block_proposal(
            title=title,
            author_info=author_info,
            block1_summary=block1_summary,
            block2_comparison=block2_comparison,
            block3_plan=block3_plan,
            block4_effects=block4_effects,
        )

        lint_res = CompositionLinter.lint_docx_proposal(spec)
        warnings = lint_res.get("issues", [])

        return {
            "doc_type": "docx",
            "spec": spec,
            "audit_warnings": warnings,
            "quality_status": "EXCELLENT" if not warnings else "NEEDS_ATTENTION",
        }
