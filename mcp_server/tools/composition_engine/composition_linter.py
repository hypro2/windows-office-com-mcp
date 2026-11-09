# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/composition_linter.py
문서 구성 완성도 린터(Linter) 및 자가 교정기.
정보 위계, Action Title 결론성, 엑셀 수식 무결성, 개조식 정합성을 점검합니다.
"""

from __future__ import annotations

import re
from typing import Any


class CompositionLinter:
    """4대 오피스 구성 품질 린터"""

    # -------------------------------------------------------------------------
    # PPT Composition Linting
    # -------------------------------------------------------------------------
    @classmethod
    def lint_ppt_slide(cls, slide_spec: dict[str, Any]) -> dict[str, Any]:
        issues = []
        auto_fixes = []

        action_title = slide_spec.get("action_title", "").strip()

        # 1. Action Title 결론성 검사 (단순 명사형 제목 방지: 개요, 현황, 분석, 전략 등)
        weak_patterns = [
            r"^[\d\.\s]*개요$",
            r"^[\d\.\s]*목차$",
            r".*(현황|분석|전략|방안|추진)$",
        ]
        for pat in weak_patterns:
            if re.search(pat, action_title):
                issues.append(
                    f"제목 '{action_title}'은 단순 명사형입니다. "
                    f"임원 보고용 슬라이드는 '핵심 수치 + 결론'을 담은 서술형 Action Title이어야 합니다."
                )
                auto_fixes.append(f"{action_title} → '{action_title} 결과 및 핵심 시사점 도출'")
                break

        # 2. 아이템 개수 밸런스 검사
        items = slide_spec.get("items", [])
        slide_type = slide_spec.get("slide_type", "")
        if slide_type == "kpi_grid_3" and len(items) != 3:
            issues.append(f"kpi_grid_3 아키타입은 정확히 3개 카드가 최적입니다 (현재 {len(items)}개).")
        elif slide_type == "process_horizontal_4" and len(items) != 4:
            issues.append(f"process_horizontal_4 아키타입은 4단계 프로세스가 최적입니다 (현재 {len(items)}개).")

        return {
            "valid": len(issues) == 0,
            "issues": issues,
            "auto_fixes": auto_fixes,
        }

    # -------------------------------------------------------------------------
    # Excel Composition Linting
    # -------------------------------------------------------------------------
    @classmethod
    def lint_excel_topology(cls, spec: dict[str, Any]) -> dict[str, Any]:
        issues = []
        auto_fixes = []

        grid = spec.get("grid", {})
        data_rows = grid.get("data_rows", [])
        formulas = grid.get("summary_formulas", {})

        # 1. 수식 누락 검사 (실무자 고충: 하드코딩된 숫자 방지)
        if len(data_rows) > 0 and len(formulas) == 0:
            issues.append(
                "데이터 테이블 하단에 집계 수식(SUM, AVERAGE)이 지정되지 않았습니다. "
                "실무 엑셀은 계산된 정적 숫자가 아닌 네이티브 수식이 연결되어야 합니다."
            )
            # B열부터 마지막 열까지 자동 SUM 제안
            auto_fixes.append("마지막 행에 '=SUM(Col8:ColN)' 자동 주입 권장")

        # 2. 숫자 포맷팅 검사 (천단위 콤마)
        headers = grid.get("headers", [])
        numeric_col_found = False
        for r in data_rows:
            for cell in r:
                if isinstance(cell, (int, float)) and cell > 1000:
                    numeric_col_found = True
                    break

        return {
            "valid": len(issues) == 0,
            "issues": issues,
            "auto_fixes": auto_fixes,
            "numeric_formatting_recommended": numeric_col_found,
        }

    # -------------------------------------------------------------------------
    # HWP Composition Linting
    # -------------------------------------------------------------------------
    @classmethod
    def lint_hwp_one_page(cls, spec: dict[str, Any]) -> dict[str, Any]:
        issues = []
        auto_fixes = []

        # 1. 상단 요약 상자 존재 여부
        sum_box = spec.get("executive_summary_box", [])
        if not sum_box:
            issues.append("상단 요약 상자(Executive Summary Box)가 누락되었습니다. 공문서는 두괄식 상자가 필수입니다.")

        # 2. 개조식 위계 순서 점검 (□ -> ○ -> - -> *)
        sections = spec.get("sections", [])
        for s_idx, sec in enumerate(sections, start=1):
            items = sec.get("items", [])
            for i_idx, it in enumerate(items, start=1):
                level = it.get("level", 1)
                symbol = it.get("symbol", "")
                expected_symbol = {1: "□", 2: "○", 3: "-", 4: "*"}.get(level, "-")
                if symbol and symbol != expected_symbol:
                    issues.append(
                        f"섹션 #{s_idx} 항목 #{i_idx}의 기호 '{symbol}'가 레벨 {level} 표준('{expected_symbol}')과 다릅니다."
                    )
                    auto_fixes.append(f"기호를 '{expected_symbol}'로 자동 보정 권장")

        return {
            "valid": len(issues) == 0,
            "issues": issues,
            "auto_fixes": auto_fixes,
        }

    # -------------------------------------------------------------------------
    # DOCX Composition Linting
    # -------------------------------------------------------------------------
    @classmethod
    def lint_docx_proposal(cls, spec: dict[str, Any]) -> dict[str, Any]:
        issues = []
        blocks = spec.get("blocks", {})

        required_blocks = ["block1_summary", "block2_comparison", "block3_plan", "block4_effects"]
        for rb in required_blocks:
            if not blocks.get(rb):
                issues.append(f"비즈니스 품의서 필수 블록 '{rb}'가 누락되었습니다.")

        return {
            "valid": len(issues) == 0,
            "issues": issues,
        }
