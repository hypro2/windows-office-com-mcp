# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/spatial_budgeter.py
공간 예산제(Spatial Budgeting) 및 시각적 제약 집행기.
텍스트 넘침(Overflow), A4 1페이지 초과, 슬라이드 과밀도를 사전 방지합니다.
"""

from __future__ import annotations

from typing import Any


class SpatialBudgeter:
    """도구별 공간 및 분량 예산 집행기"""

    # -------------------------------------------------------------------------
    # HWP 1-Page Budgeting
    # -------------------------------------------------------------------------
    @staticmethod
    def audit_hwp_one_page(spec: dict[str, Any]) -> dict[str, Any]:
        """
        한글 1페이지 보고서 분량 예산 검사 (기준: 최대 28줄)
        """
        max_lines = spec.get("typography_budget", {}).get("max_lines", 28)
        current_lines = 0
        line_breakdown = []

        # 1. 제목 및 상단 여백 (약 3줄 차지)
        current_lines += 3
        line_breakdown.append(("제목부", 3))

        # 2. 요약 상자 (줄 수 + 테두리 마진 2줄)
        sum_box = spec.get("executive_summary_box", [])
        box_lines = len(sum_box) + 2
        current_lines += box_lines
        line_breakdown.append(("요약상자", box_lines))

        # 3. 본문 섹션
        sections = spec.get("sections", [])
        for sec in sections:
            # 섹션 제목 (□ 1줄 + 여백 1줄 = 2줄)
            current_lines += 2
            items = sec.get("items", [])
            for item in items:
                text = str(item.get("text", ""))
                # 한글 기준 1줄당 약 35자 내외 (장평 100%, 15pt 들여쓰기 감안)
                item_lines = max(1, (len(text) // 35) + (1 if len(text) % 35 > 0 else 0))
                current_lines += item_lines
            line_breakdown.append((f"섹션: {sec.get('title', '무제')}", 2 + len(items)))

        # 4. 하단 표 (헤더 1줄 + 데이터행 수 * 1.5줄)
        action_table = spec.get("action_table")
        if action_table:
            rows = len(action_table.get("rows", []))
            table_lines = int(1 + (rows * 1.5))
            current_lines += table_lines
            line_breakdown.append(("실행계획 표", table_lines))

        is_overflow = current_lines > max_lines
        overflow_amount = max(0, current_lines - max_lines)

        recommendation = ""
        if is_overflow:
            recommendation = (
                f"총 {current_lines}줄로 A4 1페이지 기준({max_lines}줄)을 {overflow_amount}줄 초과했습니다. "
                f"본문 문장의 불필요한 수식어를 제거하거나, 자간(-5%) 및 줄간격(150%) 압축 파라미터를 적용하세요."
            )

        return {
            "passed": not is_overflow,
            "max_allowed_lines": max_lines,
            "estimated_lines": current_lines,
            "overflow_lines": overflow_amount,
            "breakdown": line_breakdown,
            "recommendation": recommendation,
        }

    # -------------------------------------------------------------------------
    # PPT Slide Spatial Budgeting
    # -------------------------------------------------------------------------
    @staticmethod
    def audit_ppt_slide(slide_spec: dict[str, Any]) -> dict[str, Any]:
        """
        PPT 슬라이드 카드별 텍스트 글자 수 및 시각적 여백 검사
        - Action Title: 최대 65자
        - 카드당 본문: 최대 2줄(45자 이내)
        """
        warnings = []
        action_title = slide_spec.get("action_title", "")
        if len(action_title) > 65:
            warnings.append(
                f"Action Title이 너무 깁니다 ({len(action_title)}자 > 권장 65자). "
                f"핵심 결론만 두괄식으로 압축하세요."
            )

        layout_cfg = slide_spec.get("layout_config", {})
        max_chars = layout_cfg.get("max_body_chars", 50)

        items = slide_spec.get("items", [])
        for idx, it in enumerate(items, start=1):
            body = str(it.get("desc") or it.get("body") or it.get("subtext") or "")
            if len(body) > max_chars:
                warnings.append(
                    f"카드 #{idx} 본문이 {len(body)}자로 상한({max_chars}자)을 초과했습니다. "
                    f"글자가 넘쳐 카드가 답답해집니다. 요약이 필요합니다."
                )

        return {
            "passed": len(warnings) == 0,
            "warnings": warnings,
            "card_count": len(items),
            "action_title_length": len(action_title),
        }

    # -------------------------------------------------------------------------
    # Excel Grid Topology Budgeting
    # -------------------------------------------------------------------------
    @staticmethod
    def audit_excel_topology(spec: dict[str, Any]) -> dict[str, Any]:
        """
        엑셀 토폴로지 검사:
        - 열 너비 합계가 가로 1페이지 인쇄(약 10~12열)를 초과하지 않는지
        - 상단 KPI 타일 개수가 적정한지 (최대 4개)
        """
        warnings = []
        headers = spec.get("grid", {}).get("headers", [])
        col_count = len(headers)
        if col_count > 12:
            warnings.append(
                f"데이터 열 수가 {col_count}개로 가로 1페이지 인쇄 규격(최대 12열)을 초과할 위험이 있습니다. "
                f"불필요한 열을 분리하거나 축약하세요."
            )

        kpis = spec.get("kpi_tiles", [])
        if len(kpis) == 0:
            warnings.append("상단 KPI 요약 카드가 없습니다. 임원 보고용 핵심 수치 3개를 상단에 배치하세요.")
        elif len(kpis) > 4:
            warnings.append("상단 KPI 카드가 4개를 초과하여 화면이 복잡합니다. 핵심 3~4개로 압축하세요.")

        return {
            "passed": len(warnings) == 0,
            "warnings": warnings,
            "column_count": col_count,
            "kpi_count": len(kpis),
        }
