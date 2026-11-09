# -*- coding: utf-8 -*-
"""
mcp_server/tools/dynamic_form_tools.py
자연어 지시 기반 범용 동적 서식 조립 엔진 (Universal Dynamic Form Builder).

"7대 컴포넌트"라는 정형화된 틀을 벗어나, 사용자가 자연어로 지시하는 모든 서식
(회의록, 1페이지 기획서, 업무 점검표, 대안 평가 매트릭스, 예산 견적서 등)을
10대 레고 블록(Composable Blocks)으로 즉석에서 조립하여 Word, Excel, PPT, HWP에 렌더링합니다.

[지원 10대 레고 블록]
1. title_header: 대제목, 서브제목, 문서번호, 일자
2. approval_line: 우측 상단 N단 결재선 (담당/팀장/본부장/대표이사)
3. key_value_grid: 2열 또는 4열 메타데이터 표 (일시, 장소, 참석자, 주관부서)
4. callout_box: 핵심 요약/주의/안내 상자 (좌측 2.5pt 강조선 + 배경 음영)
5. table: 겉바속촉 스마트 표 (헤더 음영, 이중선, 오픈 테이블, 숫자 우측 정렬)
6. comparison: AS-IS vs TO-BE 대칭 비교 카드
7. checklist: 점검 문항 + 담당자 + 기한 + [✔]/[ ] 상태 배지
8. process_flow: N단계 카드 + ➜ 화살표
9. kpi_cards: 핵심 지표 대시보드 카드 블록
10. signature: 하단 선언문 + 일자 + 서명/직인란
"""

from __future__ import annotations

import logging
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr, word_undo_group
from .parsers import strip_markdown_syntax
from .component_tools import (
    COLOR_BORDER_GRAY_BGR,
    COLOR_LIGHT_BG_BGR,
    COLOR_LIGHT_BLUE_BGR,
    COLOR_NAVY_BGR,
    COLOR_STEEL_BLUE_BGR,
    COLOR_TEXT_DARK_BGR,
    COLOR_WHITE_BGR,
    WD_BORDER_BOTTOM,
    ExcelComponentRenderer,
    WordComponentRenderer,
)

logger = logging.getLogger("mcp_server.dynamic_form_tools")


def _resolve_base_form_style(base_form: str | None) -> dict[str, Any]:
    """참조 양식(base_form)의 AST에서 추출된 실제 배경색·글자색·테두리 정보를 동적 빌더용 스타일로 변환"""
    default_style = {
        "accent_bgr": COLOR_NAVY_BGR,
        "header_bg_bgr": COLOR_NAVY_BGR,
        "header_text_bgr": COLOR_WHITE_BGR,
        "has_double_border": True,
        "is_open_sides": True,
    }
    if not base_form:
        return default_style
    try:
        from .document_forms import _get_bm25_searcher
        meta = _get_bm25_searcher().knowledge.get(base_form, {})
        if not meta:
            return default_style
        bg_colors = meta.get("bg_colors", [])
        if bg_colors:
            hex_c = bg_colors[0].lstrip("#")
            r, g, b = int(hex_c[0:2], 16), int(hex_c[2:4], 16), int(hex_c[4:6], 16)
            lum = 0.299 * r + 0.587 * g + 0.114 * b
            hdr_bg = hex_to_bgr(f"#{hex_c}")
            if lum > 165:
                # 밝은 회색/연청색 헤더인 경우 (예: 금천기본표 #C0C0C0, 경기신보3단표 #F0F0F0)
                return {
                    "accent_bgr": COLOR_TEXT_DARK_BGR,
                    "header_bg_bgr": hdr_bg,
                    "header_text_bgr": COLOR_TEXT_DARK_BGR,
                    "has_double_border": meta.get("has_double_border", True),
                    "is_open_sides": meta.get("is_open_sides", True),
                }
            return {
                "accent_bgr": hdr_bg,
                "header_bg_bgr": hdr_bg,
                "header_text_bgr": COLOR_WHITE_BGR,
                "has_double_border": meta.get("has_double_border", True),
                "is_open_sides": meta.get("is_open_sides", True),
            }
        color_theme = meta.get("color", "기본")
        theme_hex = {
            "네이비": "#002060",
            "스카이블루": "#2B579A",
            "그레이": "#505457",
            "다크그레이": "#333333",
            "레드": "#C00000",
        }.get(color_theme)
        if theme_hex:
            c_bgr = hex_to_bgr(theme_hex)
            default_style["accent_bgr"] = c_bgr
            default_style["header_bg_bgr"] = c_bgr
        default_style["has_double_border"] = meta.get("has_double_border", True)
        default_style["is_open_sides"] = meta.get("is_open_sides", True)
    except Exception:
        pass
    return default_style


# =============================================================================
# 1. Word 동적 서식 조립기
# =============================================================================
class WordDynamicFormBuilder:
    """Word 문서 내에 임의의 블록들을 순서대로 조립하여 완성형 서식 렌더링"""

    @staticmethod
    def render_form(
        word_app: Any,
        doc: Any,
        title: str,
        blocks: list[dict[str, Any]],
        base_form: str | None = None,
    ) -> str:
        sel = word_app.Selection
        style = _resolve_base_form_style(base_form)
        accent_bgr = style["accent_bgr"]
        header_bg_bgr = style["header_bg_bgr"]
        header_text_bgr = style["header_text_bgr"]
        has_double_border = style["has_double_border"]

        for block in blocks:
            b_type = str(block.get("type", "")).strip().lower()

            # 0. 참조 양식 직접 삽입 블록 (Reference Form Block)
            if b_type == "reference_form":
                ref_name = block.get("form_name", base_form or "")
                if ref_name:
                    from .document_forms import _resolve_form_name, render_form as render_ref_form
                    from .document_forms import _get_bm25_searcher, _recipes
                    if not _recipes():
                        from .multi_doc_form_renderer import MultiDocFormRenderer
                        meta = _get_bm25_searcher().knowledge.get(ref_name, {})
                        if meta:
                            MultiDocFormRenderer.render_to_word(
                                form_name=ref_name,
                                meta=meta,
                                args=block.get("args"),
                                headers=block.get("headers"),
                                rows=block.get("rows"),
                                replacements=block.get("replacements"),
                            )
                            sel = word_app.Selection
                        continue
                    recipe_name, err = _resolve_form_name(ref_name)
                    if not err and recipe_name:
                        render_ref_form(
                            word_app,
                            doc,
                            recipe_name,
                            args=block.get("args"),
                            headers=block.get("headers"),
                            rows=block.get("rows"),
                            replacements=block.get("replacements"),
                        )
                        sel = word_app.Selection
                continue

            # 1. 타이틀 헤더 블록
            if b_type == "title_header":
                sub = block.get("subtitle", "")
                doc_no = block.get("doc_no", "")
                date_str = block.get("date", "")

                # 메타정보(문서번호/일자) 우측 정렬
                if doc_no or date_str:
                    meta_p = sel.Range.Paragraphs.Add()
                    meta_p.Range.Text = f"문서번호: {doc_no or '-'}   |   일자: {date_str or '-'}\n"
                    meta_p.Range.Font.Name = "맑은 고딕"
                    meta_p.Range.Font.Size = 9
                    meta_p.Range.Font.Color = hex_to_bgr("#6B7280")
                    meta_p.Range.ParagraphFormat.Alignment = 2  # 우측
                    sel.Start = meta_p.Range.End

                # 대제목
                t_p = sel.Range.Paragraphs.Add()
                t_p.Range.Text = f"{block.get('title', title)}\n"
                t_p.Range.Font.Name = "맑은 고딕"
                t_p.Range.Font.Size = 18
                t_p.Range.Font.Bold = True
                t_p.Range.Font.Color = accent_bgr
                t_p.Range.ParagraphFormat.Alignment = 1  # 중앙
                t_p.Range.ParagraphFormat.SpaceAfter = 6
                sel.Start = t_p.Range.End

                # 서브제목
                if sub:
                    s_p = sel.Range.Paragraphs.Add()
                    s_p.Range.Text = f"- {sub} -\n"
                    s_p.Range.Font.Name = "맑은 고딕"
                    s_p.Range.Font.Size = 11
                    s_p.Range.Font.Color = hex_to_bgr("#4B5563")
                    s_p.Range.ParagraphFormat.Alignment = 1
                    s_p.Range.ParagraphFormat.SpaceAfter = 12
                    sel.Start = s_p.Range.End

            # 2. 결재선 블록 (Approval Line)
            elif b_type == "approval_line":
                positions = block.get("positions", ["담당", "팀장", "부서장", "대표이사"])
                if not positions:
                    continue
                n = len(positions)
                tbl = doc.Tables.Add(sel.Range, 2, n)
                tbl.AllowAutoFit = False
                tbl.Rows.Alignment = 2  # 우측 정렬 표

                col_w = 55
                for c, pos in enumerate(positions, 1):
                    c1 = tbl.Cell(1, c)
                    c1.Range.Text = pos
                    c1.Range.Font.Name = "맑은 고딕"
                    c1.Range.Font.Size = 9
                    c1.Range.Font.Bold = True
                    c1.Range.ParagraphFormat.Alignment = 1
                    c1.Shading.BackgroundPatternColor = hex_to_bgr("#E9ECEF")
                    c1.Width = col_w

                    c2 = tbl.Cell(2, c)
                    c2.Height = 35
                    c2.Range.ParagraphFormat.Alignment = 1
                    c2.Width = col_w

                for b in (-1, -2, -3, -4, -5, -6):
                    try:
                        tbl.Borders(b).LineStyle = 1
                        tbl.Borders(b).LineWidth = 4
                        tbl.Borders(b).Color = COLOR_BORDER_GRAY_BGR
                    except Exception:
                        pass

                sel.Start = tbl.Range.End
                sel.InsertParagraphAfter()
                sel.MoveDown(Unit=5, Count=1)

            # 3. 메타데이터 그리드 (Key-Value Grid)
            elif b_type == "key_value_grid":
                items = block.get("items", [])
                if not items:
                    continue
                cols_count = block.get("cols", 4)
                if cols_count == 4:
                    num_rows = max((len(items) + 1) // 2, 1)
                    tbl = doc.Tables.Add(sel.Range, num_rows, 4)
                    for idx, it in enumerate(items):
                        r = idx // 2 + 1
                        c_base = (idx % 2) * 2 + 1
                        k_cell = tbl.Cell(r, c_base)
                        k_cell.Range.Text = strip_markdown_syntax(it.get("key", ""))
                        k_cell.Range.Font.Name = "맑은 고딕"
                        k_cell.Range.Font.Size = 9.5
                        k_cell.Range.Font.Bold = True
                        k_cell.Range.Font.Color = COLOR_TEXT_DARK_BGR
                        k_cell.Shading.BackgroundPatternColor = COLOR_LIGHT_BG_BGR
                        k_cell.Range.ParagraphFormat.Alignment = 1
                        k_cell.Width = 70

                        v_cell = tbl.Cell(r, c_base + 1)
                        v_cell.Range.Text = strip_markdown_syntax(it.get("value", ""))
                        v_cell.Range.Font.Name = "맑은 고딕"
                        v_cell.Range.Font.Size = 9.5
                        v_cell.Range.ParagraphFormat.Alignment = 0
                        v_cell.Width = 160
                else:
                    num_rows = max(len(items), 1)
                    tbl = doc.Tables.Add(sel.Range, num_rows, 2)
                    for r, it in enumerate(items, 1):
                        k_cell = tbl.Cell(r, 1)
                        k_cell.Range.Text = strip_markdown_syntax(it.get("key", ""))
                        k_cell.Range.Font.Name = "맑은 고딕"
                        k_cell.Range.Font.Size = 9.5
                        k_cell.Range.Font.Bold = True
                        k_cell.Shading.BackgroundPatternColor = COLOR_LIGHT_BG_BGR
                        k_cell.Range.ParagraphFormat.Alignment = 1
                        k_cell.Width = 100

                        v_cell = tbl.Cell(r, 2)
                        v_cell.Range.Text = strip_markdown_syntax(it.get("value", ""))
                        v_cell.Range.Font.Name = "맑은 고딕"
                        v_cell.Range.Font.Size = 9.5
                        v_cell.Range.ParagraphFormat.Alignment = 0
                        v_cell.Width = 360

                for b in (-1, -2, -3, -4, -5, -6):
                    try:
                        tbl.Borders(b).LineStyle = 1
                        tbl.Borders(b).LineWidth = 4
                        tbl.Borders(b).Color = COLOR_BORDER_GRAY_BGR
                    except Exception:
                        pass

                sel.Start = tbl.Range.End
                sel.InsertParagraphAfter()
                sel.MoveDown(Unit=5, Count=1)

            # 4. 콜아웃/요약 박스
            elif b_type == "callout_box":
                WordComponentRenderer.insert_callout_box(
                    word_app, doc, block.get("text", ""), block.get("box_type", "info")
                )

            # 5. 스마트 비즈니스 표 (Table)
            elif b_type == "table":
                headers = block.get("headers", [])
                rows = block.get("rows", [])
                if not headers:
                    continue

                section_title = block.get("title", "")
                if section_title:
                    sec_p = sel.Range.Paragraphs.Add()
                    sec_p.Range.Text = f"■ {section_title}\n"
                    sec_p.Range.Font.Name = "맑은 고딕"
                    sec_p.Range.Font.Size = 11.5
                    sec_p.Range.Font.Bold = True
                    sec_p.Range.Font.Color = accent_bgr
                    sec_p.Range.ParagraphFormat.SpaceBefore = 8
                    sec_p.Range.ParagraphFormat.SpaceAfter = 4
                    sel.Start = sec_p.Range.End

                WordComponentRenderer.insert_smart_table(word_app, doc, headers, rows)
                if base_form and doc.Tables.Count >= 1:
                    tbl = doc.Tables(doc.Tables.Count)
                    for c in range(1, len(headers) + 1):
                        cell = tbl.Cell(1, c)
                        cell.Range.Font.Color = header_text_bgr
                        cell.Shading.BackgroundPatternColor = header_bg_bgr
                        cell.Borders(-1).Color = accent_bgr
                        cell.Borders(WD_BORDER_BOTTOM).LineStyle = 7 if has_double_border else 1
                        cell.Borders(WD_BORDER_BOTTOM).Color = accent_bgr
                    last_r = tbl.Rows(tbl.Rows.Count)
                    for c in range(1, len(headers) + 1):
                        last_r.Cells(c).Borders(WD_BORDER_BOTTOM).Color = accent_bgr

            # 6. 대안 비교표 (Comparison)
            elif b_type == "comparison":
                WordComponentRenderer.insert_comparison(
                    word_app,
                    doc,
                    block.get("title", "비교 분석"),
                    block.get("as_is", block.get("left_items", [])),
                    block.get("to_be", block.get("right_items", [])),
                )

            # 7. 체크리스트 / 점검표 (Checklist)
            elif b_type == "checklist":
                items = block.get("items", [])
                if not items:
                    continue

                section_title = block.get("title", "점검 및 조치 사항")
                if section_title:
                    sec_p = sel.Range.Paragraphs.Add()
                    sec_p.Range.Text = f"■ {section_title}\n"
                    sec_p.Range.Font.Name = "맑은 고딕"
                    sec_p.Range.Font.Size = 11.5
                    sec_p.Range.Font.Bold = True
                    sec_p.Range.Font.Color = accent_bgr
                    sec_p.Range.ParagraphFormat.SpaceBefore = 8
                    sec_p.Range.ParagraphFormat.SpaceAfter = 4
                    sel.Start = sec_p.Range.End

                rows_count = len(items) + 1
                tbl = doc.Tables.Add(sel.Range, rows_count, 5)
                headers = ["상태", "점검 과제 및 확인 사항", "담당자", "마감 기한", "완료 여부"]
                widths = [45, 230, 65, 70, 50]

                # 헤더
                for c, h in enumerate(headers, 1):
                    cell = tbl.Cell(1, c)
                    cell.Range.Text = h
                    cell.Range.Font.Name = "맑은 고딕"
                    cell.Range.Font.Size = 9.5
                    cell.Range.Font.Bold = True
                    cell.Range.Font.Color = header_text_bgr
                    cell.Shading.BackgroundPatternColor = header_bg_bgr
                    cell.Range.ParagraphFormat.Alignment = 1
                    cell.Width = widths[c - 1]
                    cell.Borders(-1).LineStyle = 1
                    cell.Borders(-1).LineWidth = 12
                    cell.Borders(-1).Color = accent_bgr

                # 데이터 행
                for r, it in enumerate(items, 2):
                    is_done = it.get("done", False)
                    badge_icon = "✔" if is_done else "•"
                    done_text = "[완료]" if is_done else "[대기]"

                    c1 = tbl.Cell(r, 1)
                    c1.Range.Text = badge_icon
                    c1.Range.Font.Name = "Segoe UI Symbol"
                    c1.Range.Font.Size = 11
                    c1.Range.Font.Bold = True
                    c1.Range.Font.Color = hex_to_bgr("#059669") if is_done else hex_to_bgr("#9CA3AF")
                    c1.Range.ParagraphFormat.Alignment = 1
                    c1.Width = widths[0]

                    c2 = tbl.Cell(r, 2)
                    c2.Range.Text = strip_markdown_syntax(it.get("task", "-"))
                    c2.Range.Font.Name = "맑은 고딕"
                    c2.Range.Font.Size = 9.5
                    c2.Range.ParagraphFormat.Alignment = 0
                    c2.Width = widths[1]

                    c3 = tbl.Cell(r, 3)
                    c3.Range.Text = strip_markdown_syntax(it.get("assignee", "-"))
                    c3.Range.Font.Name = "맑은 고딕"
                    c3.Range.Font.Size = 9.5
                    c3.Range.ParagraphFormat.Alignment = 1
                    c3.Width = widths[2]

                    c4 = tbl.Cell(r, 4)
                    c4.Range.Text = strip_markdown_syntax(it.get("due", "-"))
                    c4.Range.Font.Name = "맑은 고딕"
                    c4.Range.Font.Size = 9.5
                    c4.Range.ParagraphFormat.Alignment = 1
                    c4.Width = widths[3]

                    c5 = tbl.Cell(r, 5)
                    c5.Range.Text = done_text
                    c5.Range.Font.Name = "맑은 고딕"
                    c5.Range.Font.Size = 9
                    c5.Range.Font.Bold = True
                    c5.Range.Font.Color = hex_to_bgr("#059669") if is_done else hex_to_bgr("#D97706")
                    c5.Range.ParagraphFormat.Alignment = 1
                    c5.Width = widths[4]

                    for c_idx in range(1, 6):
                        c_cell = tbl.Cell(r, c_idx)
                        c_cell.Borders(WD_BORDER_BOTTOM).LineStyle = 1
                        c_cell.Borders(WD_BORDER_BOTTOM).LineWidth = 4
                        c_cell.Borders(WD_BORDER_BOTTOM).Color = COLOR_BORDER_GRAY_BGR

                for c_idx in range(1, 6):
                    tbl.Cell(rows_count, c_idx).Borders(WD_BORDER_BOTTOM).LineStyle = 1
                    tbl.Cell(rows_count, c_idx).Borders(WD_BORDER_BOTTOM).LineWidth = 12
                    tbl.Cell(rows_count, c_idx).Borders(WD_BORDER_BOTTOM).Color = COLOR_NAVY_BGR

                sel.Start = tbl.Range.End
                sel.InsertParagraphAfter()
                sel.MoveDown(Unit=5, Count=1)

            # 8. 프로세스 흐름도 (Process Flow)
            elif b_type == "process_flow":
                steps = block.get("steps", [])
                if steps:
                    WordComponentRenderer.insert_process_flow(word_app, doc, steps)

            # 9. KPI 대시보드 카드 (KPI Cards)
            elif b_type == "kpi_cards":
                cards = block.get("cards", block.get("kpi_data", []))
                if cards:
                    WordComponentRenderer.insert_kpi_cards(word_app, doc, cards)

            # 10. 서명 및 확인란 (Signature)
            elif b_type == "signature":
                notice = block.get("notice", "위와 같이 이상 없이 확인 및 보고합니다.")
                sign_date = block.get("date", "")
                signers = block.get("signers", ["보고자 (서명/인)"])

                p_notice = sel.Range.Paragraphs.Add()
                p_notice.Range.Text = f"\n{notice}\n"
                p_notice.Range.Font.Name = "맑은 고딕"
                p_notice.Range.Font.Size = 10.5
                p_notice.Range.ParagraphFormat.Alignment = 1
                p_notice.Range.ParagraphFormat.SpaceBefore = 16
                p_notice.Range.ParagraphFormat.SpaceAfter = 12

                if sign_date:
                    p_date = sel.Range.Paragraphs.Add()
                    p_date.Range.Text = f"{sign_date}\n\n"
                    p_date.Range.Font.Name = "맑은 고딕"
                    p_date.Range.Font.Size = 10.5
                    p_date.Range.ParagraphFormat.Alignment = 1

                for s in signers:
                    p_s = sel.Range.Paragraphs.Add()
                    p_s.Range.Text = f"{s}\n"
                    p_s.Range.Font.Name = "맑은 고딕"
                    p_s.Range.Font.Size = 11
                    p_s.Range.Font.Bold = True
                    p_s.Range.Font.Color = COLOR_TEXT_DARK_BGR
                    p_s.Range.ParagraphFormat.Alignment = 2
                    p_s.Range.ParagraphFormat.SpaceAfter = 6

                sel.Start = doc.Content.End
                sel.End = doc.Content.End

        return f"Word 문서에 {len(blocks)}개 블록으로 구성된 맞춤형 서식('{title}')이 성공적으로 생성되었습니다."


# =============================================================================
# 2. Excel 동적 서식 조립기
# =============================================================================
class ExcelDynamicFormBuilder:
    """Excel 시트 내에 임의의 블록들을 순차적 행으로 조립하여 완성형 서식 렌더링"""

    @staticmethod
    def render_form(excel_app: Any, sheet: Any, title: str, blocks: list[dict[str, Any]]) -> str:
        curr_row = 1
        max_col = 5

        for block in blocks:
            b_type = str(block.get("type", "")).strip().lower()

            # 1. 타이틀 헤더
            if b_type == "title_header":
                sub = block.get("subtitle", "")
                t_val = block.get("title", title)

                # A1:E1 병합 타이틀
                title_rng = sheet.Range(sheet.Cells(curr_row, 1), sheet.Cells(curr_row, max_col))
                title_rng.Merge()
                title_rng.Value = t_val
                title_rng.Font.Name = "맑은 고딕"
                title_rng.Font.Size = 15
                title_rng.Font.Bold = True
                title_rng.Font.Color = COLOR_WHITE_BGR
                title_rng.Interior.Color = COLOR_NAVY_BGR
                title_rng.HorizontalAlignment = -4108  # xlCenter
                title_rng.VerticalAlignment = -4108
                title_rng.RowHeight = 32
                curr_row += 1

                if sub:
                    sub_rng = sheet.Range(sheet.Cells(curr_row, 1), sheet.Cells(curr_row, max_col))
                    sub_rng.Merge()
                    sub_rng.Value = f"- {sub} -"
                    sub_rng.Font.Name = "맑은 고딕"
                    sub_rng.Font.Size = 10
                    sub_rng.Font.Color = hex_to_bgr("#4B5563")
                    sub_rng.HorizontalAlignment = -4108
                    sub_rng.RowHeight = 18
                    curr_row += 1

                curr_row += 1  # 공백 행

            # 2. 메타데이터 그리드 (key_value_grid)
            elif b_type == "key_value_grid":
                items = block.get("items", [])
                if not items:
                    continue
                for it in items:
                    sheet.Cells(curr_row, 1).Value = it.get("key", "")
                    sheet.Cells(curr_row, 1).Font.Name = "맑은 고딕"
                    sheet.Cells(curr_row, 1).Font.Size = 9.5
                    sheet.Cells(curr_row, 1).Font.Bold = True
                    sheet.Cells(curr_row, 1).Interior.Color = COLOR_LIGHT_BG_BGR
                    sheet.Cells(curr_row, 1).HorizontalAlignment = -4108

                    val_rng = sheet.Range(sheet.Cells(curr_row, 2), sheet.Cells(curr_row, max_col))
                    val_rng.Merge()
                    val_rng.Value = it.get("value", "")
                    val_rng.Font.Name = "맑은 고딕"
                    val_rng.Font.Size = 9.5
                    val_rng.HorizontalAlignment = -4131  # xlLeft
                    sheet.Rows(curr_row).RowHeight = 20
                    curr_row += 1
                curr_row += 1

            # 3. 테이블 블록 (table)
            elif b_type == "table":
                headers = block.get("headers", [])
                rows = block.get("rows", [])
                if not headers:
                    continue

                sec_title = block.get("title", "")
                if sec_title:
                    sec_cell = sheet.Cells(curr_row, 1)
                    sec_cell.Value = f"■ {sec_title}"
                    sec_cell.Font.Name = "맑은 고딕"
                    sec_cell.Font.Size = 11
                    sec_cell.Font.Bold = True
                    sec_cell.Font.Color = COLOR_NAVY_BGR
                    curr_row += 1

                start_table_row = curr_row
                num_cols = len(headers)
                # 헤더 입력
                for c_idx, h in enumerate(headers, 1):
                    c = sheet.Cells(curr_row, c_idx)
                    c.Value = str(h)
                    c.Font.Name = "맑은 고딕"
                    c.Font.Size = 10
                    c.Font.Bold = True
                    c.Font.Color = COLOR_WHITE_BGR
                    c.Interior.Color = COLOR_NAVY_BGR
                    c.HorizontalAlignment = -4108
                sheet.Rows(curr_row).RowHeight = 22
                curr_row += 1

                # 데이터 행 입력
                for r_data in rows:
                    for c_idx, val in enumerate(r_data, 1):
                        if c_idx > num_cols:
                            break
                        c = sheet.Cells(curr_row, c_idx)
                        s_val = str(val)
                        c.Value = s_val
                        c.Font.Name = "맑은 고딕"
                        c.Font.Size = 9.5
                        clean_num = s_val.replace(",", "").replace("%", "").strip()
                        if clean_num.isdigit():
                            c.NumberFormatLocal = "#,##0"
                            c.HorizontalAlignment = -4152  # xlRight
                        else:
                            c.HorizontalAlignment = -4131  # xlLeft
                    sheet.Rows(curr_row).RowHeight = 19
                    curr_row += 1

                # 겉바속촉 테두리 적용
                tbl_range = sheet.Range(sheet.Cells(start_table_row, 1), sheet.Cells(curr_row - 1, num_cols))
                ExcelComponentRenderer.auto_format_table(excel_app, sheet, tbl_range.Address)
                curr_row += 1

            # 4. 콜아웃/요약 박스 (callout_box)
            elif b_type == "callout_box":
                txt = block.get("text", "")
                box_rng = sheet.Range(sheet.Cells(curr_row, 1), sheet.Cells(curr_row, max_col))
                box_rng.Merge()
                box_rng.Value = f"📌 {txt}"
                box_rng.Font.Name = "맑은 고딕"
                box_rng.Font.Size = 9.5
                box_rng.Interior.Color = COLOR_LIGHT_BLUE_BGR
                box_rng.Borders(7).LineStyle = 1  # 좌측 실선
                box_rng.Borders(7).Weight = -4138
                box_rng.Borders(7).Color = COLOR_STEEL_BLUE_BGR
                box_rng.RowHeight = 28
                curr_row += 2

        sheet.Columns.AutoFit()
        return f"Excel '{sheet.Name}' 시트에 {len(blocks)}개 블록으로 구성된 맞춤형 서식('{title}')이 성공적으로 생성되었습니다."


# =============================================================================
# MCPServer 도구 등록
# =============================================================================
def register_dynamic_form_tools(mcp: MCPServer) -> None:
    """자연어 지시 기반 범용 동적 서식 생성 도구를 MCPServer 서버에 등록"""

    @mcp.tool()
    def office_build_custom_form(
        title: str,
        blocks: list[dict[str, Any]],
        base_form: str | None = None,
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [2순위 폴백: 참조 스타일 상속 다중 블록 서식 빌더]
        단일 참조 양식(office_insert_form)만으로 구성하기 어려운 다중 블록 복합 문서(회의록, 기획서, 점검표 등)나
        참조 양식과 열 개수가 다른 표를 조립할 때 사용합니다.
        base_form에 office_list_forms로 찾은 참조 양식 이름을 지정하면, 해당 참조 양식의 실제 AST 배경색·글자색·이중선 스타일을 상속받아 렌더링합니다.

        [지원 블록 타입 (blocks 리스트 요소)]:
        0. {'type': 'reference_form', 'form_name': '...', 'args': [...], 'headers': [...], 'rows': [[...]], 'replacements': {...}} (표준 참조 양식 블록 직접 삽입 및 커스텀)
        1. {'type': 'title_header', 'title': '...', 'subtitle': '...', 'doc_no': '...', 'date': '...'}
        2. {'type': 'approval_line', 'positions': ['담당', '팀장', '본부장', '대표이사']}
        3. {'type': 'key_value_grid', 'cols': 4, 'items': [{'key': '일시', 'value': '...'}, ...]}
        4. {'type': 'callout_box', 'text': '...', 'box_type': 'info'|'tip'|'warning'|'quote'}
        5. {'type': 'table', 'title': '...', 'headers': ['항목', '금액'], 'rows': [['인프라', '10,000,000']]}
        6. {'type': 'comparison', 'title': '...', 'as_is': ['...'], 'to_be': ['...']}
        7. {'type': 'checklist', 'title': '...', 'items': [{'task': '...', 'assignee': '...', 'due': '...', 'done': True|False}]}
        8. {'type': 'process_flow', 'steps': [{'title': '...', 'desc': '...'}]}
        9. {'type': 'kpi_cards', 'cards': [{'label': '...', 'value': '...', 'change': '...'}]}
        10. {'type': 'signature', 'notice': '...', 'date': '...', 'signers': ['...']}
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, f"Build Custom Form: {title}"):
                return WordDynamicFormBuilder.render_form(word_app, doc, title, blocks, base_form=base_form)
        elif target_app == "excel":
            excel_app, wb, sheet = com_manager.get_excel()
            if not sheet:
                return "열려 있는 Excel 시트를 찾을 수 없습니다."
            return ExcelDynamicFormBuilder.render_form(excel_app, sheet, title, blocks)

        return f"현재 {target_app}용 동적 서식 빌더는 순차 확장 중입니다. Word 또는 Excel에서 먼저 확인해 보세요."
