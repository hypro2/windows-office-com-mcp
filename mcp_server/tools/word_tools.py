# -*- coding: utf-8 -*-
"""
mcp_server/tools/word_tools.py
MS Word 실시간 문서 읽기(전체/마우스 선택영역/표), 대필 작성, 서식 적용 및 비즈니스 DOCX 양식 도구.
"""

from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import Any, Literal

try:
    import winreg
except ImportError:  # pragma: no cover
    winreg = None

_INSTALLED_FONT_KEYS: set[str] | None = None


def _get_installed_font_keys() -> set[str]:
    """Windows 레지스트리에서 설치된 글꼴 표시명 목록(소문자) 1회 캐시 조회."""
    global _INSTALLED_FONT_KEYS
    if _INSTALLED_FONT_KEYS is not None:
        return _INSTALLED_FONT_KEYS
    keys: set[str] = set()
    if winreg is not None:
        subkey = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, subkey) as hk:
                    idx = 0
                    while True:
                        try:
                            val_name, _val_data, _ = winreg.EnumValue(hk, idx)
                            clean_name = re.sub(r"\s*\(.*?\)\s*$", "", str(val_name)).strip().lower()
                            if clean_name:
                                keys.add(clean_name)
                            idx += 1
                        except OSError:
                            break
            except OSError:
                continue
    _INSTALLED_FONT_KEYS = keys
    return keys


def _resolve_word_font(font_name: str | None, fallback: str = "맑은 고딕") -> str:
    """Word COM용 안전한 글꼴명 반환(미설치 전용 서체 시 ClearType '맑은 고딕'으로 자동 폴백)."""
    if not font_name:
        return fallback
    norm = font_name.strip().lower()
    if norm in ("맑은 고딕", "malgun gothic", "times new roman", "consolas", "calibri", "arial", "batang", "바탕"):
        return font_name
    installed = _get_installed_font_keys()
    if not installed:
        return font_name
    base_token = re.sub(r"\s+(bold|regular|ttf|light|medium)\b", "", norm).strip()
    if any(norm in k or (base_token and base_token in k) for k in installed):
        return font_name
    return fallback

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr, word_undo_group
from .parsers import parse_document_elements, parse_inline_spans, strip_markdown_syntax

# Word 정렬 상수 (WdParagraphAlignment)
WD_ALIGN_MAP = {
    "left": 0,
    "center": 1,
    "right": 2,
    "justify": 3,
}

# 선택 영역의 서식이 섞여 있을 때 Font.Size, Font.Bold 등이 돌려주는 값 (wdUndefined)
WD_UNDEFINED = 9999999

# Word 서식 프리셋 정의
WORD_STYLE_PRESETS: dict[str, dict[str, Any]] = {
    "navy_title": {
        "font_name": "맑은 고딕",
        "font_size": 20,
        "bold": True,
        "color_hex": "#1F3864",
        "align": "center",
        "space_after": 14,
    },
    "heading1": {
        "font_name": "맑은 고딕",
        "font_size": 15,
        "bold": True,
        "color_hex": "#002060",
        "align": "left",
        "space_before": 10,
        "space_after": 6,
    },
    "heading2": {
        "font_name": "맑은 고딕",
        "font_size": 13,
        "bold": True,
        "color_hex": "#1F3864",
        "align": "left",
        "space_before": 6,
        "space_after": 4,
    },
    "body": {
        "font_name": "맑은 고딕",
        "font_size": 10.5,
        "bold": False,
        "color_hex": "#333333",
        "align": "justify",
        "space_after": 4,
    },
}


def _table_to_markdown(word_table: Any) -> str:
    """Word COM Table 객체를 Markdown 표 문자열로 변환"""
    rows_count = word_table.Rows.Count
    cols_count = word_table.Columns.Count
    md_lines: list[str] = []
    for r in range(1, rows_count + 1):
        row_vals: list[str] = []
        for c in range(1, cols_count + 1):
            try:
                raw = word_table.Cell(r, c).Range.Text
                # Word 셀 텍스트 끝에는 \r\x07 이 붙음
                cleaned = raw.replace("\r\x07", "").replace("\r", " ").replace("\n", " ").strip()
            except Exception:
                cleaned = ""
            row_vals.append(cleaned)
        md_lines.append("| " + " | ".join(row_vals) + " |")
        if r == 1:
            md_lines.append("| " + " | ".join(["---"] * cols_count) + " |")
    return "\n".join(md_lines)


def _apply_font_and_para_to_range(
    rng: Any,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color_hex: str | None = None,
    align: str | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing_pt: float | None = None,
    left_indent_pt: float | None = None,
    right_indent_pt: float | None = None,
) -> None:
    if font_name:
        resolved_font = _resolve_word_font(font_name)
        rng.Font.Name = resolved_font
        try:
            rng.Font.NameFarEast = resolved_font
        except Exception:
            pass
    if font_size is not None:
        rng.Font.Size = float(font_size)
    if bold is not None:
        rng.Font.Bold = -1 if bold else 0
    if italic is not None:
        rng.Font.Italic = -1 if italic else 0
    if underline is not None:
        rng.Font.Underline = 1 if underline else 0
    if color_hex:
        rng.Font.Color = hex_to_bgr(color_hex)
    if align and align.lower() in WD_ALIGN_MAP:
        rng.ParagraphFormat.Alignment = WD_ALIGN_MAP[align.lower()]
    if space_before is not None:
        rng.ParagraphFormat.SpaceBefore = float(space_before)
    if space_after is not None:
        rng.ParagraphFormat.SpaceAfter = float(space_after)
    if line_spacing_pt is not None:
        rng.ParagraphFormat.LineSpacingRule = 4  # wdLineSpaceExactly, 없으면 18이 1.5줄로 해석됨
        rng.ParagraphFormat.LineSpacing = float(line_spacing_pt)
    if left_indent_pt is not None:
        rng.ParagraphFormat.LeftIndent = float(left_indent_pt)
    if right_indent_pt is not None:
        rng.ParagraphFormat.RightIndent = float(right_indent_pt)


def _set_paragraph_border(rng: Any, border_id: int, color_hex: str, line_width: int) -> None:
    """range가 속한 문단의 한쪽 테두리를 실선으로 설정 (-2: 왼쪽, -3: 아래)"""
    border = rng.ParagraphFormat.Borders(border_id)
    border.LineStyle = 1
    border.LineWidth = line_width
    border.Color = hex_to_bgr(color_hex)


def _type_text_with_inline_markdown(
    sel: Any,
    text: str,
    default_bold: bool | None = None,
    default_italic: bool | None = None,
    default_underline: bool | None = None,
) -> None:
    """
    Word Selection에 텍스트를 작성할 때:
    1) 줄 시작의 '# ', '## ' 헤더 기호는 제거하고 해당 줄 전체를 볼드로 처리
    2) parse_inline_spans를 통해 **볼드**, *이탤릭*, `코드`를 실제 Word Font 서식으로 적용
    3) 모든 원시 마크다운 기호가 워드 문서에 단 1자도 노출되지 않도록 완전 보장
    """
    if not text:
        return

    init_bold = bool(default_bold) if default_bold is not None else False
    init_italic = bool(default_italic) if default_italic is not None else False

    lines = str(text).splitlines()
    for l_idx, line in enumerate(lines):
        line_clean = line.rstrip("\r")
        line_bold = init_bold

        # 헤더 마크다운 (# 제목, ## 소제목) 감지 및 기호 제거
        m_head = re.match(r"^(#{1,6})\s*(.*)", line_clean)
        if m_head:
            line_bold = True
            line_clean = m_head.group(2)

        # 리스트 마크다운 (- 항목, * 항목) 감지 -> 깔끔한 글머리 기호(• )로 변환
        m_bullet = re.match(r"^[\*\-]\s+(.*)", line_clean)
        if m_bullet:
            line_clean = "• " + m_bullet.group(1)

        # parse_inline_spans 활용한 인라인 서식 파싱 및 타이핑
        spans = parse_inline_spans(line_clean)
        for span in spans:
            s_text = span.get("text", "")
            if not s_text:
                continue
            is_b = span.get("bold", False) or line_bold
            is_i = span.get("italic", False) or init_italic
            is_code = span.get("code", False)

            sel.Font.Bold = is_b
            sel.Font.Italic = is_i
            if is_code:
                orig_color = sel.Font.Color
                sel.Font.Name = "Consolas"
                sel.TypeText(s_text)
                sel.Font.Name = "맑은 고딕"
                sel.Font.Color = orig_color
            else:
                sel.TypeText(s_text)

            # 기본 서식으로 복귀
            sel.Font.Bold = line_bold
            sel.Font.Italic = init_italic

        if l_idx < len(lines) - 1:
            sel.TypeParagraph()


def register_word_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def word_read_document(
        scope: Literal["all", "selection", "tables"] = "all",
        max_chars: int = 12000,
    ) -> dict[str, Any]:
        """
        현재 열려 있는 MS Word 활성 문서의 내용을 읽어옵니다.
        - scope="all": 문서 전체 본문 텍스트와 표(Markdown 형식)를 함께 읽어옵니다.
        - scope="selection": 사용자가 현재 마우스로 드래그/선택한 영역(Selection)의 텍스트와 서식 정보를 읽어옵니다. 서식이 섞여 있으면 font_size, bold는 None입니다.
        - scope="tables": 문서 내 모든 표를 Markdown 표 형식으로 추출합니다.
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "현재 열려 있는 MS Word 문서가 없습니다."}

        try:
            if scope == "selection":
                sel = word_app.Selection
                text = (sel.Text or "").replace("\r\x07", "\n").replace("\r", "\n").strip()
                return {
                    "success": True,
                    "document_name": doc.Name,
                    "scope": "selection",
                    "text": text[:max_chars],
                    "font_name": sel.Font.Name,
                    "font_size": None if sel.Font.Size == WD_UNDEFINED else sel.Font.Size,
                    "bold": None if sel.Font.Bold == WD_UNDEFINED else bool(sel.Font.Bold),
                }

            tables_md: list[str] = []
            for t_idx in range(1, doc.Tables.Count + 1):
                try:
                    tables_md.append(f"### [표 {t_idx}]\n" + _table_to_markdown(doc.Tables(t_idx)))
                except Exception:
                    continue

            if scope == "tables":
                return {
                    "success": True,
                    "document_name": doc.Name,
                    "tables_count": doc.Tables.Count,
                    "tables_markdown": "\n\n".join(tables_md)[:max_chars],
                }

            full_text = (doc.Content.Text or "").replace("\r\x07", " | ").replace("\r", "\n").strip()
            return {
                "success": True,
                "document_name": doc.Name,
                "paragraphs_count": doc.Paragraphs.Count,
                "tables_count": doc.Tables.Count,
                "text": full_text[:max_chars],
                "truncated": len(full_text) > max_chars,
                "tables_markdown": "\n\n".join(tables_md)[: max_chars // 2],
            }
        except Exception as e:
            return {"success": False, "error": f"Word 문서 읽기 오류: {e}"}

    @mcp.tool()
    def word_write_text(
        text: str,
        mode: Literal["insert", "replace_selection", "append"] = "insert",
        style_preset: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color_hex: str | None = None,
        align: Literal["left", "center", "right", "justify"] | None = None,
        new_paragraph: bool = True,
    ) -> dict[str, Any]:
        """
        MS Word 문서에 텍스트를 작성하고 동시에 원하는 서식을 적용합니다.

        Args:
            text: 작성할 텍스트 내용
            mode: "insert"(현재 커서 위치에 삽입), "replace_selection"(선택한 블록을 대체), "append"(문서 맨 끝에 추가)
            style_preset: 서식 프리셋 ("navy_title", "heading1", "heading2", "body")
            font_name: 글꼴 이름 (예: "맑은 고딕", "휴먼명조", "함초롬바탕")
            font_size: 글자 크기 (pt)
            bold: 굵게 여부
            color_hex: 글자 색상 (예: "#1F3864", "#2E608A", "navy")
            align: 문단 정렬 ("left", "center", "right", "justify")
            new_paragraph: 작성 후 줄바꿈(새 문단) 추가 여부
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=True, visible=True)
        if word_app is None or doc is None:
            return {"success": False, "error": "Word 연결 실패"}

        preset = WORD_STYLE_PRESETS.get(style_preset or "", {})
        f_name = font_name or preset.get("font_name")
        f_size = font_size if font_size is not None else preset.get("font_size")
        f_bold = bold if bold is not None else preset.get("bold")
        f_color = color_hex or preset.get("color_hex")
        f_align = align or preset.get("align")
        s_before = preset.get("space_before")
        s_after = preset.get("space_after")

        try:
            with word_undo_group(word_app, "MCP 텍스트 작성"):
                if mode == "append":
                    rng = doc.Content
                    rng.Collapse(0)  # wdCollapseEnd
                    rng.Select()
                sel = word_app.Selection
                if mode == "insert" and sel.Start != sel.End:
                    sel.Collapse(1)  # wdCollapseStart

                _apply_font_and_para_to_range(
                    sel,
                    font_name=f_name,
                    font_size=f_size,
                    bold=f_bold,
                    italic=italic,
                    underline=underline,
                    color_hex=f_color,
                    align=f_align,
                    space_before=s_before,
                    space_after=s_after,
                )
                _type_text_with_inline_markdown(
                    sel,
                    text,
                    default_bold=f_bold,
                    default_italic=italic,
                    default_underline=underline,
                )
                if new_paragraph:
                    sel.TypeParagraph()
            com_manager.activate_app("word")
            return {"success": True, "document_name": doc.Name, "written_length": len(text)}
        except Exception as e:
            return {"success": False, "error": f"Word 텍스트 작성 오류: {e}"}

    @mcp.tool()
    def word_format_selection(
        style_preset: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        font_size_delta: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color_hex: str | None = None,
        bg_highlight_hex: str | None = None,
        align: Literal["left", "center", "right", "justify"] | None = None,
        line_spacing_pt: float | None = None,
        paragraph_pattern: str | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        left_indent_pt: float | None = None,
        right_indent_pt: float | None = None,
        left_bar_hex: str | None = None,
        bottom_line_hex: str | None = None,
        activate_window: bool = True,
    ) -> dict[str, Any]:
        """
        사용자가 현재 MS Word에서 마우스로 선택한 영역(또는 현재 커서가 위치한 문단)에 서식을 적용하고 즉시 화면을 활성화합니다.
        paragraph_pattern을 주면 선택 영역 대신 정규식에 맞는 문단 전체에 적용합니다.
        문서를 다듬을 때는 요소 1종마다 1번씩 호출하세요. 호출 1회가 실행취소 1회입니다.

        Args:
            paragraph_pattern: 대상 문단 정규식, 표 안 문단은 제외 (예: r"^[1-7]\\.\\s"는 섹션 제목, r"^※"는 콜아웃). 일치하는 문단이 없으면 실패를 반환
            style_preset: "navy_title", "heading1", "heading2", "body" 등
            font_name: 글꼴명 (예: "맑은 고딕", "휴먼명조", "함초롬바탕")
            font_size: 절대 글자 크기 (pt)
            font_size_delta: 상대 글자 크기 조정 (예: +2.0은 2pt 크게, -1.5는 1.5pt 작게). 크기가 섞인 대상은 실패를 반환
            bold: 굵게 여부 (True / False)
            italic: 기울임 여부
            underline: 밑줄 여부
            color_hex: 글자 색상 (예: "#1F3864", "#2E608A", "navy")
            bg_highlight_hex: 배경 음영 색상 (예: "#EEF2F7", "#FFF2CC")
            align: 문단 정렬 ("left", "center", "right", "justify")
            line_spacing_pt: 고정 줄간격 (pt). 글자 크기보다 작으면 글자가 잘림
            space_before: 문단 앞 간격 (pt, 프리셋보다 우선)
            space_after: 문단 뒤 간격 (pt, 프리셋보다 우선)
            left_indent_pt: 왼쪽 들여쓰기 (pt)
            right_indent_pt: 오른쪽 들여쓰기 (pt)
            left_bar_hex: 문단 왼쪽 2.25pt 강조 바 색상 (콜아웃용)
            bottom_line_hex: 문단 아래 0.75pt 구분선 색상 (섹션 제목용)
            activate_window: 적용 후 Word 창을 화면 최상단으로 활성화할지 여부 (기본 True)
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "열려 있는 Word 문서가 없습니다. 먼저 Word를 열어주세요."}

        preset = WORD_STYLE_PRESETS.get(style_preset or "", {})
        try:
            if paragraph_pattern:
                regex = re.compile(paragraph_pattern)
                targets = [
                    p.Range
                    for p in doc.Paragraphs
                    if regex.search((p.Range.Text or "").strip()) and not p.Range.Information(12)  # wdWithInTable
                ]
                if not targets:
                    return {"success": False, "error": f"패턴에 맞는 본문 문단이 없습니다: {paragraph_pattern}"}
            else:
                sel = word_app.Selection
                targets = [sel.Range if sel.Start != sel.End else sel.Paragraphs(1).Range]

            calc_font_size = font_size if font_size is not None else preset.get("font_size")
            use_delta = calc_font_size is None and font_size_delta is not None
            if use_delta and any(rng.Font.Size == WD_UNDEFINED for rng in targets):
                return {
                    "success": False,
                    "error": "글자 크기가 섞인 대상에는 font_size_delta를 적용할 수 없습니다. font_size로 지정하세요.",
                }

            with word_undo_group(word_app, "MCP 선택영역 서식 적용"):
                for target_rng in targets:
                    # 상대 델타는 문단마다 현재 크기 기준으로 계산
                    if use_delta:
                        calc_font_size = max(5.0, float(target_rng.Font.Size) + font_size_delta)

                    _apply_font_and_para_to_range(
                        target_rng,
                        font_name=font_name or preset.get("font_name"),
                        font_size=calc_font_size,
                        bold=bold if bold is not None else preset.get("bold"),
                        italic=italic,
                        underline=underline,
                        color_hex=color_hex or preset.get("color_hex"),
                        align=align or preset.get("align"),
                        space_before=space_before if space_before is not None else preset.get("space_before"),
                        space_after=space_after if space_after is not None else preset.get("space_after"),
                        line_spacing_pt=line_spacing_pt,
                        left_indent_pt=left_indent_pt,
                        right_indent_pt=right_indent_pt,
                    )
                    if bg_highlight_hex:
                        target_rng.Shading.BackgroundPatternColor = hex_to_bgr(bg_highlight_hex)
                    if left_bar_hex:
                        _set_paragraph_border(target_rng, -2, left_bar_hex, 18)  # wdLineWidth225pt
                    if bottom_line_hex:
                        _set_paragraph_border(target_rng, -3, bottom_line_hex, 6)  # wdLineWidth075pt

            if activate_window:
                com_manager.activate_app("word")

            return {
                "success": True,
                "document_name": doc.Name,
                "matched_paragraphs": len(targets),
                "formatted_text_preview": (targets[0].Text or "").strip()[:100],
                "applied_font_size": calc_font_size,
            }
        except Exception as e:
            return {"success": False, "error": f"서식 적용 오류: {e}"}

    @mcp.tool()
    def word_format_table(
        table_index: int | None = None,
        theme: Literal["navy", "ib_steel_blue", "sage_grid", "slate_contrast", "gray_form", "gray"] = "navy",
        auto_fit: bool = True,
        header_bold: bool = True,
        zebra_shading: bool = True,
        activate_window: bool = True,
    ) -> dict[str, Any]:
        """
        MS Word에서 현재 커서가 위치한 표(또는 table_index로 지정한 표)를 전문적인 '겉바속촉'(헤더 음영, 상하 실선, 내부 점선) 서식으로 다듬습니다.
        병합 셀이 있는 표는 첫 행을 머리글로 볼 수 없어 변경 없이 실패를 반환합니다.

        Args:
            table_index: 다듬을 표 번호 (1부터 시작). 생략 시 현재 커서가 위치한 표를 찾고, 커서가 표 밖이면 문서에 표가 1개일 때만 그 표를 다듬습니다.
            theme: 표 컬러 테마 ("navy"=#1F3864 네이비, "ib_steel_blue"=#2E608A 스틸블루, "sage_grid"=#C5DFB3 실무격자, "slate_contrast"=#E6EEF7 대비표, "gray_form"=#F2F2F2 제출양식, "gray"=#404040)
            auto_fit: 창 크기/내용에 맞게 열 너비 자동 맞춤 여부
            header_bold: 첫 행(헤더)을 굵은 글씨 및 가운데 정렬할지 여부
            zebra_shading: 데이터 행에 짝수/홀수 번갈아 연한 배경색(얼룩무늬) 적용 여부
            activate_window: 작업 후 Word 창을 화면 최상단으로 활성화할지 여부
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "열려 있는 Word 문서가 없습니다."}

        tbl = None
        try:
            if table_index is not None:
                if not 1 <= table_index <= doc.Tables.Count:
                    return {"success": False, "error": f"표 번호 {table_index}가 없습니다. 문서의 표는 {doc.Tables.Count}개입니다."}
                tbl = doc.Tables(table_index)
            else:
                sel = word_app.Selection
                if sel.Tables.Count > 0:
                    tbl = sel.Tables(1)
                elif doc.Tables.Count == 1:
                    tbl = doc.Tables(1)

            if tbl is None:
                return {"success": False, "error": "다듬을 표(Table)를 찾을 수 없습니다. 표 내부에 커서를 두거나 table_index를 지정하세요."}
            # Table.Uniform은 병합 없는 표에서도 False라 셀 수로 판별
            if tbl.Range.Cells.Count != tbl.Rows.Count * tbl.Columns.Count:
                return {"success": False, "error": "병합 셀이 있는 표는 다듬지 않았습니다. word_format_selection으로 셀 단위로 처리하세요."}

            theme_colors = {
                "navy": "1F3864",
                "ib_steel_blue": "2E608A",
                "sage_grid": "C5DFB3",
                "slate_contrast": "E6EEF7",
                "gray_form": "F2F2F2",
                "gray": "404040",
            }
            hdr_hex = theme_colors.get(theme, "1F3864")
            hdr_bg = hex_to_bgr(f"#{hdr_hex}")
            zebra_bg = hex_to_bgr("#F4F6F9" if theme in ("navy", "ib_steel_blue") else "#F9F9F9")
            border_dark = hex_to_bgr("#595959")
            border_light = hex_to_bgr("#D9D9D9")

            with word_undo_group(word_app, "MCP 표 서식 단장"):
                # 1. 테두리 초기화 및 '겉바속촉' 테두리 적용
                # -1: wdBorderTop, -2: wdBorderLeft, -3: wdBorderBottom, -4: wdBorderRight, -5: wdBorderHorizontal, -6: wdBorderVertical
                tbl.Borders.Enable = False

                # 상단 테두리: 굵은 실선 (1.5pt = 12)
                tbl.Borders(-1).LineStyle = 1
                tbl.Borders(-1).LineWidth = 12
                tbl.Borders(-1).Color = border_dark

                # 하단 테두리: 굵은 실선 (1.5pt = 12)
                tbl.Borders(-3).LineStyle = 1
                tbl.Borders(-3).LineWidth = 12
                tbl.Borders(-3).Color = border_dark

                # 내부 가로 테두리: 얇은 실선 또는 점선 (0.5pt = 4), 1행 표에는 없음
                if tbl.Rows.Count > 1:
                    tbl.Borders(-5).LineStyle = 1
                    tbl.Borders(-5).LineWidth = 4
                    tbl.Borders(-5).Color = border_light

                # 좌/우 외곽선 및 내부 세로선은 제거하여 탁 트인 세련된 느낌
                tbl.Borders(-2).LineStyle = 0
                tbl.Borders(-4).LineStyle = 0
                tbl.Borders(-6).LineStyle = 0

                # 2. 첫 행(헤더) 서식
                resolved_title_font = _resolve_word_font("맑은 고딕")
                resolved_body_font = _resolve_word_font("맑은 고딕")
                if tbl.Rows.Count >= 1 and header_bold:
                    hdr_row = tbl.Rows(1)
                    hdr_row.Range.Font.Name = resolved_title_font
                    hdr_row.Range.Font.Bold = -1
                    hdr_row.Range.Font.Color = hex_to_bgr("#000000" if theme in ("sage_grid", "slate_contrast", "gray_form") else "#FFFFFF")
                    hdr_row.Range.ParagraphFormat.Alignment = 1  # Center
                    for cell_idx in range(1, hdr_row.Cells.Count + 1):
                        hdr_row.Cells(cell_idx).Shading.BackgroundPatternColor = hdr_bg
                    # 헤더 아래 구분선: 1.0pt 실선
                    tbl.Rows(1).Borders(-3).LineStyle = 1
                    tbl.Rows(1).Borders(-3).LineWidth = 8
                    tbl.Rows(1).Borders(-3).Color = border_dark

                # 3. 데이터 행 얼룩무늬 및 폰트 정돈
                for r in range(2, tbl.Rows.Count + 1):
                    row = tbl.Rows(r)
                    row.Range.Font.Name = resolved_body_font
                    row.Range.Font.Size = 10
                    row.Range.Font.Bold = 0
                    row.Range.Font.Color = hex_to_bgr("#1E293B")
                    if zebra_shading and r % 2 == 1:
                        for c in range(1, row.Cells.Count + 1):
                            row.Cells(c).Shading.BackgroundPatternColor = zebra_bg

                # 4. 셀 안쪽 여백(Padding) 넉넉하게
                tbl.TopPadding = 5.0
                tbl.BottomPadding = 5.0
                tbl.LeftPadding = 6.0
                tbl.RightPadding = 6.0

                # 5. 열 너비 자동맞춤
                if auto_fit:
                    try:
                        tbl.AutoFitBehavior(2)  # wdAutoFitWindow = 2
                    except Exception:
                        pass

            if activate_window:
                com_manager.activate_app("word")

            return {
                "success": True,
                "document_name": doc.Name,
                "rows_count": tbl.Rows.Count,
                "cols_count": tbl.Columns.Count,
                "theme": theme,
            }
        except Exception as e:
            return {"success": False, "error": f"표 서식 단장 실패: {e}"}

    @mcp.tool()
    def word_insert_table(
        data: list[list[Any]],
        has_header: bool = True,
        theme: Literal["navy", "ib_steel_blue", "sage_grid", "slate_contrast", "gray_form", "gray"] = "navy",
    ) -> dict[str, Any]:
        """
        현재 커서 위치에 2차원 배열 데이터를 표(Table)로 생성하고 브랜드 서식(헤더 배경색, 짝수행 얼룩무늬, 테두리)을 적용합니다.

        Args:
            data: 2차원 리스트 (첫 행은 헤더 권장)
            has_header: 첫 행을 헤더로 서식 지정할지 여부
            theme: 표 컬러 테마 ("navy"=#1F3864 네이비, "ib_steel_blue"=#2E608A 스틸블루, "sage_grid"=#C5DFB3 실무양식, "slate_contrast"=#E6EEF7 대비표, "gray_form"=#F2F2F2 제출양식, "gray")
        """
        ensure_com()
        if not data or not data[0]:
            return {"success": False, "error": "표 데이터(data)가 비어 있습니다."}

        word_app, doc = com_manager.get_word(create_if_missing=True, visible=True)
        if word_app is None or doc is None:
            return {"success": False, "error": "Word 연결 실패"}

        rows_cnt = len(data)
        cols_cnt = max(len(r) for r in data)
        header_colors = {
            "navy": "#1F3864",
            "ib_steel_blue": "#2E608A",
            "sage_grid": "#C5DFB3",
            "slate_contrast": "#E6EEF7",
            "gray_form": "#F2F2F2",
            "gray": "#404040",
        }
        is_sage_or_grid = theme in ("sage_grid", "slate_contrast", "gray_form")
        hdr_fg_hex = "#000000" if is_sage_or_grid else "#FFFFFF"
        hdr_bg = hex_to_bgr(header_colors.get(theme, "#1F3864"))
        zebra_bg = hex_to_bgr("#FFFFFF" if is_sage_or_grid else ("#F4F6F9" if theme in ("navy", "ib_steel_blue") else "#F2F2F2"))
        border_bgr = hex_to_bgr("#70AD47" if theme == "sage_grid" else ("#404040" if theme == "slate_contrast" else ("#7F7F7F" if theme == "gray_form" else "#CBD5E1")))

        try:
            with word_undo_group(word_app, "MCP 표 삽입"):
                sel = word_app.Selection
                if sel.Start != sel.End:
                    sel.Collapse(0)
                tbl = doc.Tables.Add(sel.Range, rows_cnt, cols_cnt)
                try:
                    tbl.TopPadding = 6
                    tbl.BottomPadding = 6
                    tbl.LeftPadding = 8
                    tbl.RightPadding = 8
                    for r_obj in tbl.Rows:
                        r_obj.AllowBreakAcrossPages = False
                    if has_header and rows_cnt > 1:
                        tbl.Rows(1).HeadingFormat = -1
                except Exception:
                    pass
                # 1. 오픈 테이블 테두리 (상하 1.5pt 실선, 내부 가로 0.5pt 실선, 좌우 세로선 없음)
                tbl.Borders.Enable = False
                for b_idx in (-1, -3):  # 상단, 하단
                    try:
                        tbl.Borders(b_idx).LineStyle = 1
                        tbl.Borders(b_idx).LineWidth = 4 if is_sage_or_grid else 12  # grid 0.5pt / open 1.5pt
                        tbl.Borders(b_idx).Color = border_bgr if is_sage_or_grid else hex_to_bgr("#1F3864" if theme == "navy" else "#333333")
                    except Exception:
                        pass
                if rows_cnt > 1:
                    try:
                        tbl.Borders(-5).LineStyle = 1  # 내부 가로선
                        tbl.Borders(-5).LineWidth = 4  # 0.5pt
                        tbl.Borders(-5).Color = border_bgr
                    except Exception:
                        pass
                if is_sage_or_grid:
                    for b_idx in (-2, -4, -6):
                        try:
                            tbl.Borders(b_idx).LineStyle = 1
                            tbl.Borders(b_idx).LineWidth = 4
                            tbl.Borders(b_idx).Color = border_bgr
                        except Exception:
                            pass
                for b_idx in (tuple() if is_sage_or_grid else (-2, -4, -6)):  # 좌/우 외곽선 및 내부 세로선 제거 (오픈테이블)
                    try:
                        tbl.Borders(b_idx).LineStyle = 0
                    except Exception:
                        pass

                # 2. 셀 데이터 채우기 및 폰트/정렬
                for r_idx, row_vals in enumerate(data, start=1):
                    for c_idx in range(1, cols_cnt + 1):
                        raw_val = str(row_vals[c_idx - 1]) if c_idx - 1 < len(row_vals) else ""
                        val = strip_markdown_syntax(raw_val)
                        cell = tbl.Cell(r_idx, c_idx)
                        cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                        c_rng = cell.Range
                        c_rng.End = c_rng.End - 1  # 셀 종료 마커 제외
                        c_rng.Text = val
                        cell_font = _resolve_word_font("맑은 고딕")
                        c_rng.Font.Name = cell_font
                        try:
                            c_rng.Font.NameFarEast = cell_font
                        except Exception:
                            pass
                        try:
                            c_rng.ParagraphFormat.FirstLineIndent = 0
                            c_rng.ParagraphFormat.LeftIndent = 0
                            c_rng.ParagraphFormat.SpaceBefore = 0
                            c_rng.ParagraphFormat.SpaceAfter = 1.5
                            c_rng.ParagraphFormat.LineSpacingRule = 5
                            try:
                                c_rng.ParagraphFormat.LineSpacing = float(word_app.LinesToPoints(1.15))
                            except Exception:
                                c_rng.ParagraphFormat.LineSpacing = 13.8
                        except Exception:
                            pass
                        c_rng.Font.Size = 10
                        if has_header and r_idx == 1:
                            c_rng.Font.Bold = -1
                            c_rng.Font.Color = hex_to_bgr(hdr_fg_hex)
                            cell.Shading.BackgroundPatternColor = hdr_bg
                            c_rng.ParagraphFormat.Alignment = 1  # 가운데 정렬
                        else:
                            c_rng.Font.Bold = 0
                            c_rng.Font.Color = hex_to_bgr("#000000" if is_sage_or_grid else "#1F2937")
                            if r_idx % 2 == 0 and not is_sage_or_grid:  # 보고서 표만 짝수행 얼룩말 음영
                                cell.Shading.BackgroundPatternColor = zebra_bg

                # 3. 열 데이터 성격별 지능형 정렬 (2행 이상일 때 본문 열 자동 정렬)
                if rows_cnt > 1:
                    for col_i in range(1, cols_cnt + 1):
                        col_vals = []
                        hdr_text = str(data[0][col_i - 1]) if col_i - 1 < len(data[0]) else ""
                        for r in range(2, rows_cnt + 1):
                            if col_i - 1 < len(data[r - 1]):
                                col_vals.append(str(data[r - 1][col_i - 1]).strip())

                        # 순번, 날짜, 수량/금액 감지
                        hdr_clean = hdr_text.upper()
                        is_idx = any(k in hdr_clean for k in ["NO", "순번", "연번", "번호", "단계"]) or all(v.isdigit() and len(v) <= 3 for v in col_vals if v)
                        has_num = any(sym in "".join(col_vals) for sym in [",", "%", "원", "$", "₩"]) or (sum(1 for v in col_vals if re.match(r'^-?\d+(?:\.\d+)?$', re.sub(r'[,%\s원\$₩]', '', v))) >= len(col_vals) * 0.7 if col_vals else False)

                        align_val = 1 if is_idx else (2 if has_num else (1 if max((len(v) for v in col_vals), default=0) <= 8 else 0))
                        for r in range(2, rows_cnt + 1):
                            try:
                                tbl.Cell(r, col_i).Range.ParagraphFormat.Alignment = align_val
                            except Exception:
                                pass

                # 4. 헤더 줄바꿈 방지 최소폭 가드 및 비율 조정
                try:
                    tbl.AllowAutoFit = True
                    tbl.AutoFitBehavior(2)  # wdAutoFitWindow
                except Exception:
                    pass

                # 표 아래로 커서 이동 후 빈 줄 추가
                after_rng = tbl.Range
                after_rng.Collapse(0)
                after_rng.Select()
                word_app.Selection.TypeParagraph()
                try:
                    ps = word_app.Selection
                    ps.Font.Name = _resolve_word_font("맑은 고딕")
                    ps.Font.Size = 10.5
                    ps.Font.Bold = 0
                    ps.Font.Color = hex_to_bgr("#1F2937")
                    ps.ParagraphFormat.FirstLineIndent = 0
                    ps.ParagraphFormat.LeftIndent = 0
                    ps.ParagraphFormat.Borders.Enable = False
                except Exception:
                    pass

            return {"success": True, "rows": rows_cnt, "cols": cols_cnt, "theme": theme}
        except Exception as e:
            return {"success": False, "error": f"Word 표 생성 오류: {e}"}

    @mcp.tool()
    def word_render_markdown_report(
        markdown_text: str,
        theme: Literal["navy", "steel_blue"] = "navy",
        apply_brand_header: bool = False,
    ) -> dict[str, Any]:
        """
        마크다운 형식의 보고서 텍스트(# 대제목, ## 섹션 제목, - 개조식 불릿, | 표 |, > 인용박스)를
        파싱하여 현재 열려 있는 Word 문서에 전문 비즈니스 보고서 서식으로 한 번에 작성합니다.
        - theme="navy": 표준 비즈니스 네이비(#1F3864 대제목/표 헤더, #2E608A 소제목, 맑은 고딕) 적용
        - theme="steel_blue": 스틸 블루(#2E608A 대제목/표 헤더, #1F3864 강조, 맑은 고딕) 적용
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=True, visible=True)
        if word_app is None or doc is None:
            return {"success": False, "error": "Word 연결 실패"}

        elements = parse_document_elements(markdown_text)
        title_color = "#1F3864" if theme == "navy" else "#2E608A"
        h1_color = "#1F3864" if theme == "navy" else "#002060"
        h2_color = "#2E608A" if theme == "navy" else "#0284C7"
        body_color = "#1F2937"
        t_font = "맑은 고딕"
        b_font = "맑은 고딕"

        rendered_count = 0

        def _fmt_last_para(
            sb: float,
            sa: float,
            line_mult: float = 1.25,
            left_indent: float = 0.0,
            first_indent: float = 0.0,
            keep_next: bool = False,
            bottom_rule: tuple[str, int] | None = None,
        ) -> None:
            try:
                cnt = doc.Paragraphs.Count
                cur_p = doc.Paragraphs(cnt)
                is_empty_tail = cnt >= 2 and not (cur_p.Range.Text or "").replace("\r", "").strip()
                p = doc.Paragraphs(cnt - 1 if is_empty_tail else cnt)
                pf = p.Range.ParagraphFormat
                pf.SpaceBefore = sb
                pf.SpaceAfter = sa
                try:
                    pf.CharacterUnitLeftIndent = 0
                    pf.CharacterUnitFirstLineIndent = 0
                except Exception:
                    pass
                pf.LeftIndent = left_indent
                pf.FirstLineIndent = first_indent
                pf.LineSpacingRule = 5
                try:
                    pf.LineSpacing = float(word_app.LinesToPoints(line_mult))
                except Exception:
                    pf.LineSpacing = 12.0 * line_mult
                pf.KeepWithNext = -1 if keep_next else 0
                if bottom_rule:
                    _set_paragraph_border(p.Range, -3, bottom_rule[0], bottom_rule[1])
                else:
                    p.Range.Borders(-3).LineStyle = 0
                if is_empty_tail:
                    cp = cur_p.Range
                    cp.Borders(-3).LineStyle = 0
                    cp.ParagraphFormat.SpaceBefore = 0
                    cp.ParagraphFormat.SpaceAfter = 0
                    cp.ParagraphFormat.LeftIndent = 0
                    cp.ParagraphFormat.FirstLineIndent = 0
                    cp.ParagraphFormat.KeepWithNext = 0
            except Exception:
                pass
        try:
            with word_undo_group(word_app, "MCP 마크다운 보고서 렌더링"):
                idx = 0
                n = len(elements)
                while idx < n:
                    elem = elements[idx]
                    etype = elem.get("type")
                    text_val = str(elem.get("text") or elem.get("content") or "").strip()

                    if etype == "TABLE_ROW":
                        table_rows: list[list[str]] = []
                        while idx < n and elements[idx].get("type") in ("TABLE_ROW", "TABLE_SEP"):
                            if elements[idx].get("type") == "TABLE_ROW":
                                cells = elements[idx].get("cells", [])
                                if cells:
                                    table_rows.append([str(c) for c in cells])
                            idx += 1
                        if table_rows:
                            word_insert_table(table_rows, has_header=True, theme="navy" if theme == "navy" else "ib_steel_blue")
                            rendered_count += 1
                        continue
                    elif etype == "TABLE_SEP":
                        idx += 1
                        continue
                    elif etype in ("TITLE_BOX", "title"):
                        word_write_text(
                            text_val, font_name=t_font, font_size=18, bold=True, color_hex=title_color, align="left"
                        )
                        try:
                            last_p = doc.Paragraphs(doc.Paragraphs.Count)
                            last_p = None
                            _fmt_last_para(10, 8, 1.2, keep_next=True, bottom_rule=("#1F3864" if theme == "navy" else "#2E608A", 12))
                        except Exception:
                            pass
                    elif etype in ("SQUARE_HEADER", "h1"):
                        prefix = ""
                        word_write_text(
                            prefix + text_val, font_name=t_font, font_size=13, bold=True, color_hex=h1_color, align="left"
                        )
                        try:
                            last_p = doc.Paragraphs(doc.Paragraphs.Count)
                            _fmt_last_para(12, 5, 1.2, keep_next=True, bottom_rule=("#CBD5E1", 4))
                            _fmt_last_para(12, 5, 1.2, keep_next=True, bottom_rule=("#CBD5E1", 4))
                        except Exception:
                            pass
                    elif etype in ("CIRCLE_ITEM", "h2", "h3"):
                        prefix = ""
                        word_write_text(
                            prefix + text_val, font_name=b_font, font_size=11.0, bold=True, color_hex=h2_color, align="left"
                        )
                        try:
                            last_p = doc.Paragraphs(doc.Paragraphs.Count)
                            _fmt_last_para(6, 3, 1.2, keep_next=True)
                            _fmt_last_para(6, 3, 1.2, keep_next=True)
                        except Exception:
                            pass
                    elif etype in ("DASH_ITEM", "bullet"):
                        prefix = "" if text_val.startswith(("•", "-", "–", "※", "☐", "☑")) else "• "
                        word_write_text(
                            prefix + text_val, font_name=b_font, font_size=10.5, bold=False, color_hex=body_color
                        )
                        _fmt_last_para(0, 3, 1.25, left_indent=14.0, first_indent=-10.0)
                        (
                        )
                    elif etype == "CODE_BLOCK":
                        if text_val:
                            word_write_text(
                                f"※ {text_val}", font_name=b_font, font_size=10, bold=False, color_hex="#64748B"
                            )
                    elif etype == "HORIZONTAL_RULE":
                        # 마크다운 수평선(---, ***) -> Word 하단 실선 구분선 문단으로 렌더링
                        try:
                            sel = word_app.Selection
                            sel.EndKey(Unit=6)  # wdStory (문서 끝으로 이동)
                            p = doc.Paragraphs.Add()
                            border = p.Range.Borders(-3)  # wdBorderBottom
                            border.LineStyle = 1  # wdLineStyleSingle
                            border.LineWidth = 6  # 0.75pt
                            border.Color = 0xD9D9D9  # 밝은 회색선
                            p.Range.ParagraphFormat.SpaceBefore = 8
                            p.Range.ParagraphFormat.SpaceAfter = 8
                            cp_hr = doc.Paragraphs(doc.Paragraphs.Count).Range
                            cp_hr.Borders(-3).LineStyle = 0
                            cp_hr.ParagraphFormat.LeftIndent = 0
                            cp_hr.ParagraphFormat.FirstLineIndent = 0
                        except Exception:
                            pass
                    else:
                        if text_val:
                            word_write_text(
                                text_val, font_name=b_font, font_size=10.5, bold=False, color_hex=body_color
                            )
                            _fmt_last_para(0, 4, 1.28, left_indent=0.0, first_indent=0.0)
                            (
                            )
                    rendered_count += 1
                    idx += 1
            return {
                "success": True,
                "document_name": doc.Name,
                "theme": theme,
                "elements_rendered": rendered_count,
            }
        except Exception as e:
            return {"success": False, "error": f"마크다운 보고서 렌더링 오류: {e}"}

    @mcp.tool()
    def word_find_and_replace(
        find_text: str,
        replace_text: str,
        replace_all: bool = True,
    ) -> dict[str, Any]:
        """현재 열려 있는 Word 문서 본문에서 특정 텍스트를 찾아 일괄 변경합니다."""
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "열려 있는 Word 문서가 없습니다."}

        try:
            with word_undo_group(word_app, "MCP 찾기 및 바꾸기"):
                find_obj = doc.Content.Find
                find_obj.ClearFormatting()
                find_obj.Replacement.ClearFormatting()
                # Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
                #         MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
                # wdReplaceAll = 2, wdReplaceOne = 1
                ok = find_obj.Execute(
                    find_text, False, False, False, False, False, True, 1, False, replace_text, 2 if replace_all else 1
                )
            return {"success": True, "replaced": bool(ok), "find_text": find_text, "replace_text": replace_text}
        except Exception as e:
            return {"success": False, "error": f"찾기/바꾸기 실패: {e}"}

    @mcp.tool()
    def word_save_document(file_path: str | None = None) -> dict[str, Any]:
        """현재 활성 Word 문서를 저장합니다. file_path 지정 시 다른 이름으로 저장(SaveAs2)합니다."""
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "열려 있는 Word 문서가 없습니다."}
        try:
            if file_path:
                abs_path = str(Path(file_path).resolve())
                Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
                prev_alerts = word_app.DisplayAlerts
                try:
                    word_app.DisplayAlerts = 0
                    doc.SaveAs2(abs_path, FileFormat=16)
                finally:
                    word_app.DisplayAlerts = prev_alerts
                return {"success": True, "path": abs_path}
            doc.Save()
            return {"success": True, "path": doc.FullName}
        except Exception as e:
            return {"success": False, "error": f"Word 저장 실패: {e}"}
