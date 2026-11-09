# -*- coding: utf-8 -*-
"""
mcp_server/tools/hwp_tools.py
한컴오피스(HWP) 실시간 ROT(!HwpObject) 바인딩 기반:
- 전체 문서 및 마우스 선택 블록(saveblock:true) 읽기
- 행안부 / KISA 범피스 1310 공기업 표준 서식 / 비즈니스 서식 프리셋 적용 및 대필
- KISA 1310 표준 공문서 마크업(표지:, 제목:, 네모:, 원:, 바:, 별:, 표:) 자동 렌더링
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr
from .parsers import parse_document_elements, parse_inline_spans, strip_markdown_syntax

# 행안부 / KISA 1310 공기업 / 비즈니스 HWP 서식 프리셋
HWP_STYLE_PRESETS: dict[str, dict[str, Any]] = {
    "행안제목": {"font": "HY헤드라인M", "size": 18, "bold": True, "align": "center", "spacing": 0, "color_hex": "#000000"},
    "행안요약": {"font": "함초롬바탕", "size": 12, "bold": True, "align": "justify", "spacing": -5, "color_hex": "#000000"},
    "행안보고자": {"font": "맑은 고딕", "size": 10, "bold": False, "align": "right", "spacing": 0, "color_hex": "#000000"},
    "행안본문1": {"font": "휴먼명조", "size": 15, "bold": True, "align": "justify", "spacing": -2, "color_hex": "#000000"},
    "행안본문2": {"font": "휴먼명조", "size": 14, "bold": False, "align": "justify", "spacing": -2, "color_hex": "#000000"},
    "행안본문3": {"font": "휴먼명조", "size": 13, "bold": False, "align": "justify", "spacing": -2, "color_hex": "#000000"},
    "공기업제목": {"font": "맑은 고딕", "size": 18, "bold": True, "align": "center", "spacing": -3, "color_hex": "#1F3864"},
    "공기업본문": {"font": "맑은 고딕", "size": 11, "bold": False, "align": "justify", "spacing": -2, "color_hex": "#000000"},
    "비즈니스제목": {"font": "맑은 고딕", "size": 18, "bold": True, "align": "left", "spacing": 0, "color_hex": "#1F3864"},
    "비즈니스소제목": {"font": "맑은 고딕", "size": 14, "bold": True, "align": "left", "spacing": 0, "color_hex": "#2E608A"},
    "비즈니스본문": {"font": "맑은 고딕", "size": 11, "bold": False, "align": "justify", "spacing": 0, "color_hex": "#1E293B"},
}

HWP_ALIGN_ACTIONS = {
    "left": "ParagraphShapeAlignLeft",
    "center": "ParagraphShapeAlignCenter",
    "right": "ParagraphShapeAlignRight",
    "justify": "ParagraphShapeAlignJustify",
    "distribute": "ParagraphShapeAlignDistribute",
}

# KISA 범피스 1310 한컴라이브러리_3.py 에서 복원한 공문서/보고서 작성 표준 프롬프트 규칙
KISA1310_OFFICIAL_REPORT_GUIDE = """[KISA 범피스 1310 표준 공문서/보고서 마크업 규격]
각 줄의 콜론(:) 앞 기호를 인식하여 한글(HWP) 표준 서식과 표로 자동 변환합니다:
- 제목: [보고서 대제목] -> HY헤드라인M 18pt 굵게 가운데 정렬
- 부서: [담당부서명] -> 맑은 고딕 10pt 오른쪽 정렬
- 네모: [대항목 제목, 4~8글자] -> '□ ' 기호 + 휴먼명조 15pt 굵게
- 원: [세부 추진 내용] -> '  ○ ' 기호 + 휴먼명조 14pt
- 바: [구체적 실행 설명] -> '    - ' 기호 + 휴먼명조 13pt
- 별: [강조/보충 설명] -> '    ※ ' 기호 + 맑은 고딕 11pt
- 표:[컬럼1]:[컬럼2]:...:[비고] (연속된 '표:' 줄들을 묶어 한글 표 및 헤더 음영으로 자동 생성)
"""


def _apply_hwp_char_and_para(
    hwp: Any,
    font_name: str = "휴먼명조",
    font_size: float = 12.0,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    color_hex: str = "#000000",
    spacing: int = 0,
    align: str = "justify",
) -> None:
    act = hwp.CreateAction("CharShape")
    pset = act.CreateSet()
    act.GetDefault(pset)
    for face_key in ("FaceNameHangul", "FaceNameLatin", "FaceNameHanja", "FaceNameJapanese", "FaceNameOther", "FaceNameSymbol", "FaceNameUser"):
        pset.SetItem(face_key, font_name)
    pset.SetItem("Height", int(float(font_size) * 100))
    pset.SetItem("Bold", 1 if bold else 0)
    pset.SetItem("Italic", 1 if italic else 0)
    pset.SetItem("UnderlineType", 1 if underline else 0)
    pset.SetItem("TextColor", hex_to_bgr(color_hex))
    pset.SetItem("SpacingHangul", int(spacing))
    pset.SetItem("SpacingLatin", int(spacing))
    act.Execute(pset)

    align_act = HWP_ALIGN_ACTIONS.get(align.lower(), "ParagraphShapeAlignJustify")
    hwp.HAction.Run(align_act)


def _hwp_insert_raw_text(hwp: Any, text: str) -> None:
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)


def register_hwp_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def hwp_read_document(
        only_selection: bool = False,
        format_type: Literal["TEXT", "HWPML2X"] = "TEXT",
        max_chars: int = 12000,
    ) -> dict[str, Any]:
        """
        현재 열려 있는 한컴오피스(HWP) 창(ROT 실시간 바인딩)에서 문서 내용을 읽어옵니다.
        - only_selection=False: 문서 전체 본문 텍스트를 추출합니다.
        - only_selection=True: 사용자가 마우스로 드래그/선택한 블록(`saveblock:true`)만 정확히 추출합니다.
        - format_type="HWPML2X": 표와 서식 구조를 포함한 XML(HML) 형식으로 추출합니다.
        """
        ensure_com()
        hwp = com_manager.get_hwp(create_if_missing=False)
        if hwp is None:
            return {"success": False, "error": "현재 실행 중인 한컴오피스(HWP) 창을 찾을 수 없습니다."}

        try:
            option = "saveblock:true" if only_selection else ""
            raw_text = hwp.GetTextFile(format_type, option) or ""
            cleaned = raw_text.replace("\r\n", "\n").strip()
            return {
                "success": True,
                "page_count": hwp.PageCount,
                "only_selection": only_selection,
                "format_type": format_type,
                "text": cleaned[:max_chars],
                "truncated": len(cleaned) > max_chars,
            }
        except Exception as e:
            return {"success": False, "error": f"HWP 문서 읽기 실패: {e}"}

    @mcp.tool()
    def hwp_write_styled_text(
        text: str,
        style_preset: Literal[
            "행안제목", "행안요약", "행안보고자", "행안본문1", "행안본문2", "행안본문3",
            "공기업제목", "공기업본문", "비즈니스제목", "비즈니스소제목", "비즈니스본문"
        ] | None = "행안본문2",
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        color_hex: str | None = None,
        align: Literal["left", "center", "right", "justify", "distribute"] | None = None,
        new_para: bool = True,
    ) -> dict[str, Any]:
        """
        현재 열려 있는 한컴오피스(HWP) 커서 위치에 서식(행안부 공문서 표준, KISA 1310 공기업 서식, 비즈니스 서식 등)을
        적용하여 텍스트를 실시간으로 입력합니다.
        """
        ensure_com()
        hwp = com_manager.get_hwp(create_if_missing=True, visible=True)
        if hwp is None:
            return {"success": False, "error": "HWP 연결 실패"}

        preset = HWP_STYLE_PRESETS.get(style_preset or "", HWP_STYLE_PRESETS["행안본문2"])
        f_name = font_name or preset["font"]
        f_size = font_size if font_size is not None else preset["size"]
        f_bold = bold if bold is not None else preset["bold"]
        f_color = color_hex or preset["color_hex"]
        f_align = align or preset["align"]
        f_spacing = preset.get("spacing", 0)

        try:
            _apply_hwp_char_and_para(
                hwp,
                font_name=f_name,
                font_size=f_size,
                bold=f_bold,
                color_hex=f_color,
                spacing=f_spacing,
                align=f_align,
            )
            # 인라인 마크다운(**볼드**, *이탤릭*) 파싱 후 HWP 서식 타이핑
            spans = parse_inline_spans(text)
            for span in spans:
                s_text = span.get("text", "")
                if not s_text:
                    continue
                is_b = span.get("bold", False) or f_bold
                is_i = span.get("italic", False)
                if is_b != f_bold or is_i:
                    _apply_hwp_char_and_para(
                        hwp,
                        font_name=f_name,
                        font_size=f_size,
                        bold=is_b,
                        italic=is_i,
                        color_hex=f_color,
                        spacing=f_spacing,
                        align=f_align,
                    )
                _hwp_insert_raw_text(hwp, s_text)
                if is_b != f_bold or is_i:
                    _apply_hwp_char_and_para(
                        hwp,
                        font_name=f_name,
                        font_size=f_size,
                        bold=f_bold,
                        color_hex=f_color,
                        spacing=f_spacing,
                        align=f_align,
                    )
            if new_para:
                hwp.HAction.Run("BreakPara")
            return {"success": True, "style_preset": style_preset, "written_length": len(text)}
        except Exception as e:
            return {"success": False, "error": f"HWP 텍스트 입력 오류: {e}"}

    @mcp.tool()
    def hwp_format_selection(
        style_preset: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        color_hex: str | None = None,
        align: Literal["left", "center", "right", "justify", "distribute"] | None = None,
        select_current_line_if_empty: bool = True,
    ) -> dict[str, Any]:
        """
        한컴오피스(HWP)에서 현재 블록 선택된 영역(또는 현재 줄)에 행안부/공기업/비즈니스 서식을 즉시 적용합니다.
        """
        ensure_com()
        hwp = com_manager.get_hwp(create_if_missing=False)
        if hwp is None:
            return {"success": False, "error": "열려 있는 HWP 창이 없습니다."}

        preset = HWP_STYLE_PRESETS.get(style_preset or "", {})
        try:
            sel_text = (hwp.GetTextFile("TEXT", "saveblock:true") or "").strip()
            if not sel_text and select_current_line_if_empty:
                hwp.HAction.Run("MoveLineBegin")
                hwp.HAction.Run("Select")
                hwp.HAction.Run("MoveLineEnd")

            _apply_hwp_char_and_para(
                hwp,
                font_name=font_name or preset.get("font", "휴먼명조"),
                font_size=font_size if font_size is not None else preset.get("size", 12),
                bold=bold if bold is not None else preset.get("bold", False),
                color_hex=color_hex or preset.get("color_hex", "#000000"),
                spacing=preset.get("spacing", 0),
                align=align or preset.get("align", "justify"),
            )
            hwp.HAction.Run("Cancel")
            return {"success": True, "style_preset": style_preset}
        except Exception as e:
            return {"success": False, "error": f"HWP 선택 영역 서식 적용 오류: {e}"}

    @mcp.tool()
    def hwp_insert_table(
        data: list[list[Any]],
        header_style: Literal["official_gray", "navy"] = "official_gray",
    ) -> dict[str, Any]:
        """
        현재 열려 있는 한컴오피스(HWP) 커서 위치에 표(Table)를 생성하고 2차원 배열 데이터를 각 셀에 채워 넣습니다.
        """
        ensure_com()
        if not data or not data[0]:
            return {"success": False, "error": "표 데이터(data)가 비어 있습니다."}

        hwp = com_manager.get_hwp(create_if_missing=True, visible=True)
        if hwp is None:
            return {"success": False, "error": "HWP 연결 실패"}

        rows_cnt = len(data)
        cols_cnt = max(len(r) for r in data)
        try:
            pset = hwp.HParameterSet.HTableCreation
            hwp.HAction.GetDefault("TableCreate", pset.HSet)
            pset.Rows = rows_cnt
            pset.Cols = cols_cnt
            pset.WidthType = 2
            pset.HeightType = 0
            pset.TableProperties.TreatAsChar = 1  # 글자처럼 취급
            hwp.HAction.Execute("TableCreate", pset.HSet)

            # 열별 데이터 성격 분석 (순번/코드 중앙, 수량/금액 우측, 텍스트 좌측)
            col_aligns = []
            for c_i in range(cols_cnt):
                c_vals = [str(data[r][c_i]).strip() for r in range(1, rows_cnt) if c_i < len(data[r])]
                hdr_val = str(data[0][c_i]).strip().upper() if c_i < len(data[0]) else ""
                is_idx = any(k in hdr_val for k in ["NO", "순번", "연번", "번호", "단계"]) or all(v.isdigit() and len(v) <= 3 for v in c_vals if v)
                has_num = any(sym in "".join(c_vals) for sym in [",", "%", "원", "$", "₩"]) or (sum(1 for v in c_vals if re.match(r'^-?\d+(?:\.\d+)?$', re.sub(r'[,%\s원\$₩]', '', v))) >= len(c_vals) * 0.7 if c_vals else False)
                if is_idx:
                    col_aligns.append("center")
                elif has_num:
                    col_aligns.append("right")
                elif max((len(v) for v in c_vals), default=0) <= 8:
                    col_aligns.append("center")
                else:
                    col_aligns.append("left")

            for r_idx, row_vals in enumerate(data):
                for c_idx in range(cols_cnt):
                    raw_val = str(row_vals[c_idx]) if c_idx < len(row_vals) else ""
                    val = strip_markdown_syntax(raw_val)
                    if r_idx == 0:
                        _apply_hwp_char_and_para(
                            hwp,
                            font_name="맑은 고딕",
                            font_size=10,
                            bold=True,
                            color_hex="#000000",
                            align="center",
                        )
                    else:
                        _apply_hwp_char_and_para(
                            hwp,
                            font_name="함초롬바탕",
                            font_size=10,
                            bold=False,
                            color_hex="#000000",
                            align=col_aligns[c_idx] if c_idx < len(col_aligns) else "left",
                        )
                    _hwp_insert_raw_text(hwp, val)
                    if not (r_idx == rows_cnt - 1 and c_idx == cols_cnt - 1):
                        hwp.HAction.Run("TableRightCell")

            # 표 밖으로 빠져나와 줄바꿈
            hwp.HAction.Run("CloseEx")
            hwp.HAction.Run("MoveLineEnd")
            hwp.HAction.Run("BreakPara")

            return {"success": True, "rows": rows_cnt, "cols": cols_cnt}
        except Exception as e:
            return {"success": False, "error": f"HWP 표 생성 오류: {e}"}

    @mcp.tool()
    def hwp_render_kisa1310_report(
        report_text: str,
    ) -> dict[str, Any]:
        """
        KISA 범피스 1310의 공문서/계획서 마크업 규격(`제목:`, `부서:`, `네모:`, `원:`, `바:`, `별:`, `표:`)으로
        작성된 텍스트를 파싱하여 현재 열려 있는 한컴오피스(HWP) 문서에 행안부/공문서 표준 서식과 표로 자동 작성합니다.

        예시 입력:
        제목: 범정부 오피스 AI 자동화 추진 계획
        부서: 디지털혁신담당관
        네모: 추진 배경 및 목적
        원: 반복적인 공문서 서식 편집 시간 단축 및 업무 생산성 제고
        바: 실시간 COM 제어 기반 AI 어시스턴트 연동
        네모: 세부 추진 일정
        표:연번:추진단계:주요내용:비고
        표:1:기반구축:MCP 서버 및 서식 엔진 구축:완료
        표:2:고도화:전 부서 배포 및 사용자 교육:추진중
        """
        ensure_com()
        hwp = com_manager.get_hwp(create_if_missing=True, visible=True)
        if hwp is None:
            return {"success": False, "error": "HWP 연결 실패"}

        elements = parse_document_elements(report_text)
        table_buffer: list[list[str]] = []
        rendered_items = 0

        def flush_table() -> None:
            nonlocal table_buffer, rendered_items
            if table_buffer:
                hwp_insert_table(table_buffer)
                rendered_items += 1
                table_buffer = []

        try:
            for elem in elements:
                etype = elem.get("type")
                text_val = str(elem.get("text") or "").strip()

                if etype == "TABLE_ROW":
                    cells = elem.get("cells", [])
                    if cells:
                        table_buffer.append(cells)
                    continue
                elif etype == "TABLE_SEP":
                    continue

                flush_table()

                if etype == "TITLE_BOX":
                    hwp_write_styled_text(text_val, style_preset="행안제목")
                elif etype == "SQUARE_HEADER":
                    prefix = "" if text_val.startswith(("□", "■", "1", "2", "3", "가", "나")) else "□ "
                    hwp_write_styled_text(prefix + text_val, style_preset="행안본문1")
                elif etype == "CIRCLE_ITEM":
                    prefix = "" if text_val.startswith(("○", "●", "①", "②", "③")) else "  ○ "
                    hwp_write_styled_text(prefix + text_val, style_preset="행안본문2")
                elif etype == "DASH_ITEM":
                    prefix = "" if text_val.startswith(("-", "–", "※", "☐", "☑")) else "    - "
                    hwp_write_styled_text(prefix + text_val, style_preset="행안본문3")
                elif etype == "CODE_BLOCK":
                    if text_val:
                        hwp_write_styled_text(f"    ※ {text_val}", style_preset="공기업본문")
                elif etype == "HORIZONTAL_RULE":
                    # 한글 구분선 문단
                    hwp_write_styled_text("──────────────────────────────────────────", style_preset="공기업본문", align="center")
                else:
                    if text_val:
                        hwp_write_styled_text(text_val, style_preset="행안본문2")
                rendered_items += 1

            flush_table()
            return {"success": True, "rendered_items": rendered_items, "page_count": hwp.PageCount}
        except Exception as e:
            return {"success": False, "error": f"KISA 1310 공문서 렌더링 실패: {e}"}

    @mcp.tool()
    def hwp_save_document(file_path: str | None = None) -> dict[str, Any]:
        """현재 열려 있는 한컴오피스(HWP) 문서를 저장합니다."""
        ensure_com()
        hwp = com_manager.get_hwp(create_if_missing=False)
        if hwp is None:
            return {"success": False, "error": "열려 있는 HWP 창이 없습니다."}
        try:
            if file_path:
                abs_path = str(Path(file_path).resolve())
                Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
                hwp.SaveAs(abs_path)
                return {"success": True, "path": abs_path}
            hwp.Save()
            return {"success": True, "path": getattr(hwp, "Path", "")}
        except Exception as e:
            return {"success": False, "error": f"HWP 저장 실패: {e}"}
