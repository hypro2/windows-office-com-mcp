# -*- coding: utf-8 -*-
"""
mcp_server/tools/ppt_tools.py
MS PowerPoint 실시간 읽기/편집 및 비즈니스 프레젠테이션 생성 도구.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr
from .parsers import strip_markdown_syntax
from .ppt_slice_engine import PptSliceEngine, SliceRegionCalculator, CATALOG_DATA
from .ppt_recipe_engine import PptRecipeEngine


def _extract_slide_summary(slide: Any, idx: int) -> dict[str, Any]:
    shapes_info: list[dict[str, Any]] = []
    texts: list[str] = []
    for sh_idx in range(1, slide.Shapes.Count + 1):
        try:
            sh = slide.Shapes(sh_idx)
            sh_text = ""
            if sh.HasTextFrame and sh.TextFrame.HasText:
                sh_text = (sh.TextFrame.TextRange.Text or "").replace("\r", "\n").strip()
            elif getattr(sh, "HasTable", 0):
                t_rows: list[str] = []
                tbl = sh.Table
                for r in range(1, tbl.Rows.Count + 1):
                    r_vals = [
                        (tbl.Cell(r, c).Shape.TextFrame.TextRange.Text or "").replace("\r", " ").strip()
                        for c in range(1, tbl.Columns.Count + 1)
                    ]
                    t_rows.append("| " + " | ".join(r_vals) + " |")
                sh_text = "\n".join(t_rows)
            if sh_text:
                texts.append(sh_text)
                shapes_info.append({"index": sh_idx, "name": sh.Name, "text": sh_text[:500]})
        except Exception:
            continue
    return {
        "slide_index": idx,
        "layout_name": getattr(slide.CustomLayout, "Name", ""),
        "shapes_count": slide.Shapes.Count,
        "text_summary": "\n".join(texts)[:1500],
        "shapes": shapes_info,
    }


# =============================================================================
# COM 도형(Shape) 상수 및 매핑
# =============================================================================
import ctypes
from contextlib import contextmanager
import logging

logger = logging.getLogger("mcp_server.ppt_tools")
user32 = ctypes.windll.user32


@contextmanager
def ppt_freeze_display(ppt_app: Any):
    """대량 도형 생성/삭제/이동 시 화면 갱신을 동결(LockWindowUpdate)하여 깜빡임 방지 및 3배 가속."""
    hwnd = 0
    try:
        if ppt_app and hasattr(ppt_app, "Caption"):
            caption = getattr(ppt_app, "Caption", None)
            if isinstance(caption, str) and caption:
                hwnd = user32.FindWindowW(None, caption)
        if hwnd:
            user32.LockWindowUpdate(hwnd)
    except Exception:
        hwnd = 0
    try:
        yield
    finally:
        if hwnd:
            try:
                user32.LockWindowUpdate(0)
            except Exception:
                pass


MSO_SHAPE_MAP: dict[str, int] = {
    "rect": 1,
    "rectangle": 1,
    "rounded_rect": 5,
    "rounded_rectangle": 5,
    "diamond": 4,
    "oval": 9,
    "circle": 9,
    "right_arrow": 33,
    "arrow": 33,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "chevron": 52,
}

ALIGN_MAP: dict[str, int] = {
    "left": 1,
    "center": 2,
    "right": 3,
    "justify": 4,
}

V_ANCHOR_MAP: dict[str, int] = {
    "top": 1,
    "middle": 3,
    "bottom": 4,
}

ZORDER_MAP: dict[str, int] = {
    "front": 0,    # msoBringToFront
    "back": 1,     # msoSendToBack
    "forward": 2,  # msoBringForward
    "backward": 3, # msoSendBackward
}

ALIGN_MSO_MAP: dict[str, int] = {
    "left": 0,    # msoAlignLefts
    "center": 1,  # msoAlignCenters
    "right": 2,   # msoAlignRights
    "top": 3,     # msoAlignTops
    "middle": 4,  # msoAlignMiddles
    "bottom": 5,  # msoAlignBottoms
}

DISTRIBUTE_MSO_MAP: dict[str, int] = {
    "horizontally": 0,  # msoDistributeHorizontally
    "vertically": 1,    # msoDistributeVertically
}

DEFAULT_ROADMAP_PALETTE: list[str] = [
    "#1F3864",  # Business Navy
    "#2E608A",  # Steel Blue
    "#0F766E",  # Modern Teal
    "#0284C7",  # Sky Blue
    "#D97706",  # Amber
    "#7C3AED",  # Purple
    "#475569",  # Slate
    "#E11D48",  # Rose
]


def _resolve_target_slide(
    ppt_app: Any,
    prs: Any,
    slide_index: int | None = None,
    create_new_slide: bool = False,
) -> tuple[Any, int]:
    """대상 슬라이드 객체 및 1-based 인덱스 반환."""
    if create_new_slide:
        new_idx = prs.Slides.Count + 1
        slide = prs.Slides.Add(new_idx, 12)  # 12 = ppLayoutBlank
        try:
            ppt_app.ActiveWindow.View.GotoSlide(new_idx)
        except Exception:
            pass
        return slide, new_idx

    if slide_index is not None:
        if slide_index < 1 or slide_index > prs.Slides.Count:
            raise ValueError(f"슬라이드 번호({slide_index})가 유효 범위(1~{prs.Slides.Count})를 벗어났습니다.")
        slide = prs.Slides(slide_index)
        try:
            ppt_app.ActiveWindow.View.GotoSlide(slide_index)
        except Exception:
            pass
        return slide, slide_index

    try:
        slide = ppt_app.ActiveWindow.View.Slide
        return slide, slide.SlideIndex
    except Exception:
        if prs.Slides.Count > 0:
            slide = prs.Slides(1)
            return slide, 1
        slide = prs.Slides.Add(1, 12)
        return slide, 1


def _apply_shape_styling(shp: Any, spec: dict[str, Any]) -> None:
    """도형(Shape/Textbox)의 채우기, 선, 텍스트, 정렬 서식 일괄 적용."""
    # 1. Fill
    fill_color = spec.get("fill_color")
    if fill_color is not None:
        fill_lower = str(fill_color).strip().lower()
        if fill_lower in ("none", "transparent"):
            shp.Fill.Visible = 0
        else:
            shp.Fill.Visible = -1
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = hex_to_bgr(fill_color)

    # 2. Line (Border)
    line_color = spec.get("line_color")
    if line_color is not None:
        line_lower = str(line_color).strip().lower()
        if line_lower in ("none", "transparent"):
            shp.Line.Visible = 0
        else:
            shp.Line.Visible = -1
            shp.Line.ForeColor.RGB = hex_to_bgr(line_color)
            if "line_weight" in spec and spec["line_weight"] is not None:
                shp.Line.Weight = float(spec["line_weight"])
    elif spec.get("line_visible") is False:
        shp.Line.Visible = 0

    # 3. Shadow
    if not spec.get("shadow", False):
        try:
            shp.Shadow.Visible = 0
        except Exception:
            pass
    else:
        try:
            shp.Shadow.Visible = -1
        except Exception:
            pass

    # 4. Text & TextFrame
    text = spec.get("text")
    if text is not None:
        tf = shp.TextFrame
        tf.WordWrap = 1 if spec.get("word_wrap", True) else 0

        v_align = str(spec.get("vertical_align", "middle")).lower()
        if v_align in V_ANCHOR_MAP:
            tf.VerticalAnchor = V_ANCHOR_MAP[v_align]

        # 기본 여백(margin): 텍스트가 테두리에 달라붙지 않도록 좌우 8pt, 상하 6pt 안전 여백 기본 적용
        default_pad_x = 8.0
        default_pad_y = 6.0
        if "margin" in spec and spec["margin"] is not None:
            m_val = float(spec["margin"])
            tf.MarginLeft = m_val
            tf.MarginRight = m_val
            tf.MarginTop = m_val
            tf.MarginBottom = m_val
        else:
            tf.MarginLeft = float(spec.get("margin_left", default_pad_x))
            tf.MarginRight = float(spec.get("margin_right", default_pad_x))
            tf.MarginTop = float(spec.get("margin_top", default_pad_y))
            tf.MarginBottom = float(spec.get("margin_bottom", default_pad_y))

        tr = tf.TextRange
        tr.Text = strip_markdown_syntax(str(text)).replace("\n", "\r")

        font_name = spec.get("font_name", "맑은 고딕")
        tr.Font.Name = font_name
        try:
            tr.Font.NameFarEast = font_name
        except Exception:
            pass

        if spec.get("font_size") is not None:
            tr.Font.Size = float(spec["font_size"])

        if "bold" in spec and spec["bold"] is not None:
            tr.Font.Bold = -1 if spec["bold"] else 0

        font_color = spec.get("font_color")
        if font_color:
            tr.Font.Color.RGB = hex_to_bgr(font_color)

        align = str(spec.get("align", "left")).lower()
        if align in ALIGN_MAP:
            tr.ParagraphFormat.Alignment = ALIGN_MAP[align]

        # 가독성 개선: 한국어 단어 단위 워드랩 및 행간/문단간격 기본 적용
        try:
            tr.LanguageID = 1042  # msoLanguageIDKorean
        except Exception:
            pass

        try:
            l_spacing = float(spec.get("line_spacing", 1.2))
            tr.ParagraphFormat.LineRuleWithin = -1  # msoTrue (줄 단위)
            tr.ParagraphFormat.SpaceWithin = l_spacing

            sp_after = float(spec.get("space_after", 4.0))
            tr.ParagraphFormat.LineRuleAfter = 0   # msoFalse (포인트 단위)
            tr.ParagraphFormat.SpaceAfter = sp_after
        except Exception:
            pass


def _estimate_text_width_pt(text: str, font_size_pt: float) -> float:
    """텍스트의 1줄 필요 너비(pt) 추정. 한글/전각=1.0em, 대문자=0.75em, 숫자/소문자=0.6em, 공백=0.35em."""
    units = 0.0
    for char in text:
        code = ord(char)
        if char == " ":
            units += 0.35
        elif code > 127:  # 한글, 전각 기호 등
            units += 1.02
        elif char.isupper():
            units += 0.75
        elif char.isdigit():
            units += 0.65
        else:
            units += 0.55
    return units * font_size_pt


def _create_single_shape(slide: Any, spec: dict[str, Any]) -> Any:
    """단일 도형 또는 텍스트박스 생성 후 서식 적용 (너비 부족으로 인한 강제 줄바꿈 자동 방어)."""
    stype_raw = spec.get("type", "rect")
    left = float(spec.get("left", 0))
    top = float(spec.get("top", 0))
    width = float(spec.get("width", 100))
    height = float(spec.get("height", 50))

    # [가독성/줄바꿈 방어] 텍스트가 박스 너비보다 길어 강제 줄바꿈되는 현상 방어
    text = spec.get("text")
    if text:
        raw_lines = [line.strip() for line in str(text).split("\n") if line.strip()]
        longest_line = max(raw_lines, key=len) if raw_lines else ""
        font_size = float(spec.get("font_size", 11.0))
        pad_x = float(spec.get("margin_left", spec.get("margin", 8.0))) + float(spec.get("margin_right", spec.get("margin", 8.0)))
        needed_w = _estimate_text_width_pt(longest_line, font_size) + pad_x + 6.0

        # auto_fit_width(기본 True): 텍스트에 비해 너비가 좁으면 슬라이드 허용 범위 내에서 너비 자동 확장
        auto_fit_width = spec.get("auto_fit_width", True)
        if auto_fit_width and width < needed_w:
            slide_w = 960.0
            try:
                slide_w = float(getattr(slide.Parent.PageSetup, "SlideWidth", 960.0))
            except Exception:
                pass
            max_avail_w = max(50.0, slide_w - left - 15.0)
            if needed_w <= max_avail_w:
                width = needed_w
            else:
                width = max_avail_w
                # 공간이 부족한 경우 폰트를 축소해 1줄 맞춤 유지
                scale = max(0.1, (max_avail_w - pad_x - 6.0) / max(1.0, _estimate_text_width_pt(longest_line, font_size)))
                new_font_size = max(8.5, round(font_size * scale * 2) / 2)
                spec["font_size"] = new_font_size

    if isinstance(stype_raw, str):
        stype_str = stype_raw.strip().lower()
        if stype_str in ("delete", "remove", "del"):
            raise ValueError("도형 삭제는 'ppt_delete_shapes' 도구를 사용해야 합니다. (잘못된 사각형 박스 생성 차단)")
        if stype_str in ("move", "resize", "transform"):
            raise ValueError("도형 이동 및 크기 변경은 'ppt_transform_shapes' 도구를 사용해야 합니다. (잘못된 사각형 박스 생성 차단)")
        if stype_str in ("textbox", "text_box", "text"):
            shp = slide.Shapes.AddTextbox(1, left, top, width, height)  # 1 = msoTextOrientationHorizontal
            _apply_shape_styling(shp, spec)
            return shp
        mso_type = MSO_SHAPE_MAP.get(stype_str, 1)
    elif isinstance(stype_raw, int):
        mso_type = stype_raw
    else:
        mso_type = 1

    shp = slide.Shapes.AddShape(mso_type, left, top, width, height)
    _apply_shape_styling(shp, spec)
    return shp


def register_ppt_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def ppt_read_presentation(
        slide_index: int | None = None,
        only_selection: bool = False,
    ) -> dict[str, Any]:
        """
        현재 데스크톱에 열려 있는 MS PowerPoint 프레젠테이션의 슬라이드 목록과 텍스트/표 내용을 읽어옵니다.
        - slide_index: 특정 슬라이드 번호(1부터 시작)만 상세 조회할 때 지정 (생략 시 전체 슬라이드 조회)
        - only_selection=True: 현재 사용자가 선택한 슬라이드 또는 도형(Shape)의 텍스트만 조회
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "현재 열려 있는 MS PowerPoint 프레젠테이션이 없습니다."}

        try:
            if only_selection:
                sel = ppt_app.ActiveWindow.Selection
                # ppSelectionText=3, ppSelectionShapes=2, ppSelectionSlides=1
                if sel.Type == 3:
                    return {"success": True, "selection_type": "text", "text": sel.TextRange.Text}
                elif sel.Type == 2:
                    sh_texts = []
                    for i in range(1, sel.ShapeRange.Count + 1):
                        sh = sel.ShapeRange(i)
                        if sh.HasTextFrame and sh.TextFrame.HasText:
                            sh_texts.append(sh.TextFrame.TextRange.Text)
                    return {"success": True, "selection_type": "shapes", "texts": sh_texts}
                elif sel.Type == 1 and sel.SlideRange.Count > 0:
                    s = sel.SlideRange(1)
                    return {"success": True, "selection_type": "slide", "slide": _extract_slide_summary(s, s.SlideIndex)}

            if slide_index is not None:
                if slide_index < 1 or slide_index > prs.Slides.Count:
                    return {"success": False, "error": f"슬라이드 번호 범위(1~{prs.Slides.Count})를 벗어났습니다."}
                return {
                    "success": True,
                    "presentation_name": prs.Name,
                    "slides_count": prs.Slides.Count,
                    "slide": _extract_slide_summary(prs.Slides(slide_index), slide_index),
                }

            slides_data = [
                _extract_slide_summary(prs.Slides(i), i) for i in range(1, prs.Slides.Count + 1)
            ]
            return {
                "success": True,
                "presentation_name": prs.Name,
                "slides_count": prs.Slides.Count,
                "slides": slides_data,
            }
        except Exception as e:
            return {"success": False, "error": f"PowerPoint 읽기 실패: {e}"}

    @mcp.tool()
    def ppt_create_report_deck_shell(
        title: str,
        subtitle: str = "",
        department: str = "기획팀",
        report_date: str = "",
        output_path: str = "output_presentation.pptx",
        open_in_desktop: bool = True,
        theme: Literal[
            "business_navy", "steel_blue", "modern_teal", "clean_charcoal",
            "emerald_green", "warm_beige", "indigo_violet"
        ] = "business_navy",
    ) -> dict[str, Any]:
        """
        [오피스 단계별 분할 생성 1단계: 기본 틀(Shell) 생성 및 즉시 데스크톱 팝업]
        기본값 정책: 표준 비즈니스 네이비(business_navy) 또는 7대 범용 비즈니스 테마 기반 기본 틀(Cover Shell)을 생성합니다.
        """
        # 글로벌 표준 비즈니스 표지 외피(Cover Shell) 순수 COM 생성
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 창을 열 수 없습니다."}

        try:
            prs.PageSetup.SlideWidth = 960.0
            prs.PageSetup.SlideHeight = 540.0
            slide = prs.Slides.Add(1, 12)  # ppLayoutBlank
            resolved_theme = PptSliceEngine.resolve_theme(theme)
            PptSliceEngine.render_shell_frame(
                slide,
                frame_type="cover",
                data={
                    "title": title,
                    "subtitle": subtitle,
                    "category": "BUSINESS REPORT",
                    "author": department,
                    "date": report_date or "2026. 11",
                },
                theme=resolved_theme,
                slide_w=960.0,
                slide_h=540.0
            )
            safe_out = com_manager.prepare_output_path(output_path)
            prs.SaveAs(str(safe_out))
            return {
                "success": True,
                "theme": theme,
                "path": str(safe_out),
                "slides_count": prs.Slides.Count,
                "guide": (
                    f"[{theme} 비즈니스 마스터 틀 생성 완료] PowerPoint 창이 즉시 활성화되었습니다. "
                    "이제 1장씩 'ppt_compose_slide' 또는 'ppt_append_report_slide'로 본문 슬라이스를 주입하세요."
                )
            }
        except Exception as e:
            logger.error(f"Failed to create deck shell: {e}")
            return {"success": False, "error": str(e)}

    def ppt_append_report_slide(
        title: str,
        section: str = "Ⅰ. 주요 추진 과제",
        message: str | None = None,
        layout: Literal["cards", "process", "bullets", "table", "timeline", "kpi", "matrix", "roadmap", "gantt", "chevron"] = "cards",
        cards: list[dict[str, Any]] | None = None,
        bullets: list[str] | None = None,
        table: dict[str, Any] | None = None,
        notes: str | None = None,
        theme: Literal[
            "business_navy", "steel_blue", "modern_teal", "clean_charcoal",
            "emerald_green", "warm_beige", "indigo_violet"
        ] = "business_navy",
    ) -> dict[str, Any]:
        """
        [단계적 분할 생성 2단계 (브릿지)] 현재 열려 있는 활성 PowerPoint에 표준 비즈니스 서식 헤더와
        14대 카탈로그 슬라이스 레이아웃을 결합하여 슬라이드를 1장씩 실시간 추가합니다.

        - section: 대목차 인덱스 라벨 (예: "Ⅰ. 디지털혁신 전략과제")
        - title: 슬라이드 대제목
        - message: 핵심 요약 메시지
        - layout: "cards" | "process" | "bullets" | "table" | "timeline" | "kpi" | "matrix"
        - cards: [{"title": "...", "bullets": ["..."], "tag": "01"}]
        - bullets: ["불릿1", "불릿2"]
        - table: {"headers": ["구분", "내용"], "rows": [["항목A", "값A"]]}
        - theme: "business_navy"(표준 비즈니스 네이비 기본) 또는 범용 테마
        - notes: 하단 주석 ("※ 출처: ...")
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 연결 실패"}

        try:
            total_slides = prs.Slides.Count
            insert_idx = total_slides if total_slides >= 2 else total_slides + 1
            slide = prs.Slides.Add(insert_idx, 12)  # 12 = ppLayoutBlank
            try:
                ppt_app.ActiveWindow.View.GotoSlide(insert_idx)
            except Exception:
                pass

            resolved_theme = PptSliceEngine.resolve_theme(theme)
            slide_w = float(prs.PageSetup.SlideWidth)
            slide_h = float(prs.PageSetup.SlideHeight)

            # 1. 표준 비즈니스 헤더 렌더링
            h_bounds = SliceRegionCalculator.calculate_bounds(prs, region="header")
            h_data = {
                "section": section,
                "title": title,
                "message": message or "",
            }
            h_ids = PptSliceEngine.render_header(slide, h_bounds, h_data, resolved_theme)
            PptSliceEngine.group_shapes(slide, h_ids)

            # 2. 본문 바운딩 박스 계산 (SliceRegionCalculator)
            body_bounds = SliceRegionCalculator.calculate_bounds(prs, region="content")

            # 3. 레이아웃별 카탈로그 슬라이스 렌더링
            created_ids: list[int] = []
            if layout == "process":
                proc_data = {"steps": cards or []}
                created_ids = PptSliceEngine.render_process_flow(slide, body_bounds, proc_data, resolved_theme)
            elif layout == "table":
                tbl_data = table or {}
                created_ids = PptSliceEngine.render_smart_table(slide, body_bounds, tbl_data, resolved_theme)
            elif layout == "bullets":
                bul_data = {"title": "", "bullets": bullets or []}
                created_ids = PptSliceEngine.render_list_bullets(slide, body_bounds, bul_data, resolved_theme)
            elif layout == "timeline":
                tl_data = {"milestones": cards or []}
                created_ids = PptSliceEngine.render_timeline(slide, body_bounds, tl_data, resolved_theme)
            elif layout in ("kpi", "kpi_cards"):
                kpi_data = {"kpis": cards or []}
                created_ids = PptSliceEngine.render_kpi_cards(slide, body_bounds, kpi_data, resolved_theme)
            elif layout in ("matrix", "matrix_2x2"):
                mat_data = {"quadrants": cards or []}
                created_ids = PptSliceEngine.render_matrix_2x2(slide, body_bounds, mat_data, resolved_theme)
            else:  # "cards"
                # 데이터 정규화: items -> bullets
                norm_cards = []
                for c in (cards or []):
                    c_dict = dict(c)
                    if "items" in c_dict and "bullets" not in c_dict:
                        c_dict["bullets"] = c_dict["items"]
                    norm_cards.append(c_dict)
                card_data = {"cards": norm_cards}
                created_ids = PptSliceEngine.render_card_grid(slide, body_bounds, card_data, resolved_theme)

            if created_ids:
                PptSliceEngine.group_shapes(slide, created_ids)

            # 4. 하단 주석이 있을 경우
            if notes:
                margin_l = slide_w * 0.049
                margin_b = slide_h * 0.105
                nt_box = slide.Shapes.AddTextbox(1, margin_l, slide_h - margin_b - 16, slide_w * 0.892, 16)
                nt_box.TextFrame.MarginLeft = 0
                nt_box.TextFrame.MarginTop = 0
                nt_tr = nt_box.TextFrame.TextRange
                nt_tr.Text = f"※ {strip_markdown_syntax(notes)}"
                nt_tr.Font.Name = "맑은 고딕"
                nt_tr.Font.Size = 9
                nt_tr.Font.Color.RGB = resolved_theme.get("text_muted", hex_to_bgr("#64748B"))

            return {
                "success": True,
                "slide_index": insert_idx,
                "total_slides": prs.Slides.Count,
                "title": title,
                "theme_applied": theme,
                "layout_applied": layout,
                "progress_message": f"[{insert_idx}번째 슬라이드 추가 완료] '{theme}' 테마 카탈로그 규격으로 '{title}' 슬라이드가 즉시 반영되었습니다."
            }
        except Exception as e:
            return {"success": False, "error": f"슬라이드 실시간 추가 실패: {e}"}

    @mcp.tool()
    def ppt_update_shape_text(
        slide_index: int,
        shape_name_or_index: str | int,
        new_text: str,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        color_hex: str | None = None,
    ) -> dict[str, Any]:
        """
        현재 열려 있는 PowerPoint의 특정 슬라이드(`slide_index`) 내 도형/텍스트상자(`shape_name_or_index`)의
        텍스트와 서식을 실시간으로 수정합니다.
        
        `shape_name_or_index` 지원 형태:
        - 정수 또는 숫자 문자열: 1-based 도형 인덱스 (예: 1, 2)
        - 도형 명칭: PowerPoint Shape Name (예: "Title 1", "Card_0")
        - 시맨틱 키워드:
          - "title" 또는 "제목": 슬라이드 메인 타이틀 도형 자동 탐색
          - "subtitle" 또는 "소제목": 타이틀 하단 서브타이틀 도형 탐색
          - "badge", "category", "카테고리": 상단 카테고리 뱃지 탐색
          - "footer", "바닥글": 하단 푸터 텍스트 탐색
        - 텍스트 검색: "contains:키워드" (해당 키워드를 포함하는 도형 자동 탐색)
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        try:
            slide = prs.Slides(slide_index)
            sh = None

            # 1. 정수 인덱스 시도
            if isinstance(shape_name_or_index, int) or str(shape_name_or_index).isdigit():
                try:
                    sh = slide.Shapes(int(shape_name_or_index))
                except Exception:
                    sh = None

            # 2. 직접 도형 이름 일치 시도
            if sh is None:
                try:
                    sh = slide.Shapes(str(shape_name_or_index))
                except Exception:
                    sh = None

            # 3. 시맨틱 탐색 또는 텍스트 포함 탐색
            if sh is None:
                key = str(shape_name_or_index).strip().lower()
                text_candidates = []
                for s_item in slide.Shapes:
                    try:
                        if s_item.HasTextFrame and s_item.TextFrame.HasText:
                            text_candidates.append(s_item)
                    except Exception:
                        pass

                if key in ("title", "제목"):
                    best = None
                    max_size = 0.0
                    for c in text_candidates:
                        try:
                            if c.Top < 150:
                                fs = float(c.TextFrame.TextRange.Font.Size)
                                if fs > max_size:
                                    max_size = fs
                                    best = c
                        except Exception:
                            pass
                    sh = best or (text_candidates[0] if text_candidates else None)

                elif key in ("subtitle", "소제목"):
                    for c in text_candidates:
                        try:
                            if 70 <= c.Top <= 200 and float(c.TextFrame.TextRange.Font.Size) < 26:
                                sh = c
                                break
                        except Exception:
                            pass

                elif key in ("badge", "category", "카테고리"):
                    for c in text_candidates:
                        try:
                            if c.Top < 70 and c.Height < 35:
                                sh = c
                                break
                        except Exception:
                            pass

                elif key in ("footer", "바닥글"):
                    for c in text_candidates:
                        try:
                            if c.Top > 450:
                                sh = c
                                break
                        except Exception:
                            pass

                elif key.startswith("contains:"):
                    target_sub = key[9:].strip()
                    for c in text_candidates:
                        try:
                            if target_sub in c.TextFrame.TextRange.Text:
                                sh = c
                                break
                        except Exception:
                            pass

                elif not sh:
                    for c in text_candidates:
                        try:
                            if key in c.TextFrame.TextRange.Text.lower():
                                sh = c
                                break
                        except Exception:
                            pass

            if sh is None:
                return {
                    "success": False,
                    "error": f"슬라이드 {slide_index}에서 타겟 도형 '{shape_name_or_index}'을 찾을 수 없습니다."
                }

            clean_text = strip_markdown_syntax(new_text)
            tr = sh.TextFrame.TextRange
            tr.Text = clean_text.replace("\n", "\r")
            if font_name:
                tr.Font.Name = font_name
                try:
                    tr.Font.NameFarEast = font_name
                except Exception:
                    pass
            if font_size is not None:
                tr.Font.Size = float(font_size)
            if bold is not None:
                tr.Font.Bold = -1 if bold else 0
            if color_hex:
                tr.Font.Color.RGB = hex_to_bgr(color_hex)

            return {
                "success": True,
                "slide_index": slide_index,
                "shape_name": sh.Name,
                "updated_text": clean_text
            }
        except Exception as e:
            return {"success": False, "error": f"도형 텍스트 수정 실패: {e}"}

    @mcp.tool()
    def ppt_find_and_replace_text(
        target_text: str,
        replacement_text: str,
        slide_index: int | None = None,
        match_case: bool = False,
    ) -> dict[str, Any]:
        """
        열려 있는 PowerPoint 프레젠테이션의 전체 슬라이드(또는 지정 슬라이드) 내 모든 도형, 표, 그룹에서
        특정 문자열(`target_text`)을 찾아 `replacement_text`로 일괄 치환합니다.
        
        Args:
            target_text: 찾을 문자열 (예: "Project Titan", "1,250억")
            replacement_text: 바꿀 새 문자열 (예: "OmniCloud", "1,420억")
            slide_index: 특정 슬라이드만 치환할 경우 번호(1-based), None이면 전체 슬라이드 일괄 치환
            match_case: 대소문자 구분 여부 (기본값 False)
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        if not target_text:
            return {"success": False, "error": "찾을 문자열(target_text)이 비어 있습니다."}

        try:
            target_slides = []
            if slide_index is not None:
                if 1 <= slide_index <= prs.Slides.Count:
                    target_slides = [prs.Slides(slide_index)]
                else:
                    return {"success": False, "error": f"유효하지 않은 슬라이드 번호입니다: {slide_index}"}
            else:
                target_slides = [prs.Slides(i) for i in range(1, prs.Slides.Count + 1)]

            total_replaced = 0
            affected_slides = []

            def replace_in_shape(shape) -> int:
                count = 0
                try:
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        tr = shape.TextFrame.TextRange
                        curr_txt = tr.Text
                        found = False
                        if match_case:
                            if target_text in curr_txt:
                                found = True
                                new_txt = curr_txt.replace(target_text, replacement_text)
                        else:
                            pattern = re.compile(re.escape(target_text), re.IGNORECASE)
                            if pattern.search(curr_txt):
                                found = True
                                new_txt = pattern.sub(replacement_text, curr_txt)
                        if found:
                            tr.Text = new_txt.replace("\n", "\r")
                            count += 1
                except Exception:
                    pass

                try:
                    if shape.HasTable:
                        tbl = shape.Table
                        for r in range(1, tbl.Rows.Count + 1):
                            for c in range(1, tbl.Columns.Count + 1):
                                cell_tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
                                curr_txt = cell_tr.Text
                                found = False
                                if match_case:
                                    if target_text in curr_txt:
                                        found = True
                                        new_txt = curr_txt.replace(target_text, replacement_text)
                                else:
                                    pattern = re.compile(re.escape(target_text), re.IGNORECASE)
                                    if pattern.search(curr_txt):
                                        found = True
                                        new_txt = pattern.sub(replacement_text, curr_txt)
                                if found:
                                    cell_tr.Text = new_txt.replace("\n", "\r")
                                    count += 1
                except Exception:
                    pass

                try:
                    if shape.Type == 6:  # msoGroup
                        for sub_idx in range(1, shape.GroupItems.Count + 1):
                            count += replace_in_shape(shape.GroupItems(sub_idx))
                except Exception:
                    pass

                return count

            for sl in target_slides:
                slide_rep = 0
                for sh_item in sl.Shapes:
                    slide_rep += replace_in_shape(sh_item)
                if slide_rep > 0:
                    total_replaced += slide_rep
                    affected_slides.append(sl.SlideIndex)

            return {
                "success": True,
                "target_text": target_text,
                "replacement_text": replacement_text,
                "total_replaced_instances": total_replaced,
                "affected_slides": affected_slides,
                "message": f"총 {total_replaced}건의 텍스트를 성공적으로 치환했습니다. (슬라이드: {affected_slides})"
            }
        except Exception as e:
            return {"success": False, "error": f"텍스트 일괄 치환 실패: {e}"}


    @mcp.tool()
    def ppt_save_presentation(file_path: str | None = None) -> dict[str, Any]:
        """현재 활성 PowerPoint 프레젠테이션을 저장합니다."""
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}
        try:
            if file_path:
                abs_path = str(Path(file_path).resolve())
                Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
                prs.SaveAs(abs_path)
                return {"success": True, "path": abs_path}
            prs.Save()
            return {"success": True, "path": prs.FullName}
        except Exception as e:
            return {"success": False, "error": f"PowerPoint 저장 실패: {e}"}

    @mcp.tool()
    def ppt_list_slice_templates() -> dict[str, Any]:
        """
        [113개 템플릿 전수 REA 분석 기반 슬라이스 카탈로그]
        14대 슬라이스 역할(Role), 스키마 명세, 15대 세분화 영역 슬롯, 7대 실무 테마 목록을 반환합니다.
        """
        return {
            "success": True,
            "themes": CATALOG_DATA.get("themes", {}),
            "slice_roles": CATALOG_DATA.get("slice_roles", {}),
            "available_regions": [
                "full", "header", "content", "top_half", "bottom_half",
                "left_half", "right_half", "left_third", "right_two_thirds",
                "left_two_thirds", "right_third", "quadrant_1", "quadrant_2",
                "quadrant_3", "quadrant_4", "footer", "custom"
            ],
            "layout_presets": [
                "left_right_half (1:1 좌우 대칭 분할)",
                "narrative_split (1:2 좌측 33% 설명 + 우측 67% 시각화)",
                "sidebar_summary (2:1 좌측 67% 본문 + 우측 33% 요약)",
                "four_quadrants (2x2 4분면 그리드)",
                "top_bottom_half (상하 2분할)"
            ],
            "guide": (
                "슬라이스 1개를 슬라이드 특정 영역에 적재적소 주입하려면 'ppt_insert_slice'를 사용하고, "
                "1:1 또는 1:2 등 정형화된 분할 레이아웃으로 여러 슬라이스를 얹으려면 'ppt_split_slide_layout'을, "
                "상단 헤더 + 본문 등 전체 합성 슬라이드는 'ppt_compose_slide'를 사용하세요."
            ),
        }

    @mcp.tool()
    def ppt_insert_slice(
        role: Literal[
            "header", "card_grid", "process_flow", "comparison",
            "kpi_cards", "smart_table", "callout_box", "timeline",
            "matrix_2x2", "pyramid_hierarchy", "list_bullets",
            "agenda_toc", "quote_stat", "roadmap_quarterly"
        ],
        data: dict[str, Any],
        region: Literal[
            "full", "header", "content", "top_half", "bottom_half",
            "left_half", "right_half", "left_third", "right_two_thirds",
            "left_two_thirds", "right_third", "quadrant_1", "quadrant_2",
            "quadrant_3", "quadrant_4", "footer", "custom"
        ] = "content",
        custom_bounds: dict[str, float] | None = None,
        create_new_slide: bool = False,
        theme: Literal[
            "steel_blue", "business_navy", "modern_teal", "clean_charcoal",
            "emerald_green", "warm_beige", "indigo_violet"
        ] = "business_navy",
        do_group: bool = True,
    ) -> dict[str, Any]:
        """
        [적재적소 COM 전용 슬라이스 렌더러]
        외부 템플릿 파일 없이, 활성 PowerPoint 슬라이드의 지정된 영역(region)에
        14대 역할(role) 중 하나를 선택하여 고품질 비즈니스 슬라이스를 실시간 정밀 주입합니다.
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 연결 실패"}

        try:
            if create_new_slide or prs.Slides.Count == 0:
                insert_idx = prs.Slides.Count + 1
                slide = prs.Slides.Add(insert_idx, 12)  # ppLayoutBlank
                try:
                    ppt_app.ActiveWindow.View.GotoSlide(insert_idx)
                except Exception:
                    pass
            else:
                try:
                    curr_idx = ppt_app.ActiveWindow.View.Slide.SlideIndex
                    slide = prs.Slides(curr_idx)
                except Exception:
                    slide = prs.Slides(prs.Slides.Count)

            bounds = SliceRegionCalculator.calculate_bounds(prs, region=region, custom_bounds=custom_bounds)
            resolved_theme = PptSliceEngine.resolve_theme(theme)

            renderer = PptSliceEngine.get_renderer(role)
            if not renderer:
                return {"success": False, "error": f"지원하지 않는 슬라이스 역할: {role}"}

            created_ids = renderer(slide, bounds, data, resolved_theme)

            grouped = False
            if do_group and len(created_ids) >= 2:
                grp = PptSliceEngine.group_shapes(slide, created_ids)
                grouped = grp is not None

            return {
                "success": True,
                "slide_index": slide.SlideIndex,
                "role": role,
                "region": region,
                "bounds": {"x": bounds[0], "y": bounds[1], "w": bounds[2], "h": bounds[3]},
                "shapes_count": len(created_ids),
                "is_grouped": grouped,
            }
        except Exception as e:
            return {"success": False, "error": f"슬라이스 주입 실패: {e}"}


    @mcp.tool()
    def ppt_insert_recipe_slice(
        recipe_id: str,
        data: dict[str, Any] | None = None,
        create_new_slide: bool = False,
        custom_bounds: dict[str, float] | None = None,
    ) -> dict[str, Any]:
        """
        [113개 템플릿 99% 정밀 재현 슬라이스 렌더러]
        카탈로그에 등록된 113개 디자이너 템플릿의 정밀 레시피(recipe_id: recipe_001~recipe_113)를 선택하여
        활성 PowerPoint 슬라이드에 순수 COM으로 디자이너 원본과 99% 똑같이 실시간 복제 렌더링합니다.

        - recipe_id: "recipe_001" ~ "recipe_113" (ppt_list_slice_templates로 목록 조회 가능)
        - data: {"slot_0": "대제목", "slot_1": "소제목", "slot_2": "수치"} 형태로 원하는 텍스트만 실시간 치환 주입
        - create_new_slide: 새 슬라이드 생성 여부 (기본값: False, 현재 슬라이드에 주입)
        - custom_bounds: {"x": ..., "y": ..., "w": ..., "h": ...} 특정 영역에 축소/배치할 때 지정
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 연결 실패"}

        recipes = CATALOG_DATA.get("recipes", {})
        if recipe_id not in recipes:
            return {
                "success": False,
                "error": f"레시피 ID '{recipe_id}'를 찾을 수 없습니다. (허용 범위: recipe_001 ~ recipe_{len(recipes):03d})"
            }

        recipe = recipes[recipe_id]

        try:
            if create_new_slide or prs.Slides.Count == 0:
                insert_idx = prs.Slides.Count + 1
                slide = prs.Slides.Add(insert_idx, 12)  # ppLayoutBlank
                try:
                    ppt_app.ActiveWindow.View.GotoSlide(insert_idx)
                except Exception:
                    pass
            else:
                curr_idx = ppt_app.ActiveWindow.View.Slide.SlideIndex
                slide = prs.Slides(curr_idx)

            bounds_tuple = None
            if custom_bounds:
                bounds_tuple = (
                    float(custom_bounds.get("x", 0)),
                    float(custom_bounds.get("y", 0)),
                    float(custom_bounds.get("w", prs.PageSetup.SlideWidth)),
                    float(custom_bounds.get("h", prs.PageSetup.SlideHeight)),
                )

            created_ids = PptRecipeEngine.render_recipe(
                slide, recipe, data=data, custom_bounds=bounds_tuple
            )

            # 그룹화
            if len(created_ids) > 1:
                PptSliceEngine.group_shapes(slide, created_ids)

            return {
                "success": True,
                "recipe_id": recipe_id,
                "name": recipe.get("name"),
                "slide_index": slide.SlideIndex,
                "shapes_rendered": len(created_ids),
                "total_recipe_shapes": recipe.get("shapes_count"),
                "slots_provided": len(data) if data else 0,
                "progress_message": f"[{recipe_id}: '{recipe.get('name')}'] 슬라이스가 99% 디자이너 퀄리티로 화면에 즉시 렌더링되었습니다."
            }
        except Exception as e:
            return {"success": False, "error": f"레시피 슬라이스 렌더링 실패: {e}"}

    @mcp.tool()
    def ppt_split_slide_layout(
        layout_pattern: Literal[
            "left_right_half", "narrative_split", "sidebar_summary",
            "four_quadrants", "top_bottom_half"
        ],
        slots: list[dict[str, Any]],
        header_data: dict[str, Any] | None = None,
        create_new_slide: bool = True,
        theme: Literal[
            "steel_blue", "business_navy", "modern_teal", "clean_charcoal",
            "emerald_green", "warm_beige", "indigo_violet"
        ] = "business_navy",
    ) -> dict[str, Any]:
        """
        [적재적소 캔버스 분할 레이아웃 인터페이스]
        슬라이드를 정형화된 황금 분할 템플릿(1:1 대칭, 1:2 내러티브, 2:1 사이드바, 4분면 격자 등)으로 구조화하고,
        각 슬롯에 원하는 슬라이스를 한 번의 호출로 적재적소에 일괄 주입합니다.
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 연결 실패"}

        pattern_slot_regions = {
            "left_right_half": ["left_half", "right_half"],
            "narrative_split": ["left_third", "right_two_thirds"],
            "sidebar_summary": ["left_two_thirds", "right_third"],
            "four_quadrants": ["quadrant_2", "quadrant_1", "quadrant_3", "quadrant_4"],
            "top_bottom_half": ["top_half", "bottom_half"],
        }

        target_regions = pattern_slot_regions.get(layout_pattern, ["left_half", "right_half"])

        try:
            if create_new_slide or prs.Slides.Count == 0:
                insert_idx = prs.Slides.Count + 1
                slide = prs.Slides.Add(insert_idx, 12)  # ppLayoutBlank
                try:
                    ppt_app.ActiveWindow.View.GotoSlide(insert_idx)
                except Exception:
                    pass
            else:
                curr_idx = ppt_app.ActiveWindow.View.Slide.SlideIndex
                slide = prs.Slides(curr_idx)

            resolved_theme = PptSliceEngine.resolve_theme(theme)
            results = []

            # 1. Optional Header
            if header_data:
                h_bounds = SliceRegionCalculator.calculate_bounds(prs, region="header")
                h_ids = PptSliceEngine.render_header(slide, h_bounds, header_data, resolved_theme)
                PptSliceEngine.group_shapes(slide, h_ids)
                results.append({"slot": "header", "role": "header", "shapes_count": len(h_ids)})

            # 2. Injected Slots
            for idx, slot_item in enumerate(slots):
                if idx >= len(target_regions):
                    break
                reg = target_regions[idx]
                role = slot_item.get("role", "card_grid")
                data = slot_item.get("data") or {}

                bounds = SliceRegionCalculator.calculate_bounds(prs, region=reg)
                renderer = PptSliceEngine.get_renderer(role)
                if renderer:
                    created_ids = renderer(slide, bounds, data, resolved_theme)
                    grp = PptSliceEngine.group_shapes(slide, created_ids)
                    results.append({"slot_index": idx, "region": reg, "role": role, "shapes_count": len(created_ids), "grouped": grp is not None})

            return {
                "success": True,
                "slide_index": slide.SlideIndex,
                "layout_pattern": layout_pattern,
                "slots_rendered": len(results),
                "details": results,
            }
        except Exception as e:
            return {"success": False, "error": f"분할 레이아웃 생성 실패: {e}"}

    @mcp.tool()
    def ppt_compose_slide(
        slices: list[dict[str, Any]],
        frame_data: dict[str, Any] | None = None,
        create_new_slide: bool = True,
        theme: Literal[
            "steel_blue", "business_navy", "modern_teal", "clean_charcoal",
            "emerald_green", "warm_beige", "indigo_violet"
        ] = "business_navy",
    ) -> dict[str, Any]:
        """
        [완성형 복합 슬라이드 전체 합성 렌더러]
        슬라이드 1장에 여러 슬라이스(예: 상단 header + 중단 card_grid/kpi_cards + 하단 callout_box)를
        한 번의 호출로 조화롭게 결합하여 완성형 슬라이드를 실시간 생성합니다.
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 연결 실패"}

        try:
            if create_new_slide or prs.Slides.Count == 0:
                insert_idx = prs.Slides.Count + 1
                slide = prs.Slides.Add(insert_idx, 12)  # ppLayoutBlank
                try:
                    ppt_app.ActiveWindow.View.GotoSlide(insert_idx)
                except Exception:
                    pass
            else:
                curr_idx = ppt_app.ActiveWindow.View.Slide.SlideIndex
                slide = prs.Slides(curr_idx)

            slide_w = float(prs.PageSetup.SlideWidth)
            slide_h = float(prs.PageSetup.SlideHeight)
            resolved_theme = PptSliceEngine.resolve_theme(theme)

            # 0. 비즈니스 표준 서식 프레임(헤더·푸터) 옵션
            if frame_data:
                frame_type = frame_data.get("frame_type", "body")
                f_ids = PptSliceEngine.render_shell_frame(slide, frame_type=frame_type, data=frame_data, theme=resolved_theme, slide_w=slide_w, slide_h=slide_h)
                PptSliceEngine.group_shapes(slide, f_ids)

            results: list[dict[str, Any]] = []
            for sl in slices:
                role = sl.get("role", "card_grid")
                region = sl.get("region", "content")
                data = sl.get("data") or {}
                custom_bounds = sl.get("custom_bounds")

                bounds = SliceRegionCalculator.calculate_bounds(prs, region=region, custom_bounds=custom_bounds)
                resolved_theme = PptSliceEngine.resolve_theme(theme)

                renderer = PptSliceEngine.get_renderer(role)
                if renderer:
                    created_ids = renderer(slide, bounds, data, resolved_theme)
                    grp = PptSliceEngine.group_shapes(slide, created_ids)
                    results.append({"role": role, "region": region, "shapes_count": len(created_ids), "grouped": grp is not None})

            return {
                "success": True,
                "slide_index": slide.SlideIndex,
                "slices_composed": len(results),
                "details": results,
            }
        except Exception as e:
            return {"success": False, "error": f"슬라이드 합성 실패: {e}"}

    @mcp.tool()
    def ppt_add_shapes_batch(
        shapes: list[dict[str, Any]],
        slide_index: int | None = None,
        create_new_slide: bool = False,
    ) -> dict[str, Any]:
        """
        현재 열려 있는 PowerPoint 슬라이드에 다수의 도형(Shape)과 텍스트박스를 단일 호출로 일괄 생성합니다.
        PowerShell COM 방식과 동일하게 좌표(pt) 기반으로 정밀 배치할 수 있습니다.

        각 shape 딕셔너리 필드:
        - `type`: 도형 종류 ('rect', 'rounded_rect', 'oval', 'right_arrow', 'left_arrow', 'chevron', 'textbox' 또는 MSO 정수코드)
        - `left`, `top`, `width`, `height`: 좌표 및 크기 (pt 단위)
        - `fill_color`: 채우기 색상 ('#2E5CA8', '#FFFFFF', 'transparent' 등)
        - `line_color`: 테두리 색상 (생략 또는 'none' 시 테두리 없음)
        - `line_weight`: 테두리 두께 (pt, 기본 1.0)
        - `shadow`: 그림자 여부 (기본 False)
        - `text`: 내부 텍스트
        - `font_name`: 글꼴 이름 (기본 '맑은 고딕')
        - `font_size`: 글꼴 크기 (pt, 기본 11)
        - `font_color`: 글꼴 색상 (기본 '#262626')
        - `bold`: 굵게 여부 (기본 False)
        - `align`: 텍스트 가로 정렬 ('left', 'center', 'right', 기본 'left')
        - `vertical_align`: 세로 정렬 ('top', 'middle', 'bottom', 기본 'middle')
        - `margin`: 상하좌우 여백 (pt, 기본 0~4)
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 연결 실패"}

        try:
            slide, target_idx = _resolve_target_slide(ppt_app, prs, slide_index, create_new_slide)
            created_count = 0
            errors: list[str] = []

            with ppt_freeze_display(ppt_app):
                for i, shp_spec in enumerate(shapes):
                    try:
                        _create_single_shape(slide, shp_spec)
                        created_count += 1
                    except Exception as ex:
                        errors.append(f"도형 #{i+1} 생성 실패: {ex}")

            res: dict[str, Any] = {
                "success": True,
                "slide_index": target_idx,
                "created_count": created_count,
                "total_requested": len(shapes),
                "message": f"[{target_idx}번째 슬라이드] {created_count}개 도형이 일괄 생성되었습니다.",
            }
            if errors:
                res["partial_errors"] = errors[:5]
            return res
        except Exception as e:
            return {"success": False, "error": f"도형 일괄 생성 실패: {e}"}

    @mcp.tool()
    def ppt_insert_column_roadmap(
        title: str,
        columns: list[dict[str, Any]],
        subtitle: str | None = None,
        section: str | None = None,
        slide_index: int | None = None,
        create_new_slide: bool = True,
        show_arrows: bool = True,
        bottom_notes: str | None = None,
        theme: Literal[
            "business_navy", "steel_blue", "modern_teal", "clean_charcoal",
            "emerald_green", "warm_beige", "indigo_violet"
        ] = "business_navy",
    ) -> dict[str, Any]:
        """
        [최신 모던 컬럼 로드맵]
        - 113종 전수 분석 기반 황금 마진(좌우 41.6pt, content_top 146pt) 및 7대 테마 컬러 완벽 연동
        - 컬럼 간 화살표 MsoAutoShapeType 33(Right Arrow) 완벽 정렬
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 창을 열 수 없습니다."}

        try:
            slide, target_idx = _resolve_target_slide(ppt_app, prs, slide_index, create_new_slide)
            slide_w = float(getattr(prs.PageSetup, "SlideWidth", 960))
            slide_h = float(getattr(prs.PageSetup, "SlideHeight", 540))
            resolved_theme = PptSliceEngine.resolve_theme(theme)

            # 1. 외피 본문 프레임(Master Frame / Shell) 렌더링
            frame_data = {
                "section": section or "ROADMAP",
                "title": title,
                "category": section or "ROADMAP",
                "page_num": f"{target_idx:02d}",
                "company": "Enterprise Roadmap"
            }
            f_ids = PptSliceEngine.render_shell_frame(slide, "body", frame_data, resolved_theme, slide_w, slide_h)
            PptSliceEngine.group_shapes(slide, f_ids)

            # 2. 본문 가용 영역 계산 (황금 마진 자동 적응)
            b_x, b_y, b_w, b_h = SliceRegionCalculator.calculate_bounds(prs, region="content")

            n = max(1, len(columns))
            gap = 14.0 if show_arrows else 16.0
            arrow_w = 20.0 if show_arrows else 0.0
            total_gaps = (n - 1) * (gap + arrow_w)
            col_w = (b_w - total_gaps) / n
            col_h = b_h - 15.0

            created_ids: list[int] = []
            curr_x = b_x

            MSO_SHAPE_ROUNDED_RECTANGLE = 5
            MSO_SHAPE_RIGHT_ARROW = 33

            for i, col in enumerate(columns):
                c_title = str(col.get("title", f"단계 {i+1}")).strip()
                c_sub = str(col.get("subtitle") or col.get("assignee") or col.get("label") or "").strip()
                bullets = col.get("bullets") or []
                slogan = str(col.get("slogan", "")).strip()

                # 컬럼 엘리베이티드 카드
                card = PptSliceEngine._create_elevated_card(slide, curr_x, b_y, col_w, col_h, resolved_theme)
                created_ids.append(card.Id)

                # 상단 헤더 배지 (단계 번호 + 타이틀)
                badge = PptSliceEngine._create_pill_badge(slide, curr_x + 10, b_y + 12, f"0{i+1}. {c_title}", resolved_theme, bg_color=resolved_theme["primary"], text_color=resolved_theme["white"], w=col_w - 20, h=24.0)
                created_ids.append(badge.Id)

                # 서브타이틀 / 담당 (옵션)
                text_top = b_y + 42
                if c_sub:
                    sub_sh = slide.Shapes.AddTextbox(1, curr_x + 10, text_top, col_w - 20, 20)
                    sub_tr = sub_sh.TextFrame.TextRange
                    sub_tr.Text = c_sub
                    sub_tr.Font.Name = "맑은 고딕"
                    sub_tr.Font.Size = 10.0
                    sub_tr.Font.Color.RGB = resolved_theme["secondary"]
                    sub_tr.ParagraphFormat.Alignment = 2
                    created_ids.append(sub_sh.Id)
                    text_top += 22

                # 불릿 내용
                bullet_h = col_h - (text_top - b_y) - (36 if slogan else 14)
                if bullets:
                    b_sh = slide.Shapes.AddTextbox(1, curr_x + 10, text_top, col_w - 20, max(40, bullet_h))
                    b_tf = b_sh.TextFrame
                    b_tf.MarginLeft = 0
                    b_tf.MarginTop = 0
                    b_tf.WordWrap = -1
                    b_tr = b_tf.TextRange
                    b_tr.Text = "\n".join(f"• {b}" for b in bullets)
                    b_tr.Font.Name = "맑은 고딕"
                    b_tr.Font.Size = 10.0
                    b_tr.Font.Color.RGB = resolved_theme["text_dark"]
                    try:
                        b_tr.ParagraphFormat.SpaceWithin = 1.25
                    except Exception:
                        pass
                    b_tr.ParagraphFormat.SpaceAfter = 3.0
                    created_ids.append(b_sh.Id)

                # 하단 슬로건/핵심 결과물 배지
                if slogan:
                    sl_box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, curr_x + 8, b_y + col_h - 32, col_w - 16, 24)
                    sl_box.Fill.Solid()
                    sl_box.Fill.ForeColor.RGB = resolved_theme["bg_light"]
                    sl_box.Line.ForeColor.RGB = resolved_theme["card_border"]
                    sl_box.Line.Weight = 0.5
                    sl_tr = sl_box.TextFrame.TextRange
                    sl_tr.Text = f"★ {slogan}"
                    sl_tr.Font.Name = "맑은 고딕"
                    sl_tr.Font.Size = 9.5
                    sl_tr.Font.Bold = True
                    sl_tr.Font.Color.RGB = resolved_theme["primary"]
                    sl_tr.ParagraphFormat.Alignment = 2
                    created_ids.append(sl_box.Id)

                curr_x += col_w

                # 컬럼 간 오른쪽 화살표 커넥터 (msoShapeRightArrow = 33)
                if show_arrows and i < n - 1:
                    arr_x = curr_x + (gap / 2) - 4
                    arr_y = b_y + (col_h / 2) - 7
                    arr = slide.Shapes.AddShape(MSO_SHAPE_RIGHT_ARROW, arr_x, arr_y, arrow_w, 14)
                    arr.Fill.Solid()
                    arr.Fill.ForeColor.RGB = resolved_theme["secondary"]
                    arr.Line.Visible = 0
                    created_ids.append(arr.Id)
                    curr_x += gap + arrow_w

            if bottom_notes:
                n_sh = slide.Shapes.AddTextbox(1, b_x, slide_h - 30, b_w, 20)
                n_tr = n_sh.TextFrame.TextRange
                n_tr.Text = f"※ {bottom_notes}"
                n_tr.Font.Name = "맑은 고딕"
                n_tr.Font.Size = 8.5
                n_tr.Font.Color.RGB = resolved_theme["text_muted"]
                created_ids.append(n_sh.Id)

            return {
                "success": True,
                "slide_index": target_idx,
                "columns_count": n,
                "theme": theme,
                "shapes_created": len(created_ids),
                "message": f"[{theme} 모던 컬럼 로드맵] 슬라이드 {target_idx}번에 {n}단계 로드맵이 성공적으로 생성되었습니다."
            }
        except Exception as e:
            logger.error(f"Failed to insert column roadmap: {e}")
            return {"success": False, "error": str(e)}

    def ppt_delete_shapes(
        slide_index: int | None = None,
        shapes: list[str | int] | None = None,
        all: bool = False,
        exclude: list[str | int] | None = None,
        text_contains: str | None = None,
        fill_color_hex: str | None = None,
    ) -> dict[str, Any]:
        """
        지정한 슬라이드에서 조건에 부합하는 도형들을 일괄 삭제합니다.
        (ykuwai/ppt-mcp 벤치마킹 및 COM 인덱스 시프트 방지 역순 삭제 엔진)

        - slide_index: 슬라이드 번호 (생략 시 현재 활성 슬라이드)
        - shapes: 삭제할 도형의 이름(Name) 또는 1-based 인덱스 목록 (예: ["BlueBox 1", 3, 5])
        - all: True 시 슬라이드 내 모든 도형 대상 (exclude와 병행 사용 권장)
        - exclude: 삭제에서 제외할 도형 이름 또는 인덱스 목록 (예: ["Title 1"])
        - text_contains: 해당 텍스트를 포함하고 있는 도형만 삭제 (엄격 AND 필터)
        - fill_color_hex: 해당 채우기 색상(예: "#2E5CA8")을 가진 도형만 삭제 (엄격 AND 필터)
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        try:
            slide, target_idx = _resolve_target_slide(ppt_app, prs, slide_index, create_new_slide=False)
            total_shapes = slide.Shapes.Count
            if total_shapes == 0:
                return {"success": True, "slide_index": target_idx, "deleted_count": 0, "message": "슬라이드에 도형이 없습니다."}

            exclude_set = set(str(x).strip() for x in (exclude or []))
            target_names_set = set(str(x).strip() for x in shapes) if shapes is not None else None

            # BGR target color
            target_bgr = hex_to_bgr(fill_color_hex) if fill_color_hex else None

            # 1. 삭제 대상 판별: 역순 순회로 대상 도형 수집 (인덱스 시프트 누락 방지)
            to_delete_indices: list[int] = []
            deleted_names: list[str] = []

            for i in range(total_shapes, 0, -1):
                try:
                    shp = slide.Shapes(i)
                except Exception:
                    continue

                shp_name = str(getattr(shp, "Name", "")).strip()
                # Exclude 체크
                if shp_name in exclude_set or str(i) in exclude_set:
                    continue

                # Shapes 명시 목록 체크
                if target_names_set is not None:
                    if shp_name not in target_names_set and str(i) not in target_names_set:
                        continue
                elif not all and text_contains is None and fill_color_hex is None:
                    # 아무 필터도 안 줬고 all도 False면 오삭제 방지 차원에서 스킵
                    continue

                # AND 조건 1: text_contains
                if text_contains is not None:
                    sh_text = ""
                    try:
                        if shp.HasTextFrame and shp.TextFrame.HasText:
                            sh_text = shp.TextFrame.TextRange.Text or ""
                    except Exception:
                        pass
                    if text_contains not in sh_text:
                        continue

                # AND 조건 2: fill_color_hex
                if target_bgr is not None:
                    matched_color = False
                    try:
                        if shp.Fill.Visible and shp.Fill.ForeColor.RGB == target_bgr:
                            matched_color = True
                    except Exception:
                        pass
                    if not matched_color:
                        continue

                to_delete_indices.append(i)
                deleted_names.append(shp_name)

            if not to_delete_indices:
                return {
                    "success": True,
                    "slide_index": target_idx,
                    "deleted_count": 0,
                    "message": "조건에 부합하는 삭제 대상 도형이 없습니다.",
                }

            # 2. LockWindowUpdate 화면 동결 하에 역순 일괄 삭제 실행
            with ppt_freeze_display(ppt_app):
                for idx in to_delete_indices:
                    try:
                        slide.Shapes(idx).Delete()
                    except Exception as ex:
                        logger.warning(f"도형 #{idx} 삭제 중 오류: {ex}")

            return {
                "success": True,
                "slide_index": target_idx,
                "deleted_count": len(deleted_names),
                "deleted_shapes": deleted_names[:30],
                "message": f"[{target_idx}번째 슬라이드] {len(deleted_names)}개 도형이 안전하게 삭제되었습니다.",
            }
        except Exception as e:
            return {"success": False, "error": f"도형 일괄 삭제 실패: {e}"}

    @mcp.tool()
    def ppt_transform_shapes(
        slide_index: int | None = None,
        shapes: list[str | int] | None = None,
        all: bool = False,
        exclude: list[str | int] | None = None,
        transforms: list[dict[str, Any]] | None = None,
        delta_x: float = 0.0,
        dleft: float = 0.0,
        delta_y: float = 0.0,
        dtop: float = 0.0,
        delta_w: float = 0.0,
        dwidth: float = 0.0,
        delta_h: float = 0.0,
        dheight: float = 0.0,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        z_order: Literal["front", "back", "forward", "backward"] | None = None,
    ) -> dict[str, Any]:
        """
        기존 도형들의 위치(Left, Top) 및 크기(Width, Height)를 일괄 이동하거나 크기/레이어(Z-Order)를 변경합니다.
        (ykuwai/ppt-mcp 벤치마킹: dleft/dtop 오프셋 및 all/exclude 셀렉터 지원)

        - slide_index: 슬라이드 번호 (생략 시 현재 활성 슬라이드)
        - shapes: 대상 도형 이름 또는 인덱스 목록 (예: ["Card 1", "Card 2", 3])
        - all: True 시 슬라이드 내 모든 도형 대상 (exclude와 함께 사용하여 제목 제외 전체 이동 가능)
        - exclude: 이동/변형에서 제외할 도형 이름 또는 인덱스 목록
        - delta_x 또는 dleft: X축 상대 이동량 (pt, 예: +520 시 오른쪽으로 520pt 이동)
        - delta_y 또는 dtop: Y축 상대 이동량 (pt, 예: -30 시 위로 30pt 이동)
        - delta_w 또는 dwidth: 너비 상대 증감량 (pt)
        - delta_h 또는 dheight: 높이 상대 증감량 (pt)
        - left, top, width, height: 절대 좌표 및 크기 강제 지정
        - z_order: 레이어 겹침 순서 ('front', 'back', 'forward', 'backward')
        - transforms: 개별 도형마다 서로 다른 좌표를 적용할 경우의 딕셔너리 리스트
          [{"shape": "Box 1", "left": 100, "delta_y": 20}, ...]
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        try:
            slide, target_idx = _resolve_target_slide(ppt_app, prs, slide_index, create_new_slide=False)
            total_shapes = slide.Shapes.Count
            if total_shapes == 0:
                return {"success": True, "slide_index": target_idx, "transformed_count": 0, "message": "슬라이드에 도형이 없습니다."}

            eff_dx = delta_x if delta_x != 0.0 else dleft
            eff_dy = delta_y if delta_y != 0.0 else dtop
            eff_dw = delta_w if delta_w != 0.0 else dwidth
            eff_dh = delta_h if delta_h != 0.0 else dheight

            exclude_set = set(str(x).strip() for x in (exclude or []))
            transformed_names: list[str] = []

            with ppt_freeze_display(ppt_app):
                # 1. 개별 transforms 리스트가 전달된 경우
                if transforms:
                    for item in transforms:
                        sh_ref = item.get("shape") or item.get("name") or item.get("index")
                        if not sh_ref:
                            continue
                        try:
                            shp = slide.Shapes(int(sh_ref)) if str(sh_ref).isdigit() else slide.Shapes(str(sh_ref))
                        except Exception:
                            continue

                        s_name = str(getattr(shp, "Name", ""))
                        if s_name in exclude_set or str(sh_ref) in exclude_set:
                            continue

                        # 상대 오프셋
                        i_dx = item.get("delta_x", item.get("dleft", 0.0))
                        i_dy = item.get("delta_y", item.get("dtop", 0.0))
                        i_dw = item.get("delta_w", item.get("dwidth", 0.0))
                        i_dh = item.get("delta_h", item.get("dheight", 0.0))

                        if i_dx != 0.0:
                            shp.IncrementLeft(float(i_dx))
                        if i_dy != 0.0:
                            shp.IncrementTop(float(i_dy))
                        if i_dw != 0.0:
                            shp.Width = max(2.0, shp.Width + float(i_dw))
                        if i_dh != 0.0:
                            shp.Height = max(2.0, shp.Height + float(i_dh))

                        # 절대 좌표
                        if "left" in item and item["left"] is not None:
                            shp.Left = float(item["left"])
                        if "top" in item and item["top"] is not None:
                            shp.Top = float(item["top"])
                        if "width" in item and item["width"] is not None:
                            shp.Width = float(item["width"])
                        if "height" in item and item["height"] is not None:
                            shp.Height = float(item["height"])

                        # Z-Order
                        i_z = item.get("z_order")
                        if i_z and str(i_z).lower() in ZORDER_MAP:
                            shp.ZOrder(ZORDER_MAP[str(i_z).lower()])

                        transformed_names.append(s_name)

                # 2. 공통 일괄 변형 (shapes 또는 all)
                else:
                    target_shapes_list = []
                    if all:
                        for i in range(1, total_shapes + 1):
                            try:
                                shp = slide.Shapes(i)
                                s_name = str(getattr(shp, "Name", ""))
                                if s_name not in exclude_set and str(i) not in exclude_set:
                                    target_shapes_list.append(shp)
                            except Exception:
                                pass
                    elif shapes:
                        for ref in shapes:
                            try:
                                shp = slide.Shapes(int(ref)) if str(ref).isdigit() else slide.Shapes(str(ref))
                                s_name = str(getattr(shp, "Name", ""))
                                if s_name not in exclude_set and str(ref) not in exclude_set:
                                    target_shapes_list.append(shp)
                            except Exception:
                                pass

                    for shp in target_shapes_list:
                        s_name = str(getattr(shp, "Name", ""))
                        if eff_dx != 0.0:
                            shp.IncrementLeft(float(eff_dx))
                        if eff_dy != 0.0:
                            shp.IncrementTop(float(eff_dy))
                        if eff_dw != 0.0:
                            shp.Width = max(2.0, shp.Width + float(eff_dw))
                        if eff_dh != 0.0:
                            shp.Height = max(2.0, shp.Height + float(eff_dh))

                        if left is not None:
                            shp.Left = float(left)
                        if top is not None:
                            shp.Top = float(top)
                        if width is not None:
                            shp.Width = float(width)
                        if height is not None:
                            shp.Height = float(height)

                        if z_order and str(z_order).lower() in ZORDER_MAP:
                            shp.ZOrder(ZORDER_MAP[str(z_order).lower()])

                        transformed_names.append(s_name)

            return {
                "success": True,
                "slide_index": target_idx,
                "transformed_count": len(transformed_names),
                "transformed_shapes": transformed_names[:30],
                "message": f"[{target_idx}번째 슬라이드] {len(transformed_names)}개 도형의 변형/이동이 성공적으로 적용되었습니다.",
            }
        except Exception as e:
            return {"success": False, "error": f"도형 일괄 변형 실패: {e}"}

    @mcp.tool()
    def ppt_align_shapes(
        shapes: list[str | int],
        align: Literal["left", "center", "right", "top", "middle", "bottom"] | None = None,
        distribute: Literal["horizontally", "vertically"] | None = None,
        relative_to_slide: bool = False,
        slide_index: int | None = None,
    ) -> dict[str, Any]:
        """
        지정한 도형들을 일괄 정렬(Align)하거나 균등 분배(Distribute)합니다.
        (ykarapazar/pptx-mcp-live 벤치마킹)

        - shapes: 정렬할 대상 도형 이름 또는 인덱스 목록 (최소 2개 이상 필요)
        - align: 정렬 방식 ('left', 'center', 'right', 'top', 'middle', 'bottom')
        - distribute: 균등 분배 방식 ('horizontally', 'vertically')
        - relative_to_slide: True 시 슬라이드 전체 폭/높이 기준 균등 배치, False 시 선택 도형 영역 내 배치
        - slide_index: 슬라이드 번호 (생략 시 현재 활성 슬라이드)
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        if len(shapes) < 2:
            return {"success": False, "error": "정렬 및 균등 분배를 수행하려면 최소 2개 이상의 도형이 필요합니다."}

        try:
            slide, target_idx = _resolve_target_slide(ppt_app, prs, slide_index, create_new_slide=False)

            target_tuple = tuple(int(x) if str(x).isdigit() else str(x) for x in shapes)
            shape_range = slide.Shapes.Range(target_tuple)

            rel_mso = -1 if relative_to_slide else 0  # msoTrue = -1, msoFalse = 0

            with ppt_freeze_display(ppt_app):
                if align and str(align).lower() in ALIGN_MSO_MAP:
                    shape_range.Align(ALIGN_MSO_MAP[str(align).lower()], rel_mso)

                if distribute and str(distribute).lower() in DISTRIBUTE_MSO_MAP:
                    shape_range.Distribute(DISTRIBUTE_MSO_MAP[str(distribute).lower()], rel_mso)

            return {
                "success": True,
                "slide_index": target_idx,
                "aligned_count": len(shapes),
                "align": align,
                "distribute": distribute,
                "message": f"[{target_idx}번째 슬라이드] {len(shapes)}개 도형의 정렬/분배가 완료되었습니다.",
            }
        except Exception as e:
            return {"success": False, "error": f"도형 정렬 실패: {e}"}

    @mcp.tool()
    def ppt_auto_compose_from_text(
        raw_text: str,
        presentation_title: str | None = None,
        theme: Literal[
            "business_navy", "steel_blue", "modern_teal", "clean_charcoal",
            "emerald_green", "warm_beige", "indigo_violet"
        ] | None = None,
        output_path: str = "output_auto_composed_deck.pptx",
        open_in_desktop: bool = True,
    ) -> dict[str, Any]:
        """
        [원시 텍스트 덤프 기반 완결형 PPT 덱 자동 기획·구성 원샷 도구]
        사용자가 입력한 비구조화 줄글/기획서/회의록(Raw Text)을 스스로 분석하여,
        서사 구조(기승전결)와 17대 슬라이스 최적 아키타입(KPI, 프로세스, 비교, 로드맵, 매트릭스 등)을
        지능적으로 자동 판정하고 표지-목차-본문-엔딩으로 완결된 PowerPoint 프레젠테이션을 생성합니다.
        """
        from mcp_server.tools.composition_engine.auto_composer import AutoComposer

        if not raw_text or not raw_text.strip():
            return {"success": False, "error": "입력된 텍스트가 비어 있습니다."}

        # 1. 텍스트 분석 및 슬라이스 덱 스펙 자동 도출
        spec = AutoComposer.compose(
            raw_text=raw_text,
            presentation_title=presentation_title,
            theme=theme,
        )

        resolved_theme_name = spec.get("theme") or "business_navy"
        resolved_theme = PptSliceEngine.resolve_theme(resolved_theme_name)

        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=open_in_desktop)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "PowerPoint 창을 열 수 없습니다."}

        try:
            prs.PageSetup.SlideWidth = 960.0
            prs.PageSetup.SlideHeight = 540.0

            # 2. 표지 (Cover Shell) 생성
            cover_slide = prs.Slides.Add(prs.Slides.Count + 1, 12)
            PptSliceEngine.render_shell_frame(
                cover_slide,
                frame_type="cover",
                data={
                    "title": spec["title"],
                    "subtitle": spec["subtitle"],
                    "category": "AUTO COMPOSED DECK",
                    "author": "AI Transformation Office",
                    "date": "2026. 11",
                },
                theme=resolved_theme,
            )

            # 3. 목차 (Agenda Shell) 생성
            if spec.get("agenda_items"):
                agenda_slide = prs.Slides.Add(prs.Slides.Count + 1, 12)
                PptSliceEngine.render_shell_frame(
                    agenda_slide,
                    frame_type="agenda",
                    data={
                        "title": "AGENDA & OVERVIEW",
                        "subtitle": "핵심 아젠다 및 세부 추진 계획",
                        "agenda_items": spec["agenda_items"][:5],
                    },
                    theme=resolved_theme,
                )

            # 4. 본문 슬라이드군 자동 렌더링 (각 슬라이스 최적 매핑)
            for s_info in spec.get("slides", []):
                body_slide = prs.Slides.Add(prs.Slides.Count + 1, 12)
                # 외피 프레임 생성
                PptSliceEngine.render_shell_frame(
                    body_slide,
                    frame_type="body",
                    data={
                        "title": s_info["title"],
                        "subtitle": s_info.get("subtitle", ""),
                        "category": s_info.get("category", "STRATEGY"),
                        "author": "AI Transformation Office",
                    },
                    theme=resolved_theme,
                )
                # 내피 슬라이스 렌더링
                PptSliceEngine.render_slice(
                    body_slide,
                    role=s_info["role"],
                    data=s_info["data"],
                    theme=resolved_theme,
                    layout={"left": 48.0, "top": 140.0, "width": 864.0, "height": 330.0},
                )

            # 5. 엔딩 (Closing Shell) 생성
            closing_slide = prs.Slides.Add(prs.Slides.Count + 1, 12)
            PptSliceEngine.render_shell_frame(
                closing_slide,
                frame_type="closing",
                data={
                    "title": "Thank You",
                    "subtitle": "질의응답 및 향후 추진 협의",
                    "email": "contact@company.com",
                },
                theme=resolved_theme,
            )

            safe_out = com_manager.prepare_output_path(output_path)
            prs.SaveAs(str(safe_out))

            return {
                "success": True,
                "title": spec["title"],
                "theme": resolved_theme_name,
                "slides_count": prs.Slides.Count,
                "composed_slides": [
                    {"title": s["title"], "role": s["role"], "category": s["category"]}
                    for s in spec.get("slides", [])
                ],
                "path": str(safe_out),
                "message": f"총 {prs.Slides.Count}장의 완결형 프레젠테이션이 성공적으로 자동 기획·구성되었습니다.",
            }
        except Exception as e:
            logger.error(f"Failed to auto compose deck: {e}")
            return {"success": False, "error": f"자동 구성 덱 생성 실패: {e}"}

