# -*- coding: utf-8 -*-
"""
mcp_server/tools/ppt_slice_engine.py
113개 템플릿(1,495장 슬라이드, 17,446개 도형) 전수 REA 역공학 분석 기반
14대 COM 전용 HWP 스타일 적재적소 슬라이스 렌더링 엔진.

- 외부 템플릿 파일 없이 순수 COM 저수준 API를 통해 슬라이드 전체 또는 부분 영역에 슬라이스(컴포넌트) 정밀 주입
- 14대 풀스택 슬라이스:
  1. header: 대목차 + 28pt 대제목 + 16pt 핵심메시지 띠 + 구분선
  2. card_grid: 2~5열 반응형 둥근 사각형 카드 (상단 헤더 바 + 볼드 타이틀 + 불릿)
  3. process_flow: 3~5단계 화살표 연결 프로세스 (STEP 라벨 + 타이틀 + 설명 + ➜)
  4. comparison: 좌우 2단 대칭 비교 패널 (AS-IS 경고 레드 vs TO-BE 성공 그린)
  5. kpi_cards: 핵심 수치(28~32pt) 대시보드 카드 + 라벨 + 증감율(▲/▼) 배지
  6. smart_table: 겉바속촉 비즈니스 표 (상하 굵은선, 오픈형 측면, 헤더 음영, 배지)
  7. callout_box: 좌측 4pt 강조선 + 은은한 배경 음영 + 12pt 요약 박스
  8. timeline: 가로 로드맵 바 + 마일스톤 노드 원형 + 단계별 과제 카드
  9. matrix_2x2: 4분면 매트릭스 (X/Y 직교 축 + 4개 분면 독립 카드)
  10. pyramid_hierarchy: 3~4단 계층 구조 피라미드 (Top 비전 ➔ Middle 전략 ➔ Base 과제)
  11. list_bullets: KISA/HWP식 4단계 개조식 위계 리스트 (■ ➔ ○ ➔ - ➔ ※)
  12. agenda_toc: 목차 / 아젠다 슬라이스 (대형 번호 박스 + 목차 타이틀 + 요약 밴드)
  13. quote_stat: 대형 인용구 / 비전 슬라이스 (대형 따옴표 + 24pt 메시지 + 화자 배지)
  14. roadmap_quarterly: 4분기 매트릭스 로드맵 (Q1~Q4 4열 트랙별 마일스톤 카드)
- 전수 분석 기반 황금 마진율(Left 4.9%, Top 5.8%, Right 5.9%, Bottom 10.5%) 및 비례 간격 적용
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Literal

from ..com_manager import hex_to_bgr

logger = logging.getLogger("mcp_server.ppt_slice_engine")

# Load catalog
CATALOG_PATH = Path(__file__).parent / "ppt_slice_catalog.json"
CATALOG_DATA: dict[str, Any] = {}
if CATALOG_PATH.exists():
    try:
        with open(CATALOG_PATH, encoding="utf-8") as f:
            CATALOG_DATA = json.load(f)
    except Exception as e:
        logger.warning(f"Failed to load ppt_slice_catalog.json: {e}")

DEFAULT_THEMES = CATALOG_DATA.get("themes", {
    "steel_blue": {
        "primary": "#2E608A",
        "secondary": "#4A7FA7",
        "accent": "#0284C7",
        "bg_light": "#FFFFFF",
        "card_border": "#CBD5E1",
        "text_dark": "#1E293B",
        "text_muted": "#64748B",
    },
    "business_navy": {
        "primary": "#1F3864",
        "secondary": "#0055AA",
        "accent": "#F26621",
        "bg_light": "#F4F6F9",
        "card_border": "#CBD5E1",
        "text_dark": "#1E293B",
        "text_muted": "#64748B",
    },
    "modern_teal": {
        "primary": "#006D77",
        "secondary": "#00B7BF",
        "accent": "#E77547",
        "bg_light": "#F0FDF4",
        "card_border": "#80D2EC",
        "text_dark": "#0F172A",
        "text_muted": "#475569",
    },
    "clean_charcoal": {
        "primary": "#1E293B",
        "secondary": "#475569",
        "accent": "#2563EB",
        "bg_light": "#F8FAFC",
        "card_border": "#E2E8F0",
        "text_dark": "#0F172A",
        "text_muted": "#94A3B8",
    },
    "emerald_green": {
        "primary": "#16A643",
        "secondary": "#00B050",
        "accent": "#FFC000",
        "bg_light": "#F3F9FB",
        "card_border": "#A7F3D0",
        "text_dark": "#064E3B",
        "text_muted": "#047857",
    },
    "warm_beige": {
        "primary": "#78350F",
        "secondary": "#D97706",
        "accent": "#EE7700",
        "bg_light": "#FFFBEB",
        "card_border": "#FDE68A",
        "text_dark": "#451A03",
        "text_muted": "#92400E",
    },
    "indigo_violet": {
        "primary": "#3E3172",
        "secondary": "#9B88B2",
        "accent": "#E7252F",
        "bg_light": "#F3F1F9",
        "card_border": "#CCAEF0",
        "text_dark": "#1E1B4B",
        "text_muted": "#6B21A8",
    },
})

# MSO Shape types
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_RIGHT_ARROW = 33  # 13은 msoShapeCan(원통), 33이 msoShapeRightArrow(오른쪽 화살표)
MSO_SHAPE_OVAL = 9


class SliceRegionCalculator:
    """113개 템플릿 전수 분석에서 도출된 황금 마진율 기반 정밀 바운딩 박스(X, Y, W, H) 계산 엔진"""

    @staticmethod
    def calculate_bounds(
        pres: Any,
        region: Literal[
            "full", "header", "content", "top_half", "bottom_half",
            "left_half", "right_half", "custom"
        ] = "content",
        custom_bounds: dict[str, float] | None = None,
    ) -> tuple[float, float, float, float]:
        slide_w = float(pres.PageSetup.SlideWidth)
        slide_h = float(pres.PageSetup.SlideHeight)

        # REA golden median margins: left 4.9%, right 5.9%, top 5.8%, bottom 10.5%
        margin_l = slide_w * 0.049
        margin_r = slide_w * 0.059
        margin_t = slide_h * 0.058
        margin_b = slide_h * 0.105

        avail_w = slide_w - margin_l - margin_r
        avail_h = slide_h - margin_t - margin_b

        header_h = 75.0
        content_top = margin_t + header_h + 30.0  # 황금 간격 30pt 확보 (146.1pt)
        content_h = slide_h - content_top - margin_b

        if region == "custom" and custom_bounds:
            x = float(custom_bounds.get("x", margin_l))
            y = float(custom_bounds.get("y", content_top))
            w = float(custom_bounds.get("w", avail_w))
            h = float(custom_bounds.get("h", content_h))
            return x, y, w, h

        if region == "header":
            return margin_l, margin_t, avail_w, header_h

        if region == "full":
            return margin_l, margin_t, avail_w, avail_h

        if region == "content":
            return margin_l, content_top, avail_w, content_h

        if region == "top_half":
            half_h = (content_h - 15.0) / 2
            return margin_l, content_top, avail_w, half_h

        if region == "bottom_half":
            half_h = (content_h - 15.0) / 2
            return margin_l, content_top + half_h + 15.0, avail_w, half_h

        if region == "left_half":
            half_w = (avail_w - 20.0) / 2
            return margin_l, content_top, half_w, content_h

        if region == "right_half":
            half_w = (avail_w - 20.0) / 2
            return margin_l + half_w + 20.0, content_top, half_w, content_h

        if region == "left_third":
            w_third = (avail_w - 20.0) * 0.33
            return margin_l, content_top, w_third, content_h

        if region == "right_two_thirds":
            w_third = (avail_w - 20.0) * 0.33
            w_two_thirds = (avail_w - 20.0) * 0.67
            return margin_l + w_third + 20.0, content_top, w_two_thirds, content_h

        if region == "left_two_thirds":
            w_two_thirds = (avail_w - 20.0) * 0.67
            return margin_l, content_top, w_two_thirds, content_h

        if region == "right_third":
            w_two_thirds = (avail_w - 20.0) * 0.67
            w_third = (avail_w - 20.0) * 0.33
            return margin_l + w_two_thirds + 20.0, content_top, w_third, content_h

        if region in ("quadrant_1", "quadrant_2", "quadrant_3", "quadrant_4"):
            qw = (avail_w - 16.0) / 2
            qh = (content_h - 16.0) / 2
            if region == "quadrant_2":  # top-left
                return margin_l, content_top, qw, qh
            elif region == "quadrant_1":  # top-right
                return margin_l + qw + 16.0, content_top, qw, qh
            elif region == "quadrant_3":  # bottom-left
                return margin_l, content_top + qh + 16.0, qw, qh
            elif region == "quadrant_4":  # bottom-right
                return margin_l + qw + 16.0, content_top + qh + 16.0, qw, qh

        if region == "footer":
            return margin_l, slide_h - margin_b - 24.0, avail_w, 24.0

        return margin_l, content_top, avail_w, content_h


class PptSliceEngine:
    """순수 COM 기반 14대 저수준 슬라이스 렌더러"""

    @staticmethod
    def resolve_theme(theme_spec: str | dict[str, str] | None) -> dict[str, int]:
        base = DEFAULT_THEMES.get("business_navy", {})
        if isinstance(theme_spec, str) and theme_spec in DEFAULT_THEMES:
            base = DEFAULT_THEMES[theme_spec]
        elif isinstance(theme_spec, dict):
            base = {**base, **theme_spec}

        return {
            "primary": hex_to_bgr(base.get("primary", "#1F3864")),
            "secondary": hex_to_bgr(base.get("secondary", "#0055AA")),
            "accent": hex_to_bgr(base.get("accent", "#F26621")),
            "bg_light": hex_to_bgr(base.get("bg_light", "#F4F6F9")),
            "card_border": hex_to_bgr(base.get("card_border", "#CBD5E1")),
            "text_dark": hex_to_bgr(base.get("text_dark", "#1E293B")),
            "text_muted": hex_to_bgr(base.get("text_muted", "#64748B")),
            "white": hex_to_bgr("#FFFFFF"),
        }

    @classmethod
    def get_renderer(cls, role: str):
        render_map = {
            "header": cls.render_header,
            "card_grid": cls.render_card_grid,
            "card_grid_3col": cls.render_card_grid,
            "process_flow": cls.render_process_flow,
            "flow_process_steps": cls.render_process_flow,
            "comparison": cls.render_comparison,
            "split_2col": cls.render_comparison,
            "kpi_cards": cls.render_kpi_cards,
            "metric_cards": cls.render_kpi_cards,
            "smart_table": cls.render_smart_table,
            "comparison_table": cls.render_smart_table,
            "callout_box": cls.render_callout_box,
            "timeline": cls.render_timeline,
            "timeline_roadmap": cls.render_timeline,
            "matrix_2x2": cls.render_matrix_2x2,
            "pyramid_hierarchy": cls.render_pyramid_hierarchy,
            "list_bullets": cls.render_list_bullets,
            "agenda_toc": cls.render_agenda_toc,
            "quote_stat": cls.render_quote_stat,
            "roadmap_quarterly": cls.render_roadmap_quarterly,
            "roadmap_quarterly_v2": cls.render_roadmap_quarterly_v2,
            "roadmap_gantt": cls.render_roadmap_gantt,
            "roadmap_chevron": cls.render_roadmap_chevron,
        }
        return render_map.get(role)

    @classmethod
    def render_slice(
        cls,
        slide: Any,
        role: str,
        data: dict[str, Any],
        theme: dict[str, int] | None = None,
        layout: dict[str, float] | tuple[float, float, float, float] | None = None,
    ) -> list[int]:
        """
        통합 슬라이스 렌더러 호출 메서드.
        카탈로그 역할명 및 엔진 내부 역할명을 모두 수용하며 데이터 스키마를 정규화하여 렌더링합니다.
        """
        renderer = cls.get_renderer(role)
        if not renderer:
            raise ValueError(f"Unknown slice role: {role}")
        resolved_theme = theme or cls.resolve_theme("business_navy")
        if isinstance(layout, dict):
            bounds = (
                float(layout.get("left", 48.0)),
                float(layout.get("top", 140.0)),
                float(layout.get("width", 864.0)),
                float(layout.get("height", 330.0)),
            )
        elif isinstance(layout, (tuple, list)):
            bounds = tuple(layout)
        else:
            bounds = (48.0, 140.0, 864.0, 330.0)

        # 데이터 키 호환성 보정
        normalized_data = dict(data)
        if role in ("comparison", "split_2col") and "left_title" in data:
            normalized_data = {
                "left_panel": {"title": data.get("left_title", ""), "items": data.get("left_items", [])},
                "right_panel": {"title": data.get("right_title", ""), "items": data.get("right_items", [])},
            }
        elif role in ("kpi_cards", "metric_cards") and "metrics" in data:
            normalized_data["kpis"] = data["metrics"]

        return renderer(slide, bounds, normalized_data, resolved_theme)



    @classmethod
    def render_shell_frame(
        cls,
        slide: Any,
        frame_type: Literal["cover", "agenda", "divider", "body", "closing"] = "body",
        data: dict[str, Any] | None = None,
        theme: dict[str, int] | None = None,
        slide_w: float = 960.0,
        slide_h: float = 540.0,
    ) -> list[int]:
        """
        113종 전수 분석 기반 5대 슬라이드 외피(Master Shell Frames) 렌더러
        - cover: 프레젠테이션 시작 표지 (대제목, 소제목, 작성자/부서, 일자)
        - agenda: 목차 헤더 프레임
        - divider: 간지/장구분 전환 프레임 (대형 챕터 넘버, 장 제목, 키 메시지)
        - body: 표준 본문 프레임 (카테고리 뱃지/소제목, 슬라이드 대제목, 구분선, 푸터/쪽번호)
        - closing: 엔딩/Q&A 프레임 (Thank You, 연락처 안내)
        """
        if data is None:
            data = {}
        if theme is None:
            theme = cls.resolve_theme("business_navy")

        created: list[int] = []
        margin_l = slide_w * 0.043  # ~41.6pt
        margin_r = slide_w * 0.043
        margin_t = slide_h * 0.076  # ~41.1pt
        avail_w = slide_w - margin_l - margin_r

        if frame_type == "cover":
            title = str(data.get("title", "프레젠테이션 제목")).strip()
            subtitle = str(data.get("subtitle", "부제목 또는 핵심 슬로건")).strip()
            author = str(data.get("author", "")).strip()
            date = str(data.get("date", "")).strip()

            category = str(data.get("category", "")).strip()
            if category:
                badge_w = min(len(category) * 11 + 24, 180.0)
                badge = cls._create_pill_badge(slide, margin_l, margin_t + 40, category, theme, bg_color=theme["primary"], text_color=theme["white"], w=badge_w, h=24.0)
                created.append(badge.Id)

            t_box = slide.Shapes.AddTextbox(1, margin_l, margin_t + 80, avail_w, 80)
            t_box.TextFrame.MarginLeft = 0
            t_box.TextFrame.MarginTop = 0
            tr = t_box.TextFrame.TextRange
            tr.Text = title
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 34
            tr.Font.Bold = True
            tr.Font.Color.RGB = theme["primary"]
            created.append(t_box.Id)

            if subtitle:
                st_box = slide.Shapes.AddTextbox(1, margin_l, margin_t + 175, avail_w, 40)
                st_box.TextFrame.MarginLeft = 0
                st_box.TextFrame.MarginTop = 0
                str_range = st_box.TextFrame.TextRange
                str_range.Text = subtitle
                str_range.Font.Name = "맑은 고딕"
                str_range.Font.Size = 18
                str_range.Font.Color.RGB = theme["secondary"]
                created.append(st_box.Id)

            line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, margin_l, margin_t + 235, 80, 4)
            line.Fill.Solid()
            line.Fill.ForeColor.RGB = theme["accent"]
            line.Line.Visible = False
            created.append(line.Id)

            meta_text = " | ".join(filter(None, [author, date]))
            if meta_text:
                m_box = slide.Shapes.AddTextbox(1, margin_l, slide_h - 70, avail_w, 30)
                m_box.TextFrame.MarginLeft = 0
                m_box.TextFrame.MarginTop = 0
                mtr = m_box.TextFrame.TextRange
                mtr.Text = meta_text
                mtr.Font.Name = "맑은 고딕"
                mtr.Font.Size = 12
                mtr.Font.Color.RGB = theme["text_muted"]
                created.append(m_box.Id)

        elif frame_type == "agenda":
            title = str(data.get("title", "AGENDA")).strip()
            subtitle = str(data.get("subtitle", "목차 및 주요 진행 순서")).strip()

            badge = cls._create_pill_badge(slide, margin_l, margin_t, "CONTENTS", theme, bg_color=theme["secondary"], text_color=theme["white"], w=100.0, h=22.0)
            created.append(badge.Id)

            t_box = slide.Shapes.AddTextbox(1, margin_l, margin_t + 28, avail_w, 45)
            t_box.TextFrame.MarginLeft = 0
            t_box.TextFrame.MarginTop = 0
            tr = t_box.TextFrame.TextRange
            tr.Text = title
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 28
            tr.Font.Bold = True
            tr.Font.Color.RGB = theme["primary"]
            created.append(t_box.Id)

            line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, margin_l, margin_t + 78, avail_w, 1.5)
            line.Fill.Solid()
            line.Fill.ForeColor.RGB = theme["card_border"]
            line.Line.Visible = False
            created.append(line.Id)

        elif frame_type == "divider":
            part = str(data.get("part", "PART 01")).strip()
            title = str(data.get("title", "챕터 제목")).strip()
            message = str(data.get("message", "")).strip()

            badge_w = min(len(part) * 11 + 24, 140.0)
            badge = cls._create_pill_badge(slide, margin_l, slide_h * 0.32, part, theme, bg_color=theme["primary"], text_color=theme["white"], w=badge_w, h=26.0)
            created.append(badge.Id)

            t_box = slide.Shapes.AddTextbox(1, margin_l, slide_h * 0.40, avail_w, 60)
            t_box.TextFrame.MarginLeft = 0
            t_box.TextFrame.MarginTop = 0
            tr = t_box.TextFrame.TextRange
            tr.Text = title
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 32
            tr.Font.Bold = True
            tr.Font.Color.RGB = theme["primary"]
            created.append(t_box.Id)

            line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, margin_l, slide_h * 0.40 + 60, 100, 4)
            line.Fill.Solid()
            line.Fill.ForeColor.RGB = theme["accent"]
            line.Line.Visible = False
            created.append(line.Id)

            if message:
                msg_box = slide.Shapes.AddTextbox(1, margin_l, slide_h * 0.40 + 75, avail_w, 40)
                msg_box.TextFrame.MarginLeft = 0
                msg_box.TextFrame.MarginTop = 0
                mtr = msg_box.TextFrame.TextRange
                mtr.Text = message
                mtr.Font.Name = "맑은 고딕"
                mtr.Font.Size = 14
                mtr.Font.Color.RGB = theme["text_muted"]
                created.append(msg_box.Id)

        elif frame_type == "closing":
            headline = str(data.get("headline", "감사합니다")).strip()
            sub = str(data.get("subtitle", "Q & A")).strip()
            contact = str(data.get("contact", "")).strip()

            t_box = slide.Shapes.AddTextbox(1, margin_l, slide_h * 0.33, avail_w, 60)
            t_box.TextFrame.MarginLeft = 0
            t_box.TextFrame.MarginTop = 0
            tr = t_box.TextFrame.TextRange
            tr.Text = headline
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 38
            tr.Font.Bold = True
            tr.Font.Color.RGB = theme["primary"]
            created.append(t_box.Id)

            if sub:
                s_box = slide.Shapes.AddTextbox(1, margin_l, slide_h * 0.33 + 65, avail_w, 35)
                s_box.TextFrame.MarginLeft = 0
                s_box.TextFrame.MarginTop = 0
                str_range = s_box.TextFrame.TextRange
                str_range.Text = sub
                str_range.Font.Name = "맑은 고딕"
                str_range.Font.Size = 20
                str_range.Font.Color.RGB = theme["secondary"]
                created.append(s_box.Id)

            if contact:
                c_box = slide.Shapes.AddTextbox(1, margin_l, slide_h * 0.33 + 115, avail_w, 40)
                c_box.TextFrame.MarginLeft = 0
                c_box.TextFrame.MarginTop = 0
                ctr = c_box.TextFrame.TextRange
                ctr.Text = contact
                ctr.Font.Name = "맑은 고딕"
                ctr.Font.Size = 13
                ctr.Font.Color.RGB = theme["text_muted"]
                created.append(c_box.Id)

        else:  # "body"
            section = str(data.get("section", "")).strip()
            title = str(data.get("title", "슬라이드 제목을 입력하세요")).strip()
            category = str(data.get("category", "")).strip()
            page_num = str(data.get("page_num", "")).strip()
            company = str(data.get("company", "")).strip()

            badge_label = category or section
            if badge_label:
                badge_w = min(len(badge_label) * 11 + 24, 180.0)
                badge = cls._create_pill_badge(slide, margin_l, margin_t, badge_label, theme, bg_color=theme["secondary"], text_color=theme["white"], w=badge_w, h=22.0)
                created.append(badge.Id)
                title_top = margin_t + 28
            else:
                title_top = margin_t

            t_box = slide.Shapes.AddTextbox(1, margin_l, title_top, avail_w, 42)
            t_box.TextFrame.MarginLeft = 0
            t_box.TextFrame.MarginTop = 0
            tr = t_box.TextFrame.TextRange
            tr.Text = title
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 24
            tr.Font.Bold = True
            tr.Font.Color.RGB = theme["primary"]
            created.append(t_box.Id)

            show_line = data.get("show_divider_line", True)
            if show_line:
                line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, margin_l, title_top + 45, avail_w, 1.5)
                line.Fill.Solid()
                line.Fill.ForeColor.RGB = theme.get("card_border", hex_to_bgr("#E2E8F0"))
                line.Line.Visible = False
                created.append(line.Id)

            footer_y = slide_h - 35
            if company:
                c_sh = slide.Shapes.AddTextbox(1, margin_l, footer_y, 300, 20)
                c_sh.TextFrame.MarginLeft = 0
                c_sh.TextFrame.MarginTop = 0
                ctr = c_sh.TextFrame.TextRange
                ctr.Text = company
                ctr.Font.Name = "맑은 고딕"
                ctr.Font.Size = 9
                ctr.Font.Color.RGB = theme["text_muted"]
                created.append(c_sh.Id)

            if page_num:
                p_sh = slide.Shapes.AddTextbox(1, slide_w - margin_r - 80, footer_y, 80, 20)
                p_sh.TextFrame.MarginLeft = 0
                p_sh.TextFrame.MarginTop = 0
                ptr = p_sh.TextFrame.TextRange
                ptr.Text = page_num
                ptr.Font.Name = "맑은 고딕"
                ptr.Font.Size = 9
                ptr.Font.Color.RGB = theme["text_muted"]
                created.append(p_sh.Id)

        return created


    # =========================================================================
    # REA Aesthetic Helper Methods
    # =========================================================================
    @staticmethod
    def _create_elevated_card(slide: Any, x: float, y: float, w: float, h: float, theme: dict[str, int], bg_color: int | None = None, border_color: int | None = None) -> Any:
        """REA 실측: 모서리 0.0622 곡률, 0.5pt 헤어라인, 부드러운 5pt 드롭 섀도우 적용 카드"""
        card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, w, h)
        try:
            if hasattr(card, "Adjustments") and card.Adjustments.Count > 0:
                card.Adjustments.Item(1, 0.0622) # REA 황금 곡률 (6.2%)
        except Exception:
            pass

        card.Fill.Solid()
        card.Fill.ForeColor.RGB = bg_color if bg_color is not None else theme["white"]

        card.Line.Visible = -1
        card.Line.Weight = 0.50 # REA 헤어라인 보더
        card.Line.ForeColor.RGB = border_color if border_color is not None else theme["card_border"]

        # 드롭 섀도우
        try:
            if hasattr(card, "Shadow"):
                card.Shadow.Visible = -1
                card.Shadow.Blur = 5.0
                card.Shadow.Transparency = 0.60
                card.Shadow.OffsetX = 2.0
                card.Shadow.OffsetY = 2.0
        except Exception:
            pass

        return card

    @staticmethod
    def _create_pill_badge(slide: Any, x: float, y: float, text: str, theme: dict[str, int], bg_color: int | None = None, text_color: int | None = None, w: float = 48.0, h: float = 20.0) -> Any:
        """좌상단 앙증맞은 시각적 앵커 마이크로 필(Pill) 태그"""
        badge = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, w, h)
        try:
            if hasattr(badge, "Adjustments") and badge.Adjustments.Count > 0:
                badge.Adjustments.Item(1, 0.5) # 완전 둥근 알약형
        except Exception:
            pass

        badge.Fill.Solid()
        badge.Fill.ForeColor.RGB = bg_color if bg_color is not None else theme["bg_light"]
        badge.Line.Visible = -1
        badge.Line.Weight = 0.50
        badge.Line.ForeColor.RGB = theme["card_border"]

        tf = badge.TextFrame
        tf.MarginLeft = 4
        tf.MarginRight = 4
        tf.MarginTop = 1
        tf.MarginBottom = 1
        tr = tf.TextRange
        tr.Text = text
        tr.Font.Name = "맑은 고딕"
        tr.Font.Size = 9.5
        tr.Font.Bold = True
        tr.Font.Color.RGB = text_color if text_color is not None else theme["primary"]
        tr.ParagraphFormat.Alignment = 2 # ppAlignCenter
        return badge

    # =========================================================================
    # 1. Header Slice
    # =========================================================================
    @staticmethod
    def render_header(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        created: list[int] = []

        section = str(data.get("section", "Ⅰ. 개요 및 배경")).strip()
        title = str(data.get("title", "슬라이드 제목을 입력하세요")).strip()
        message = str(data.get("message", "")).strip()

        # 1. Section label
        sh1 = slide.Shapes.AddTextbox(1, x, y, w, 20)
        sh1.TextFrame.MarginLeft = 0
        sh1.TextFrame.MarginTop = 0
        tr1 = sh1.TextFrame.TextRange
        tr1.Text = section
        tr1.Font.Name = "맑은 고딕"
        tr1.Font.Size = 11
        tr1.Font.Bold = True
        tr1.Font.Color.RGB = theme["secondary"]
        created.append(sh1.Id)

        # 2. Main title (REA golden typical size 28pt)
        sh2 = slide.Shapes.AddTextbox(1, x, y + 22, w, 40)
        sh2.TextFrame.MarginLeft = 0
        sh2.TextFrame.MarginTop = 0
        tr2 = sh2.TextFrame.TextRange
        tr2.Text = title
        tr2.Font.Name = "맑은 고딕"
        tr2.Font.Size = 24
        tr2.Font.Bold = True
        tr2.Font.Color.RGB = theme["primary"]
        created.append(sh2.Id)

        # 3. Message ribbon (REA golden typical size 16pt / 12pt)
        curr_y = y + 64
        if message:
            ribbon = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, curr_y, w, 34)
            ribbon.Fill.Solid()
            ribbon.Fill.ForeColor.RGB = theme["bg_light"]
            ribbon.Line.ForeColor.RGB = theme["secondary"]
            ribbon.Line.Weight = 1.0

            tf_r = ribbon.TextFrame
            tf_r.MarginLeft = 14
            tf_r.MarginRight = 14
            tf_r.MarginTop = 6
            tf_r.MarginBottom = 6
            tr_r = tf_r.TextRange
            tr_r.Text = f"💡  {message}"
            tr_r.Font.Name = "맑은 고딕"
            tr_r.Font.Size = 12
            tr_r.Font.Bold = True
            tr_r.Font.Color.RGB = theme["primary"]
            created.append(ribbon.Id)
        else:
            line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, x, curr_y, w, 2)
            line.Fill.Solid()
            line.Fill.ForeColor.RGB = theme["secondary"]
            line.Line.Visible = 0
            created.append(line.Id)

        return created

    # =========================================================================
    # 2. Card Grid Slice (REA Aesthetic Upgrade)
    # =========================================================================
    @classmethod
    def render_card_grid(cls, slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        cards = data.get("cards") or []
        n = max(1, len(cards))
        created: list[int] = []

        spacing = 27.3 # REA 황금 간격
        card_w = (w - (spacing * (n - 1))) / n
        card_h = h

        for i, c in enumerate(cards):
            card_x = x + i * (card_w + spacing)
            tag = str(c.get("tag", f"0{i+1}")).strip()
            title = str(c.get("title", f"과제 {i+1}")).strip()
            bullets = c.get("bullets") or []
            if not bullets and c.get("desc"):
                desc_val = str(c["desc"]).strip()
                bullets = [l.strip() for l in desc_val.splitlines() if l.strip()]

            # 1. 소프트 엘리베이티드 카드
            box = cls._create_elevated_card(slide, card_x, y, card_w, card_h, theme)
            created.append(box.Id)

            # 2. 마이크로 필 태그 (좌상단 앵커)
            badge = cls._create_pill_badge(slide, card_x + 14, y + 14, tag, theme, w=max(46.0, len(tag)*10.0))
            created.append(badge.Id)

            # 3. 카드 볼드 타이틀 (태그 우측 또는 하단)
            t_sh = slide.Shapes.AddTextbox(1, card_x + 14, y + 40, card_w - 28, 28)
            t_sh.TextFrame.MarginLeft = 0
            t_sh.TextFrame.MarginTop = 0
            t_tr = t_sh.TextFrame.TextRange
            t_tr.Text = title
            t_tr.Font.Name = "맑은 고딕"
            t_tr.Font.Size = 13.5
            t_tr.Font.Bold = True
            t_tr.Font.Color.RGB = theme["text_dark"]
            created.append(t_sh.Id)

            # 4. 얇은 액센트 구분선
            div_line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, card_x + 14, y + 72, card_w - 28, 1)
            div_line.Fill.Solid()
            div_line.Fill.ForeColor.RGB = theme["bg_light"]
            div_line.Line.Visible = 0
            created.append(div_line.Id)

            # 5. 불릿 리스트 (여백과 줄간격 확보)
            if bullets:
                b_sh = slide.Shapes.AddTextbox(1, card_x + 14, y + 80, card_w - 28, card_h - 94)
                tf_b = b_sh.TextFrame
                tf_b.WordWrap = -1
                tf_b.MarginLeft = 0
                tf_b.MarginTop = 0
                tf_b.MarginRight = 0
                tr_b = tf_b.TextRange
                tr_b.Text = "\r".join([f"•  {b}" for b in bullets])
                tr_b.Font.Name = "맑은 고딕"
                tr_b.Font.Size = 11.0
                tr_b.Font.Color.RGB = theme["text_dark"]
                try:
                    for p_idx in range(1, tr_b.Paragraphs().Count + 1):
                        tr_b.Paragraphs(p_idx).ParagraphFormat.SpaceAfter = 6.0 # 단락간격
                except Exception:
                    pass
                created.append(b_sh.Id)

        return created

    # =========================================================================
    # 3. Process Flow Slice (REA Aesthetic Upgrade)
    # =========================================================================
    @classmethod
    def render_process_flow(cls, slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        steps = data.get("steps") or []
        n = max(1, len(steps))
        created: list[int] = []

        gap = 16.0
        arrow_w = 20.0
        total_gaps = (n - 1) * (gap + arrow_w)
        card_w = (w - total_gaps) / n
        card_h = min(h, 240.0)
        card_y = y + (h - card_h) / 2

        curr_x = x
        for i, s in enumerate(steps):
            step_num = str(s.get("step_num") or s.get("step") or f"0{i+1}").replace("STEP", "").strip()
            title = str(s.get("title", f"단계 {i+1}")).strip()
            desc = str(s.get("desc", "")).strip()

            # 1. 소프트 엘리베이티드 카드
            box = cls._create_elevated_card(slide, curr_x, card_y, card_w, card_h, theme)
            created.append(box.Id)

            # 2. 스텝 원형/필 배지
            badge = cls._create_pill_badge(slide, curr_x + 14, card_y + 14, f"STEP {step_num}", theme, bg_color=theme["primary"], text_color=theme["white"], w=60.0)
            created.append(badge.Id)

            # 3. 단계 타이틀
            t_sh = slide.Shapes.AddTextbox(1, curr_x + 14, card_y + 44, card_w - 28, 36)
            t_sh.TextFrame.MarginLeft = 0
            t_sh.TextFrame.MarginTop = 0
            t_sh.TextFrame.WordWrap = -1
            t_tr = t_sh.TextFrame.TextRange
            t_tr.Text = title
            t_tr.Font.Name = "맑은 고딕"
            t_tr.Font.Size = 13.0
            t_tr.Font.Bold = True
            t_tr.Font.Color.RGB = theme["text_dark"]
            created.append(t_sh.Id)

            # 4. 단계 세부 설명
            if desc:
                d_sh = slide.Shapes.AddTextbox(1, curr_x + 14, card_y + 84, card_w - 28, card_h - 96)
                tf_d = d_sh.TextFrame
                tf_d.MarginLeft = 0
                tf_d.MarginTop = 0
                tf_d.WordWrap = -1
                d_tr = tf_d.TextRange
                d_tr.Text = desc
                d_tr.Font.Name = "맑은 고딕"
                d_tr.Font.Size = 10.0
                d_tr.Font.Color.RGB = theme["text_muted"]
                created.append(d_sh.Id)

            curr_x += card_w

            # 5. 연결 화살표 커넥터
            if i < n - 1:
                arr_x = curr_x + (gap / 2) - 2
                arr_y = card_y + (card_h / 2) - 7
                arr = slide.Shapes.AddShape(MSO_SHAPE_RIGHT_ARROW, arr_x, arr_y, arrow_w, 14)
                arr.Fill.Solid()
                arr.Fill.ForeColor.RGB = theme["secondary"]
                arr.Line.Visible = 0
                created.append(arr.Id)
                curr_x += gap + arrow_w

        return created

    # =========================================================================
    # 4. Comparison Slice
    # =========================================================================
    @staticmethod
    def render_comparison(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        created: list[int] = []

        gap = 24.0
        panel_w = (w - gap) / 2
        panel_h = h

        l_dict = data.get("left_panel") or {}
        r_dict = data.get("right_panel") or {}
        left_title = str(l_dict.get("title") or data.get("left_title", "현재 상태 (AS-IS)")).strip()
        left_items = l_dict.get("bullets") or data.get("left_items") or []
        right_title = str(r_dict.get("title") or data.get("right_title", "개선 방향 (TO-BE)")).strip()
        right_items = r_dict.get("bullets") or data.get("right_items") or []

        # Left Panel (AS-IS)
        box_l = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, panel_w, panel_h)
        box_l.Fill.Solid()
        box_l.Fill.ForeColor.RGB = hex_to_bgr("#FEF2F2")
        box_l.Line.ForeColor.RGB = hex_to_bgr("#FCA5A5")
        box_l.Line.Weight = 1.5
        created.append(box_l.Id)

        hbar_l = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, panel_w, 34)
        hbar_l.Fill.Solid()
        hbar_l.Fill.ForeColor.RGB = hex_to_bgr("#DC2626")
        hbar_l.Line.Visible = 0
        hbar_l.TextFrame.TextRange.Text = f"⚠️  {left_title}"
        hbar_l.TextFrame.TextRange.Font.Name = "맑은 고딕"
        hbar_l.TextFrame.TextRange.Font.Size = 13
        hbar_l.TextFrame.TextRange.Font.Bold = True
        hbar_l.TextFrame.TextRange.Font.Color.RGB = theme["white"]
        created.append(hbar_l.Id)

        if left_items:
            tb_l = slide.Shapes.AddTextbox(1, x + 14, y + 44, panel_w - 28, panel_h - 54)
            tb_l.TextFrame.WordWrap = -1
            tr_l = tb_l.TextFrame.TextRange
            tr_l.Text = "\r".join([f"•  {it}" for it in left_items])
            tr_l.Font.Name = "맑은 고딕"
            tr_l.Font.Size = 11.5
            tr_l.Font.Color.RGB = theme["text_dark"]
            created.append(tb_l.Id)

        # Right Panel (TO-BE)
        rx = x + panel_w + gap
        box_r = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, rx, y, panel_w, panel_h)
        box_r.Fill.Solid()
        box_r.Fill.ForeColor.RGB = hex_to_bgr("#ECFDF5")
        box_r.Line.ForeColor.RGB = hex_to_bgr("#6EE7B7")
        box_r.Line.Weight = 1.5
        created.append(box_r.Id)

        hbar_r = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, rx, y, panel_w, 34)
        hbar_r.Fill.Solid()
        hbar_r.Fill.ForeColor.RGB = hex_to_bgr("#059669")
        hbar_r.Line.Visible = 0
        hbar_r.TextFrame.TextRange.Text = f"✨  {right_title}"
        hbar_r.TextFrame.TextRange.Font.Name = "맑은 고딕"
        hbar_r.TextFrame.TextRange.Font.Size = 13
        hbar_r.TextFrame.TextRange.Font.Bold = True
        hbar_r.TextFrame.TextRange.Font.Color.RGB = theme["white"]
        created.append(hbar_r.Id)

        if right_items:
            tb_r = slide.Shapes.AddTextbox(1, rx + 14, y + 44, panel_w - 28, panel_h - 54)
            tb_r.TextFrame.WordWrap = -1
            tr_r = tb_r.TextFrame.TextRange
            tr_r.Text = "\r".join([f"•  {it}" for it in right_items])
            tr_r.Font.Name = "맑은 고딕"
            tr_r.Font.Size = 11.5
            tr_r.Font.Color.RGB = theme["text_dark"]
            created.append(tb_r.Id)

        return created


    # =========================================================================
    # 5. KPI Cards Slice (REA Aesthetic Upgrade)
    # =========================================================================
    @classmethod
    def render_kpi_cards(cls, slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        kpis = data.get("cards") or data.get("kpis") or []
        n = max(1, len(kpis))
        created: list[int] = []

        spacing = 20.0
        card_w = (w - (spacing * (n - 1))) / n
        card_h = min(h, 140.0)
        card_y = y + (h - card_h) / 2

        for i, k in enumerate(kpis):
            cx = x + i * (card_w + spacing)
            label = str(k.get("label", "핵심 지표")).strip()
            val = str(k.get("value", "0")).strip()
            unit = str(k.get("unit", "")).strip()
            change = str(k.get("delta") or k.get("change") or "").strip()
            direction = str(k.get("direction", "up")).lower()

            # 1. 소프트 엘리베이티드 카드
            box = cls._create_elevated_card(slide, cx, card_y, card_w, card_h, theme)
            created.append(box.Id)

            # 2. 상단 지표 라벨 (11pt Muted)
            lbl_sh = slide.Shapes.AddTextbox(1, cx + 16, card_y + 14, card_w - 32, 22)
            lbl_sh.TextFrame.MarginLeft = 0
            lbl_sh.TextFrame.MarginTop = 0
            lbl_tr = lbl_sh.TextFrame.TextRange
            lbl_tr.Text = label
            lbl_tr.Font.Name = "맑은 고딕"
            lbl_tr.Font.Size = 11.0
            lbl_tr.Font.Bold = True
            lbl_tr.Font.Color.RGB = theme["text_muted"]
            created.append(lbl_sh.Id)

            # 3. 대형 핵심 수치 (32pt Bold) 및 단위(15pt Bold) 베이스라인 일치화
            val_sh = slide.Shapes.AddTextbox(1, cx + 16, card_y + 36, card_w - 32, 45)
            val_sh.TextFrame.MarginLeft = 0
            val_sh.TextFrame.MarginTop = 0
            val_tr = val_sh.TextFrame.TextRange
            val_tr.Text = f"{val} "
            val_tr.Font.Name = "맑은 고딕"
            val_tr.Font.Size = 32.0
            val_tr.Font.Bold = True
            val_tr.Font.Color.RGB = theme["primary"]
            if unit:
                unit_run = val_tr.InsertAfter(unit)
                unit_run.Font.Name = "맑은 고딕"
                unit_run.Font.Size = 15.0
                unit_run.Font.Bold = True
                unit_run.Font.Color.RGB = theme["primary"]
            created.append(val_sh.Id)

            # 4. 하단 증감율 마이크로 알약 배지
            if change:
                is_up = direction == "up" or change.startswith("+")
                badge_bg = hex_to_bgr("#DCFCE7") if is_up else hex_to_bgr("#FEE2E2")
                badge_fg = hex_to_bgr("#15803D") if is_up else hex_to_bgr("#B91C1C")
                badge_txt = f"▲ {change}" if is_up else f"▼ {change}"

                badge = cls._create_pill_badge(slide, cx + 16, card_y + 92, badge_txt, theme, bg_color=badge_bg, text_color=badge_fg, w=max(56.0, len(badge_txt)*9.0), h=22.0)
                created.append(badge.Id)

        return created

    # =========================================================================
    # 6. Smart Table Slice
    # =========================================================================
    @staticmethod
    def render_smart_table(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        headers = data.get("headers") or ["구분", "항목", "내용"]
        rows = data.get("rows") or []
        created: list[int] = []

        num_rows = len(rows) + 1
        num_cols = len(headers)

        table_shape = slide.Shapes.AddTable(num_rows, num_cols, x, y, w, min(h, num_rows * 34.0))
        tbl = table_shape.Table
        created.append(table_shape.Id)

        for c_idx, h_text in enumerate(headers, start=1):
            cell = tbl.Cell(1, c_idx)
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = theme["primary"]
            tr = cell.Shape.TextFrame.TextRange
            tr.Text = str(h_text)
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 11.5
            tr.Font.Bold = True
            tr.Font.Color.RGB = theme["white"]

        for r_idx, r_vals in enumerate(rows, start=2):
            bg_col = theme["bg_light"] if r_idx % 2 == 0 else theme["white"]
            for c_idx in range(1, num_cols + 1):
                val = str(r_vals[c_idx - 1]) if c_idx - 1 < len(r_vals) else ""
                cell = tbl.Cell(r_idx, c_idx)
                cell.Shape.Fill.Solid()
                cell.Shape.Fill.ForeColor.RGB = bg_col
                tr = cell.Shape.TextFrame.TextRange
                tr.Text = val
                tr.Font.Name = "맑은 고딕"
                tr.Font.Size = 10.5
                tr.Font.Color.RGB = theme["text_dark"]

                if val.startswith("[") and val.endswith("]"):
                    tr.Font.Bold = True
                    if "완료" in val or "정상" in val:
                        tr.Font.Color.RGB = hex_to_bgr("#059669")
                    elif "위험" in val or "지연" in val:
                        tr.Font.Color.RGB = hex_to_bgr("#DC2626")
                    elif "진행" in val:
                        tr.Font.Color.RGB = hex_to_bgr("#2563EB")

        return created

    # =========================================================================
    # 7. Callout Box Slice
    # =========================================================================
    @staticmethod
    def render_callout_box(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        created: list[int] = []

        tag = str(data.get("title") or data.get("tag") or "시사점").strip()
        text = str(data.get("text", "")).strip()

        box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, w, h)
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = theme["bg_light"]
        box.Line.ForeColor.RGB = theme["card_border"]
        box.Line.Weight = 1.0

        line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, x, y, 5, h)
        line.Fill.Solid()
        line.Fill.ForeColor.RGB = theme["accent"]
        line.Line.Visible = 0

        tf = box.TextFrame
        tf.MarginLeft = 24
        tf.MarginRight = 14
        tf.MarginTop = 10
        tr = tf.TextRange
        tr.Text = f"[{tag}]  {text}"
        tr.Font.Name = "맑은 고딕"
        tr.Font.Size = 12
        tr.Font.Bold = True
        tr.Font.Color.RGB = theme["text_dark"]

        created.extend([box.Id, line.Id])
        return created

    # =========================================================================
    # 8. Timeline Roadmap Slice
    # =========================================================================
    @staticmethod
    def render_timeline(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        milestones = data.get("milestones") or []
        n = max(1, len(milestones))
        created: list[int] = []

        spacing = 15.0
        node_w = (w - (spacing * (n - 1))) / n

        bar_y = y + 20.0
        hbar = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, x, bar_y, w, 4)
        hbar.Fill.Solid()
        hbar.Fill.ForeColor.RGB = theme["secondary"]
        hbar.Line.Visible = 0
        created.append(hbar.Id)

        for i, m in enumerate(milestones):
            nx = x + i * (node_w + spacing)
            phase = str(m.get("phase", f"Phase {i+1}")).strip()
            title = str(m.get("title", "")).strip()
            desc = str(m.get("desc", "")).strip()

            c_dia = 16.0
            cx = nx + (node_w / 2) - (c_dia / 2)
            cy = bar_y - (c_dia / 2) + 2
            circle = slide.Shapes.AddShape(MSO_SHAPE_OVAL, cx, cy, c_dia, c_dia)
            circle.Fill.Solid()
            circle.Fill.ForeColor.RGB = theme["accent"] if i == 0 else theme["primary"]
            circle.Line.ForeColor.RGB = theme["white"]
            circle.Line.Weight = 2.0
            created.append(circle.Id)

            card_y = bar_y + 16.0
            card_h = h - 36.0
            card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, nx, card_y, node_w, card_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = theme["bg_light"]
            card.Line.ForeColor.RGB = theme["card_border"]
            card.Line.Weight = 1.0

            tf = card.TextFrame
            tf.MarginLeft = 8
            tf.MarginTop = 10
            tr = tf.TextRange
            tr.Text = f"{phase}\r{title}\r\r{desc}"
            tr.Font.Name = "맑은 고딕"

            p1 = tr.Paragraphs(1)
            p1.Font.Size = 10
            p1.Font.Bold = True
            p1.Font.Color.RGB = theme["secondary"]

            p2 = tr.Paragraphs(2)
            p2.Font.Size = 12
            p2.Font.Bold = True
            p2.Font.Color.RGB = theme["primary"]

            if desc:
                p4 = tr.Paragraphs(4)
                p4.Font.Size = 10
                p4.Font.Color.RGB = theme["text_dark"]

            created.append(card.Id)

        return created

    # =========================================================================
    # 9. Matrix 2x2 Slice (신규 확장)
    # =========================================================================
    @staticmethod
    def render_matrix_2x2(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        created: list[int] = []

        x_axis = str(data.get("x_label") or data.get("x_axis", "실행 난이도 (저 ➔ 고)")).strip()
        y_axis = str(data.get("y_label") or data.get("y_axis", "기대 효과 (저 ➔ 고)")).strip()
        quads = data.get("quadrants") or []

        axis_margin = 25.0
        grid_w = w - axis_margin
        grid_h = h - axis_margin
        cell_w = (grid_w - 14.0) / 2
        cell_h = (grid_h - 14.0) / 2

        # Axes labels
        sh_y = slide.Shapes.AddTextbox(1, x, y, axis_margin, grid_h)
        sh_y.TextFrame.WordWrap = -1
        tr_y = sh_y.TextFrame.TextRange
        tr_y.Text = f"▲\r{y_axis}"
        tr_y.Font.Name = "맑은 고딕"
        tr_y.Font.Size = 9.5
        tr_y.Font.Bold = True
        tr_y.Font.Color.RGB = theme["text_muted"]
        created.append(sh_y.Id)

        # 4 Quadrants
        quad_positions = [
            (x + axis_margin, y, "2사분면"),                          # top-left
            (x + axis_margin + cell_w + 14.0, y, "1사분면 (핵심)"),   # top-right
            (x + axis_margin, y + cell_h + 14.0, "3사분면"),          # bottom-left
            (x + axis_margin + cell_w + 14.0, y + cell_h + 14.0, "4사분면"), # bottom-right
        ]

        for i, (qx, qy, def_lbl) in enumerate(quad_positions):
            q_info = quads[i] if i < len(quads) else {}
            title = str(q_info.get("title", def_lbl)).strip()
            items = q_info.get("items") or []

            # 1사분면은 테마 포인트 컬러 강조
            is_priority = (i == 1)
            box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, qx, qy, cell_w, cell_h)
            box.Fill.Solid()
            box.Fill.ForeColor.RGB = hex_to_bgr("#EFF6FF") if is_priority else theme["bg_light"]
            box.Line.ForeColor.RGB = theme["primary"] if is_priority else theme["card_border"]
            box.Line.Weight = 1.5 if is_priority else 1.0
            created.append(box.Id)

            tb = slide.Shapes.AddTextbox(1, qx + 8, qy + 6, cell_w - 16, cell_h - 12)
            tb.TextFrame.WordWrap = -1
            tr = tb.TextFrame.TextRange
            items_text = "\r".join([f"•  {it}" for it in items])
            tr.Text = f"[{title}]\r{items_text}"
            tr.Font.Name = "맑은 고딕"

            p1 = tr.Paragraphs(1)
            p1.Font.Size = 11.5
            p1.Font.Bold = True
            p1.Font.Color.RGB = theme["primary"]

            for p_idx in range(2, tr.Paragraphs().Count + 1):
                p = tr.Paragraphs(p_idx)
                p.Font.Size = 10
                p.Font.Color.RGB = theme["text_dark"]

            created.append(tb.Id)

        return created

    # =========================================================================
    # 10. Pyramid Hierarchy Slice (신규 확장)
    # =========================================================================
    @staticmethod
    def render_pyramid_hierarchy(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        layers = data.get("layers") or []
        n = max(1, len(layers))
        created: list[int] = []

        spacing = 10.0
        layer_h = (h - (spacing * (n - 1))) / n

        # Layer width ratios from top to base (e.g. 50% -> 75% -> 100%)
        width_steps = [0.55 + 0.45 * (i / max(1, n - 1)) for i in range(n)]

        for i, l in enumerate(layers):
            ratio = width_steps[i]
            cur_w = w * ratio
            cur_x = x + (w - cur_w) / 2
            cur_y = y + i * (layer_h + spacing)

            level_tag = str(l.get("level", f"Tier {i+1}")).strip()
            title = str(l.get("title", "")).strip()
            desc = str(l.get("desc", "")).strip()

            box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, cur_x, cur_y, cur_w, layer_h)
            box.Fill.Solid()
            box.Fill.ForeColor.RGB = theme["primary"] if i == 0 else (theme["secondary"] if i == 1 else theme["bg_light"])
            box.Line.ForeColor.RGB = theme["primary"]
            box.Line.Weight = 1.2
            created.append(box.Id)

            tf = box.TextFrame
            tf.MarginLeft = 14
            tf.MarginRight = 14
            tf.MarginTop = 8
            tr = tf.TextRange
            tr.Text = f"[{level_tag}]  {title}\r{desc}" if desc else f"[{level_tag}]  {title}"
            tr.Font.Name = "맑은 고딕"

            p1 = tr.Paragraphs(1)
            p1.Font.Size = 12.5
            p1.Font.Bold = True
            p1.Font.Color.RGB = theme["white"] if i < 2 else theme["primary"]

            if desc and tr.Paragraphs().Count >= 2:
                p2 = tr.Paragraphs(2)
                p2.Font.Size = 10.5
                p2.Font.Color.RGB = theme["white"] if i < 2 else theme["text_dark"]

        return created

    # =========================================================================
    # 11. List Bullets Slice (KISA/HWP 표준 개조식 4단 위계, 신규 확장)
    # =========================================================================
    @staticmethod
    def render_list_bullets(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        items = data.get("items") or []
        created: list[int] = []

        box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, w, h)
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = theme["white"]
        box.Line.ForeColor.RGB = theme["card_border"]
        box.Line.Weight = 1.0
        created.append(box.Id)

        tb = slide.Shapes.AddTextbox(1, x + 16, y + 14, w - 32, h - 28)
        tb.TextFrame.WordWrap = -1
        tr = tb.TextFrame.TextRange

        prefix_map = {1: "■  ", 2: "  ○  ", 3: "    -  ", 4: "      ※  "}
        size_map = {1: 13.5, 2: 12.0, 3: 11.0, 4: 9.5}
        bold_map = {1: True, 2: False, 3: False, 4: False}

        lines = []
        for it in items:
            lvl = int(it.get("level", 1))
            txt = str(it.get("text", "")).strip()
            pref = prefix_map.get(lvl, "•  ")
            lines.append(f"{pref}{txt}")

        tr.Text = "\r".join(lines)
        tr.Font.Name = "맑은 고딕"

        for idx, it in enumerate(items, start=1):
            if idx <= tr.Paragraphs().Count:
                p = tr.Paragraphs(idx)
                lvl = int(it.get("level", 1))
                p.Font.Size = size_map.get(lvl, 11.0)
                p.Font.Bold = bold_map.get(lvl, False)
                if lvl == 1:
                    p.Font.Color.RGB = theme["primary"]
                elif lvl == 4:
                    p.Font.Color.RGB = theme["text_muted"]
                else:
                    p.Font.Color.RGB = theme["text_dark"]

        created.append(tb.Id)
        return created

    # =========================================================================
    # 12. Agenda TOC Slice (신규 확장)
    # =========================================================================
    @staticmethod
    def render_agenda_toc(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        agendas = data.get("agendas") or []
        n = max(1, len(agendas))
        created: list[int] = []

        spacing = 14.0
        row_h = (h - (spacing * (n - 1))) / n

        for i, a in enumerate(agendas):
            cur_y = y + i * (row_h + spacing)
            num = str(a.get("num", f"0{i+1}")).strip()
            title = str(a.get("title", f"주제 {i+1}")).strip()
            desc = str(a.get("desc", "")).strip()

            # Number badge
            num_w = 44.0
            num_box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, cur_y, num_w, row_h)
            num_box.Fill.Solid()
            num_box.Fill.ForeColor.RGB = theme["primary"]
            num_box.Line.Visible = 0
            num_tr = num_box.TextFrame.TextRange
            num_tr.Text = num
            num_tr.Font.Name = "맑은 고딕"
            num_tr.Font.Size = 16
            num_tr.Font.Bold = True
            num_tr.Font.Color.RGB = theme["white"]
            created.append(num_box.Id)

            # Agenda row band
            band_x = x + num_w + 10.0
            band_w = w - num_w - 10.0
            band = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, band_x, cur_y, band_w, row_h)
            band.Fill.Solid()
            band.Fill.ForeColor.RGB = theme["bg_light"]
            band.Line.ForeColor.RGB = theme["card_border"]
            band.Line.Weight = 1.0

            tf = band.TextFrame
            tf.MarginLeft = 14
            tf.MarginTop = 8
            tr = tf.TextRange
            tr.Text = f"{title}\r{desc}" if desc else title
            tr.Font.Name = "맑은 고딕"

            p1 = tr.Paragraphs(1)
            p1.Font.Size = 13.5
            p1.Font.Bold = True
            p1.Font.Color.RGB = theme["primary"]

            if desc and tr.Paragraphs().Count >= 2:
                p2 = tr.Paragraphs(2)
                p2.Font.Size = 10.5
                p2.Font.Color.RGB = theme["text_dark"]

            created.append(band.Id)

        return created

    # =========================================================================
    # 13. Quote Stat Slice (신규 확장)
    # =========================================================================
    @staticmethod
    def render_quote_stat(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        created: list[int] = []

        quote = str(data.get("quote", "생각의 힘이 모든 변화의 출발점입니다.")).strip()
        author = str(data.get("author", "")).strip()
        affiliation = str(data.get("affiliation", "")).strip()

        box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, w, h)
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = theme["bg_light"]
        box.Line.ForeColor.RGB = theme["card_border"]
        box.Line.Weight = 1.2
        created.append(box.Id)

        # Quotation mark deco
        deco = slide.Shapes.AddTextbox(1, x + 16, y + 6, 60, 40)
        deco.TextFrame.TextRange.Text = "“"
        deco.TextFrame.TextRange.Font.Name = "맑은 고딕"
        deco.TextFrame.TextRange.Font.Size = 44
        deco.TextFrame.TextRange.Font.Bold = True
        deco.TextFrame.TextRange.Font.Color.RGB = theme["accent"]
        created.append(deco.Id)

        # Quote text
        tb = slide.Shapes.AddTextbox(1, x + 30, y + 40, w - 60, h - 80)
        tb.TextFrame.WordWrap = -1
        tr = tb.TextFrame.TextRange
        tr.Text = f"\"{quote}\""
        tr.Font.Name = "맑은 고딕"
        tr.Font.Size = 18
        tr.Font.Bold = True
        tr.Font.Color.RGB = theme["primary"]
        created.append(tb.Id)

        # Author / Source badge
        if author or affiliation:
            auth_box = slide.Shapes.AddTextbox(1, x + 30, y + h - 38, w - 60, 26)
            tr_a = auth_box.TextFrame.TextRange
            tr_a.Text = f"— {author} ({affiliation})" if affiliation else f"— {author}"
            tr_a.Font.Name = "맑은 고딕"
            tr_a.Font.Size = 11
            tr_a.Font.Bold = True
            tr_a.Font.Color.RGB = theme["text_muted"]
            created.append(auth_box.Id)

        return created

    # =========================================================================
    # 14. Roadmap Quarterly Slice (신규 확장)
    # =========================================================================
    @staticmethod
    def render_roadmap_quarterly(slide: Any, bounds: tuple[float, float, float, float], data: dict[str, Any], theme: dict[str, int]) -> list[int]:
        x, y, w, h = bounds
        quarters = data.get("quarters") or []
        n = max(1, len(quarters))
        created: list[int] = []

        spacing = 14.0
        col_w = (w - (spacing * (n - 1))) / n
        col_h = h

        for i, q in enumerate(quarters):
            cx = x + i * (col_w + spacing)
            q_lbl = str(q.get("quarter", f"Q{i+1}")).strip()
            focus = str(q.get("focus", "")).strip()
            items = q.get("items") or []

            # Background card
            box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, cx, y, col_w, col_h)
            box.Fill.Solid()
            box.Fill.ForeColor.RGB = theme["white"]
            box.Line.ForeColor.RGB = theme["primary"] if i == 0 else theme["card_border"]
            box.Line.Weight = 1.5 if i == 0 else 1.0
            created.append(box.Id)

            # Quarter header
            hbar_h = 32.0
            hbar = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, cx, y, col_w, hbar_h)
            hbar.Fill.Solid()
            hbar.Fill.ForeColor.RGB = theme["primary"] if i == 0 else theme["secondary"]
            hbar.Line.Visible = 0
            tr_h = hbar.TextFrame.TextRange
            tr_h.Text = q_lbl
            tr_h.Font.Name = "맑은 고딕"
            tr_h.Font.Size = 12.5
            tr_h.Font.Bold = True
            tr_h.Font.Color.RGB = theme["white"]
            created.append(hbar.Id)

            # Focus sub-header
            curr_y = y + hbar_h + 8
            if focus:
                f_box = slide.Shapes.AddTextbox(1, cx + 8, curr_y, col_w - 16, 22)
                f_tr = f_box.TextFrame.TextRange
                f_tr.Text = f"🎯 {focus}"
                f_tr.Font.Name = "맑은 고딕"
                f_tr.Font.Size = 11
                f_tr.Font.Bold = True
                f_tr.Font.Color.RGB = theme["accent"]
                created.append(f_box.Id)
                curr_y += 24

            # Items
            if items:
                tb = slide.Shapes.AddTextbox(1, cx + 8, curr_y, col_w - 16, col_h - (curr_y - y) - 8)
                tb.TextFrame.WordWrap = -1
                tr_b = tb.TextFrame.TextRange
                tr_b.Text = "\r".join([f"• {it}" for it in items])
                tr_b.Font.Name = "맑은 고딕"
                tr_b.Font.Size = 10
                tr_b.Font.Color.RGB = theme["text_dark"]
                created.append(tb.Id)

        return created

    # =========================================================================
    # 최신 모던 로드맵 3대 슬라이스 (v3.2.0)
    # =========================================================================
    @classmethod
    def render_roadmap_quarterly_v2(
        cls,
        slide: Any,
        bounds: tuple[float, float, float, float],
        data: dict[str, Any],
        theme: dict[str, int],
    ) -> list[int]:
        """
        [최신 모던 분기 프로그레스 로드맵 v2]
        - 상단: Q1~Q4 연결 타임 레일(Progress Rail) + 원형 앵커 노드 + 분기 라벨
        - 하단: 4열 소프트 엘리베이티드 카드 + 목표 배지 + 상세 마일스톤 불릿 + 상태 배지([완료]/[진행중]/[예정])
        """
        x, y, w, h = bounds
        quarters = data.get("quarters") or data.get("columns") or []
        if not quarters:
            quarters = [
                {"quarter": "Q1", "goal": "기반 구축", "status": "완료", "items": ["외피·내피 표준화", "환경 셋업"]},
                {"quarter": "Q2", "goal": "엔진 개발", "status": "진행중", "items": ["COM 제어 모듈", "슬라이스 엔진"]},
                {"quarter": "Q3", "goal": "시범 운영", "status": "예정", "items": ["실무 파일럿 검증", "피드백 반영"]},
                {"quarter": "Q4", "goal": "전사 확산", "status": "예정", "items": ["전사 포털 연동", "생산성 극대화"]},
            ]
        n = max(1, len(quarters))
        created: list[int] = []

        spacing = 16.0
        col_w = (w - (spacing * (n - 1))) / n

        rail_y = y + 22.0
        rail = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, x + 20, rail_y, w - 40, 3)
        rail.Fill.Solid()
        rail.Fill.ForeColor.RGB = theme["card_border"]
        rail.Line.Visible = 0
        created.append(rail.Id)

        card_top = y + 54.0
        card_h = h - 54.0

        for i, q in enumerate(quarters):
            cx = x + i * (col_w + spacing)
            q_lbl = str(q.get("quarter") or f"Q{i+1}").strip()
            goal = str(q.get("goal") or q.get("focus") or q.get("title") or "").strip()
            status = str(q.get("status") or "예정").strip()
            items = q.get("items") or q.get("bullets") or []

            # 앵커 노드
            node_size = 18.0
            node_x = cx + (col_w / 2) - (node_size / 2)
            node_y = rail_y - (node_size / 2) + 1.5
            node = slide.Shapes.AddShape(MSO_SHAPE_OVAL, node_x, node_y, node_size, node_size)
            node.Fill.Solid()
            node_color = theme["primary"] if status in ["완료", "진행중"] else theme["card_border"]
            node.Fill.ForeColor.RGB = node_color
            node.Line.Weight = 2.0
            node.Line.ForeColor.RGB = theme["white"]
            created.append(node.Id)

            # 분기 라벨
            lbl_sh = slide.Shapes.AddTextbox(1, cx, y, col_w, 20)
            lbl_sh.TextFrame.MarginLeft = 0
            lbl_sh.TextFrame.MarginTop = 0
            tr_l = lbl_sh.TextFrame.TextRange
            tr_l.Text = q_lbl
            tr_l.Font.Name = "맑은 고딕"
            tr_l.Font.Size = 12.0
            tr_l.Font.Bold = True
            tr_l.Font.Color.RGB = theme["primary"]
            tr_l.ParagraphFormat.Alignment = 2
            created.append(lbl_sh.Id)

            # 카드
            box = cls._create_elevated_card(slide, cx, card_top, col_w, card_h, theme)
            created.append(box.Id)

            if goal:
                g_sh = slide.Shapes.AddTextbox(1, cx + 12, card_top + 14, col_w - 24, 28)
                g_sh.TextFrame.MarginLeft = 0
                g_sh.TextFrame.MarginTop = 0
                gtr = g_sh.TextFrame.TextRange
                gtr.Text = goal
                gtr.Font.Name = "맑은 고딕"
                gtr.Font.Size = 13.0
                gtr.Font.Bold = True
                gtr.Font.Color.RGB = theme["text_dark"]
                created.append(g_sh.Id)

            if items:
                b_sh = slide.Shapes.AddTextbox(1, cx + 12, card_top + 46, col_w - 24, card_h - 86)
                b_tf = b_sh.TextFrame
                b_tf.MarginLeft = 0
                b_tf.MarginTop = 0
                b_tf.WordWrap = -1
                b_tr = b_tf.TextRange
                b_tr.Text = "\n".join(f"•  {it}" for it in items)
                b_tr.Font.Name = "맑은 고딕"
                b_tr.Font.Size = 10.5
                b_tr.Font.Color.RGB = theme["text_muted"]
                try:
                    b_tr.ParagraphFormat.SpaceWithin = 1.25
                except Exception:
                    pass
                b_tr.ParagraphFormat.SpaceAfter = 4.0
                created.append(b_sh.Id)

            if status:
                if status == "완료":
                    s_bg = hex_to_bgr("#DCFCE7")
                    s_fg = hex_to_bgr("#15803D")
                    s_text = "✔ 완료"
                elif status == "진행중":
                    s_bg = hex_to_bgr("#DBEAFE")
                    s_fg = hex_to_bgr("#1D4ED8")
                    s_text = "⏳ 진행중"
                else:
                    s_bg = hex_to_bgr("#F3F4F6")
                    s_fg = hex_to_bgr("#6B7280")
                    s_text = "📋 예정"

                s_badge = cls._create_pill_badge(slide, cx + 12, card_top + card_h - 32, s_text, theme, bg_color=s_bg, text_color=s_fg, w=64.0, h=20.0)
                created.append(s_badge.Id)

        return created

    @classmethod
    def render_roadmap_gantt(
        cls,
        slide: Any,
        bounds: tuple[float, float, float, float],
        data: dict[str, Any],
        theme: dict[str, int],
    ) -> list[int]:
        """
        [트랙별 멀티 레인 간트차트형 로드맵]
        - Y축 3~4개 업무 트랙(전략/기획, 시스템 구축, 운영/확산)
        - X축 Q1~Q4 타임라인 헤더 및 세로 가이드라인
        - 트랙별 간트 진행 바(Gantt Bar) 시각화
        """
        x, y, w, h = bounds
        tracks = data.get("tracks") or []
        if not tracks:
            tracks = [
                {"track": "전략 및 표준화", "bars": [{"title": "외피/내피 아키텍처 수립", "start": 0.0, "end": 0.35, "color": "primary"}, {"title": "전사 가이드라인 배포", "start": 0.40, "end": 0.70, "color": "secondary"}]},
                {"track": "플랫폼 개발", "bars": [{"title": "MCP 데스크톱 바인딩", "start": 0.15, "end": 0.60, "color": "primary"}, {"title": "슬라이스 엔진 v3.2 구축", "start": 0.50, "end": 1.0, "color": "accent"}]},
                {"track": "운영 및 확산", "bars": [{"title": "파일럿 검증", "start": 0.55, "end": 0.85, "color": "secondary"}, {"title": "전사 운영 체계 전환", "start": 0.80, "end": 1.0, "color": "primary"}]}
            ]
        quarters = data.get("headers") or ["Q1 (1~3월)", "Q2 (4~6월)", "Q3 (7~9월)", "Q4 (10~12월)"]
        created: list[int] = []

        track_col_w = 140.0
        timeline_w = w - track_col_w
        num_q = len(quarters)
        q_w = timeline_w / num_q

        header_h = 32.0
        for i, q in enumerate(quarters):
            qx = x + track_col_w + i * q_w
            qh_box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, qx + 2, y, q_w - 4, header_h)
            qh_box.Fill.Solid()
            qh_box.Fill.ForeColor.RGB = theme["bg_light"]
            qh_box.Line.ForeColor.RGB = theme["card_border"]
            qh_box.Line.Weight = 0.75
            tr = qh_box.TextFrame.TextRange
            tr.Text = q
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 11.0
            tr.Font.Bold = True
            tr.Font.Color.RGB = theme["primary"]
            tr.ParagraphFormat.Alignment = 2
            created.append(qh_box.Id)

        num_tracks = len(tracks)
        row_gap = 12.0
        row_h = (h - header_h - 12.0 - (row_gap * (num_tracks - 1))) / max(1, num_tracks)

        for r_idx, trk in enumerate(tracks):
            ry = y + header_h + 12.0 + r_idx * (row_h + row_gap)
            track_name = str(trk.get("track", f"트랙 {r_idx+1}")).strip()

            t_card = cls._create_elevated_card(slide, x, ry, track_col_w - 10, row_h, theme)
            tr_t = t_card.TextFrame.TextRange
            tr_t.Text = track_name
            tr_t.Font.Name = "맑은 고딕"
            tr_t.Font.Size = 12.0
            tr_t.Font.Bold = True
            tr_t.Font.Color.RGB = theme["text_dark"]
            tr_t.ParagraphFormat.Alignment = 2
            created.append(t_card.Id)

            bg_grid = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x + track_col_w, ry, timeline_w, row_h)
            bg_grid.Fill.Solid()
            bg_grid.Fill.ForeColor.RGB = hex_to_bgr("#FAFBFD")
            bg_grid.Line.ForeColor.RGB = theme["card_border"]
            bg_grid.Line.Weight = 0.5
            created.append(bg_grid.Id)

            bars = trk.get("bars") or []
            for b in bars:
                b_title = str(b.get("title", "")).strip()
                s_ratio = float(b.get("start", 0.0))
                e_ratio = float(b.get("end", 1.0))
                b_x = x + track_col_w + s_ratio * timeline_w + 4
                b_w = max(40.0, (e_ratio - s_ratio) * timeline_w - 8)
                b_h = max(24.0, row_h * 0.55)
                b_y = ry + (row_h - b_h) / 2

                color_key = b.get("color", "primary")
                fill_color = theme.get(color_key, theme["primary"])

                gbar = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, b_x, b_y, b_w, b_h)
                try:
                    if hasattr(gbar, "Adjustments") and gbar.Adjustments.Count > 0:
                        gbar.Adjustments.Item(1, 0.25)
                except Exception:
                    pass
                gbar.Fill.Solid()
                gbar.Fill.ForeColor.RGB = fill_color
                gbar.Line.Visible = 0

                tr_b = gbar.TextFrame.TextRange
                tr_b.Text = b_title
                tr_b.Font.Name = "맑은 고딕"
                tr_b.Font.Size = 10.0
                tr_b.Font.Bold = True
                tr_b.Font.Color.RGB = theme["white"]
                tr_b.ParagraphFormat.Alignment = 2
                created.append(gbar.Id)

        return created

    @classmethod
    def render_roadmap_chevron(
        cls,
        slide: Any,
        bounds: tuple[float, float, float, float],
        data: dict[str, Any],
        theme: dict[str, int],
    ) -> list[int]:
        """
        [단계별 셰브런 페이즈 로드맵]
        - 상단 3~5단계 셰브런 헤더 밴드
        - 하단 단계별 상세 실행 과제 카드
        """
        x, y, w, h = bounds
        phases = data.get("phases") or data.get("steps") or []
        if not phases:
            phases = [
                {"phase": "PHASE 01", "title": "사전 기획", "items": ["현황 진단", "목표 정의"]},
                {"phase": "PHASE 02", "title": "모델 구축", "items": ["엔진 개발", "표준 수립"]},
                {"phase": "PHASE 03", "title": "실무 파일럿", "items": ["시범 적용", "오류 개선"]},
                {"phase": "PHASE 04", "title": "전사 표준화", "items": ["전사 배포", "성과 측정"]},
            ]
        n = max(1, len(phases))
        created: list[int] = []

        spacing = 10.0
        col_w = (w - (spacing * (n - 1))) / n
        chev_h = 36.0
        card_h = h - chev_h - 12.0
        card_y = y + chev_h + 12.0

        MSO_SHAPE_CHEVRON = 52
        for i, ph in enumerate(phases):
            cx = x + i * (col_w + spacing)
            p_lbl = str(ph.get("phase", f"PHASE 0{i+1}")).strip()
            title = str(ph.get("title", "")).strip()
            items = ph.get("items") or ph.get("bullets") or []

            chev = slide.Shapes.AddShape(MSO_SHAPE_CHEVRON, cx, y, col_w, chev_h)
            chev.Fill.Solid()
            chev_color = theme["primary"] if i % 2 == 0 else theme["secondary"]
            chev.Fill.ForeColor.RGB = chev_color
            chev.Line.Visible = 0

            tr_c = chev.TextFrame.TextRange
            tr_c.Text = f"{p_lbl} {title}".strip()
            tr_c.Font.Name = "맑은 고딕"
            tr_c.Font.Size = 11.5
            tr_c.Font.Bold = True
            tr_c.Font.Color.RGB = theme["white"]
            tr_c.ParagraphFormat.Alignment = 2
            created.append(chev.Id)

            box = cls._create_elevated_card(slide, cx, card_y, col_w, card_h, theme)
            created.append(box.Id)

            if items:
                b_sh = slide.Shapes.AddTextbox(1, cx + 12, card_y + 14, col_w - 24, card_h - 28)
                b_tf = b_sh.TextFrame
                b_tf.MarginLeft = 0
                b_tf.MarginTop = 0
                b_tf.WordWrap = -1
                b_tr = b_tf.TextRange
                b_tr.Text = "\n".join(f"✔  {it}" for it in items)
                b_tr.Font.Name = "맑은 고딕"
                b_tr.Font.Size = 11.0
                b_tr.Font.Color.RGB = theme["text_dark"]
                try:
                    b_tr.ParagraphFormat.SpaceWithin = 1.25
                except Exception:
                    pass
                b_tr.ParagraphFormat.SpaceAfter = 6.0
                created.append(b_sh.Id)

        return created


    # =========================================================================
    # Grouping Helper
    # =========================================================================
    @staticmethod
    def group_shapes(slide: Any, shape_ids: list[int]) -> Any:
        if not shape_ids or len(shape_ids) < 2:
            return None
        try:
            return slide.Shapes.Range(shape_ids).Group()
        except Exception as e:
            logger.debug(f"Failed to group shapes (non-fatal): {e}")
            return None



# =============================================================================
# COM Image Helper for Aspect Ratio Fit & Masking
# =============================================================================
class ComImageHelper:
    """PowerPoint COM 저수준 이미지 정비율 크롭, 원형 아바타 마스킹 및 플레이스홀더 헬퍼"""

    @staticmethod
    def insert_image_slot(
        slide: Any,
        img_path: str | None,
        left: float,
        top: float,
        width: float,
        height: float,
        mask_shape: str = "rectangle",
        roundness: float = 0.0,
        theme: dict[str, int] | None = None,
        border_pt: float = 0.0,
        border_color: int | None = None,
    ) -> int | None:
        """
        슬롯 영역에 이미지를 삽입하거나 플레이스홀더를 생성.
        원형 아바타, 둥근 모서리 카드의 경우 AutoShape.Fill.UserPicture를 활용하여
        가장 안정적으로 마스킹 및 비율 유지를 처리함.
        """
        import os

        msoShapeRectangle = 1
        msoShapeRoundedRectangle = 5
        msoShapeOval = 9

        target_shape_type = msoShapeRectangle
        if mask_shape in ("oval", "msoShapeOval"):
            target_shape_type = msoShapeOval
        elif mask_shape in ("rounded_rectangle", "msoShapeRoundedRectangle") or roundness > 0:
            target_shape_type = msoShapeRoundedRectangle

        # 실제 이미지 파일이 존재할 경우
        if img_path and os.path.exists(img_path):
            abs_img_path = os.path.abspath(img_path)
            try:
                shape = slide.Shapes.AddShape(target_shape_type, left, top, width, height)
                shape.Fill.UserPicture(abs_img_path)

                if border_pt > 0 and border_color is not None:
                    shape.Line.Visible = -1
                    shape.Line.ForeColor.RGB = border_color
                    shape.Line.Weight = border_pt
                else:
                    shape.Line.Visible = 0

                if target_shape_type == msoShapeRoundedRectangle and roundness > 0:
                    try:
                        shape.Adjustments.Item(1, roundness)
                    except Exception:
                        pass
                return shape.Id
            except Exception as e:
                logger.warning(f"Failed to insert picture via Fill.UserPicture: {e}")

        # 이미지가 없거나 삽입 실패 시: 디자이너 감성의 점선 벡터 플레이스홀더 생성
        try:
            placeholder = slide.Shapes.AddShape(target_shape_type, left, top, width, height)
            placeholder.Fill.Visible = -1
            placeholder.Fill.Solid()
            placeholder.Fill.ForeColor.RGB = theme.get("surface", 0xF3F4F6) if theme else 0xF3F4F6
            placeholder.Line.Visible = -1
            placeholder.Line.ForeColor.RGB = theme.get("border", 0xD1D5DB) if theme else 0xD1D5DB
            placeholder.Line.Weight = 1.0
            try:
                placeholder.Line.DashStyle = 2  # 점선 테두리
            except Exception:
                pass

            if target_shape_type == msoShapeRoundedRectangle and roundness > 0:
                try:
                    placeholder.Adjustments.Item(1, roundness)
                except Exception:
                    pass

            tf = placeholder.TextFrame
            tf.WordWrap = -1
            tr = tf.TextRange
            tr.Text = "🖼 Image Slot" if target_shape_type != msoShapeOval else "👤 Profile"
            tr.Font.Name = "맑은 고딕"
            tr.Font.Size = 10.0
            tr.Font.Color.RGB = theme.get("text_muted", 0x9CA3AF) if theme else 0x9CA3AF
            return placeholder.Id
        except Exception as e:
            logger.error(f"Failed to create image placeholder: {e}")
            return None


# =============================================================================
# Zero-File Dependency Blueprint Renderer
# =============================================================================
class BlueprintRenderer:
    """
    95개 순수 기법 전수(100%) 1:1 독립 함수 디스패처.
    
    핵심 원칙:
    1. 뭉뚱그리기 완전 배제: 95개 기법 각각에 대해 작성된 전용 독립 렌더링 함수와 1:1 직결.
    2. 외피-내피 간 30pt 황금 여백 강제 (content_top = 126.0, header_bottom = 96.0).
    3. 도형 겹침 0% 및 텍스트 삐져나옴 0% 원천 보장.
    """

    _blueprints_data: dict[str, Any] | None = None

    @classmethod
    def load_blueprints(cls) -> dict[str, Any]:
        if cls._blueprints_data is None:
            bp_path = Path("data/ppt_technique_blueprints.json")
            if not bp_path.exists():
                bp_path = Path(__file__).parent.parent.parent / "data" / "ppt_technique_blueprints.json"
            if bp_path.exists():
                with open(bp_path, "r", encoding="utf-8") as f:
                    cls._blueprints_data = json.load(f)
            else:
                cls._blueprints_data = {"technique_blueprints": {}, "image_layout_blueprints": {}, "context_intent_rules": {}}
        return cls._blueprints_data

    @classmethod
    def render(
        cls,
        slide: Any,
        technique_name: str,
        data: dict[str, Any],
        theme_override: dict[str, int] | None = None,
        image_files: list[str] | None = None,
        slide_width_pt: float = 960.0,
        slide_height_pt: float = 540.0,
    ) -> list[int]:
        """95개 전용 독립 렌더링 함수 1:1 호출"""
        from .ppt_technique_functions import TECHNIQUE_FUNCTIONS_REGISTRY

        theme = {
            "primary": 0x82331F,       # Navy
            "secondary": 0x8C7355,     # Slate
            "accent": 0x2A72D9,        # Orange/Gold
            "bg_light": 0xFAFAFA,      # Surface
            "card_border": 0xDDDDDD,   # Border
            "surface": 0xFAFAFA,
            "surface_alt": 0xEEEEEE,
            "border": 0xDDDDDD,
            "text_dark": 0x222222,     # Dark Text
            "text_muted": 0x666666,    # Muted Gray Text
            "white": 0xFFFFFF,
            "danger": 0x2A2AD9,
            "success": 0x388E3C,
        }
        if theme_override:
            theme.update(theme_override)

        # 95개 기법 1:1 전용 독립 함수 실행
        target_fn = TECHNIQUE_FUNCTIONS_REGISTRY.get(technique_name)
        if target_fn:
            return target_fn(slide, data, theme, image_files=image_files)

        # 매칭 실패 시 첫 번째 독립 함수로 폴백
        fallback_fn = next(iter(TECHNIQUE_FUNCTIONS_REGISTRY.values()))
        return fallback_fn(slide, data, theme, image_files=image_files)
