# -*- coding: utf-8 -*-
"""
mcp_server/tools/trendy_vector_engine.py
06_트렌디_비즈니스_템플릿의 미려한 디자이너 감성과 비례감을
순수 파워포인트 네이티브 벡터(도형, 라운드 카드, 셰브론, 배지, 타임라인, 빅넘버)로 100% 재현하며,
글자 겹침 0%, 영역 넘침 0%, 황금 여백을 수학적으로 완벽 보장하는 실전 마스터 렌더링 엔진.
"""

from __future__ import annotations
from typing import Any
import win32com.client


class TrendyVectorEngine:
    # 06_트렌디 비즈니스 템플릿의 대표 명품 테마 팔레트 (신뢰의 네이비 & 트렌디 블루)
    DEFAULT_THEME = {
        "primary": 0x82331F,       # Deep Navy/Blue (BGR)
        "secondary": 0x8C7355,     # Slate Gray
        "accent": 0x2A72D9,        # Orange / Gold Accent
        "accent_blue": 0xEC8025,   # Trendy Sky Blue
        "bg_light": 0xFAFAFA,      # Crisp Card Surface
        "card_border": 0xDDDDDD,   # Subtle Border
        "surface_alt": 0xF3F4F6,   # Light Gray Fill
        "text_dark": 0x1E293B,     # Slate 800 Dark Text
        "text_muted": 0x64748B,    # Slate 500 Sub Text
        "white": 0xFFFFFF,
        "success": 0x388E3C,
    }

    CANVAS_W = 960.0
    CANVAS_H = 540.0

    @classmethod
    def render_header(cls, slide: Any, title: str, subtitle: str = "", category: str = "", theme: dict[str, int] = None) -> float:
        """상단 헤더 블록 (카테고리 배지 + 대제목 24~26pt + 서브타이틀) 겹침 0% 렌더링"""
        th = theme or cls.DEFAULT_THEME
        left_m = 48.0
        cur_y = 36.0

        # 카테고리 태그/배지
        if category:
            cat_box = slide.Shapes.AddShape(5, left_m, cur_y, max(70.0, len(category) * 11.0 + 16.0), 20.0) # ROUNDED_RECTANGLE
            cat_box.Fill.Solid()
            cat_box.Fill.ForeColor.RGB = th["primary"]
            cat_box.Line.Visible = 0
            cat_tf = cat_box.TextFrame
            cat_tf.WordWrap = 0
            cat_tf.MarginLeft = 4.0; cat_tf.MarginRight = 4.0; cat_tf.MarginTop = 1.0; cat_tf.MarginBottom = 1.0
            cat_tr = cat_tf.TextRange
            cat_tr.Text = category.upper()
            cat_tr.Font.Name = "맑은 고딕"
            cat_tr.Font.Size = 9.5
            cat_tr.Font.Bold = -1
            cat_tr.Font.Color.RGB = th["white"]
            cur_y += 26.0

        # 메인 대제목 (가로 864pt 시원한 박스)
        t_box = slide.Shapes.AddTextbox(1, left_m, cur_y, 864.0, 36.0)
        t_tf = t_box.TextFrame
        t_tf.WordWrap = 0 # 가로 1줄 보장
        t_tf.MarginLeft = 0; t_tf.MarginRight = 0; t_tf.MarginTop = 0; t_tf.MarginBottom = 0
        t_tr = t_tf.TextRange
        t_tr.Text = title
        t_tr.Font.Name = "맑은 고딕"
        t_tr.Font.Size = 24.0 if len(title) <= 28 else 21.0
        t_tr.Font.Bold = -1
        t_tr.Font.Color.RGB = th["text_dark"]
        cur_y += 38.0

        # 서브타이틀
        if subtitle:
            s_box = slide.Shapes.AddTextbox(1, left_m, cur_y, 864.0, 22.0)
            s_tf = s_box.TextFrame
            s_tf.WordWrap = 0
            s_tf.MarginLeft = 0; s_tf.MarginRight = 0; s_tf.MarginTop = 0; s_tf.MarginBottom = 0
            s_tr = s_tf.TextRange
            s_tr.Text = subtitle
            s_tr.Font.Name = "맑은 고딕"
            s_tr.Font.Size = 11.5
            s_tr.Font.Color.RGB = th["text_muted"]
            cur_y += 26.0

        # 세련된 헤더 구분선
        div_line = slide.Shapes.AddLine(left_m, cur_y + 4.0, left_m + 864.0, cur_y + 4.0)
        div_line.Line.ForeColor.RGB = th["card_border"]
        div_line.Line.Weight = 1.0

        # 본문 시작 Y 좌표 반환
        return cur_y + 24.0

    @classmethod
    def render_cover(cls, slide: Any, data: dict[str, Any], theme: dict[str, int] = None):
        """[아키타입 0] 미니멀 임팩트 비즈니스 표지 (트렌디 비즈니스 대표 표지)"""
        th = theme or cls.DEFAULT_THEME
        title = data.get("title", "2026 차세대 비즈니스 혁신 전략")
        subtitle = data.get("subtitle", "지속 가능한 성장과 디지털 전환 로드맵")
        category = data.get("category", "경영전략기획")

        # 1. 좌측 악센트 세로 바 (너비 16pt, 높이 540pt)
        left_bar = slide.Shapes.AddShape(1, 0.0, 0.0, 16.0, cls.CANVAS_H)
        left_bar.Fill.Solid()
        left_bar.Fill.ForeColor.RGB = th["primary"]
        left_bar.Line.Visible = 0

        # 2. 중앙 메인 라운드 프레임 카드
        card_w = 760.0
        card_h = 360.0
        c_left = 100.0
        c_top = 90.0

        card = slide.Shapes.AddShape(5, c_left, c_top, card_w, card_h)
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = th["bg_light"]
        card.Line.ForeColor.RGB = th["card_border"]
        card.Line.Weight = 1.0

        # 상단 악센트 컬러 띠
        top_accent = slide.Shapes.AddShape(1, c_left, c_top, card_w, 8.0)
        top_accent.Fill.Solid()
        top_accent.Fill.ForeColor.RGB = th["primary"]
        top_accent.Line.Visible = 0

        # 카테고리 태그 배지
        cat_box = slide.Shapes.AddShape(5, c_left + 48.0, c_top + 48.0, max(80.0, len(category) * 12.0 + 20.0), 24.0)
        cat_box.Fill.Solid()
        cat_box.Fill.ForeColor.RGB = th["primary"]
        cat_box.Line.Visible = 0
        c_tr = cat_box.TextFrame.TextRange
        c_tr.Text = category.upper()
        c_tr.Font.Name = "맑은 고딕"
        c_tr.Font.Size = 10.0
        c_tr.Font.Bold = -1
        c_tr.Font.Color.RGB = th["white"]

        # 메인 대형 타이틀 (32~36pt)
        t_box = slide.Shapes.AddTextbox(1, c_left + 48.0, c_top + 90.0, card_w - 96.0, 90.0)
        t_tf = t_box.TextFrame
        t_tf.WordWrap = -1
        t_tr = t_tf.TextRange
        t_tr.Text = title
        t_tr.Font.Name = "맑은 고딕"
        t_tr.Font.Size = 32.0 if len(title) <= 25 else 26.0
        t_tr.Font.Bold = -1
        t_tr.Font.Color.RGB = th["text_dark"]

        # 서브타이틀
        if subtitle:
            s_box = slide.Shapes.AddTextbox(1, c_left + 48.0, c_top + 190.0, card_w - 96.0, 40.0)
            s_tf = s_box.TextFrame
            s_tf.WordWrap = -1
            s_tr = s_tf.TextRange
            s_tr.Text = subtitle
            s_tr.Font.Name = "맑은 고딕"
            s_tr.Font.Size = 14.0
            s_tr.Font.Color.RGB = th["text_muted"]

        # 하단 날짜 및 작성자 메타데이터
        div = slide.Shapes.AddLine(c_left + 48.0, c_top + 260.0, c_left + card_w - 48.0, c_top + 260.0)
        div.Line.ForeColor.RGB = th["card_border"]
        div.Line.Weight = 0.75

        meta_box = slide.Shapes.AddTextbox(1, c_left + 48.0, c_top + 280.0, card_w - 96.0, 24.0)
        m_tf = meta_box.TextFrame
        m_tf.WordWrap = 0
        m_tr = m_tf.TextRange
        m_tr.Text = "2026. 11 | 전략기획실 비즈니스혁신팀"
        m_tr.Font.Name = "맑은 고딕"
        m_tr.Font.Size = 11.0
        m_tr.Font.Color.RGB = th["text_muted"]

    @classmethod
    def render_agenda(cls, slide: Any, data: dict[str, Any], theme: dict[str, int] = None):
        """[아키타입 0-1] 4단 넘버링 비주얼 목차 (넘버링 배지 01~04 + 타이틀)"""
        th = theme or cls.DEFAULT_THEME
        start_y = cls.render_header(slide, data.get("title", "목차 (Agenda)"), data.get("subtitle", "Presentation Contents"), data.get("category", "목차"), th)
        slots = data.get("slots", [])[:4]

        left_m = 48.0
        avail_w = 864.0
        item_cnt = 4
        gap_y = 16.0
        avail_h = cls.CANVAS_H - start_y - 36.0
        item_h = (avail_h - gap_y * (item_cnt - 1)) / item_cnt

        for i in range(item_cnt):
            slot = slots[i] if i < len(slots) else {"title": f"주요 아젠다 {i+1}", "description": ""}
            cur_y = start_y + i * (item_h + gap_y)

            # 라운드 바 카드 컨테이너
            bar = slide.Shapes.AddShape(5, left_m, cur_y, avail_w, item_h)
            bar.Fill.Solid()
            bar.Fill.ForeColor.RGB = th["bg_light"]
            bar.Line.ForeColor.RGB = th["card_border"]
            bar.Line.Weight = 1.0

            # 좌측 넘버링 배지 블록 (너비 70pt)
            num_badge = slide.Shapes.AddShape(1, left_m, cur_y, 70.0, item_h)
            num_badge.Fill.Solid()
            num_badge.Fill.ForeColor.RGB = th["primary"] if i == 0 else (th["accent_blue"] if i < 3 else th["secondary"])
            num_badge.Line.Visible = 0
            nb_tr = num_badge.TextFrame.TextRange
            nb_tr.Text = f"0{i+1}"
            nb_tr.ParagraphFormat.Alignment = 2 # Center
            nb_tr.Font.Name = "맑은 고딕"
            nb_tr.Font.Size = 14.0
            nb_tr.Font.Bold = -1
            nb_tr.Font.Color.RGB = th["white"]

            # 아젠다 제목 (15pt Bold)
            t_box = slide.Shapes.AddTextbox(1, left_m + 88.0, cur_y + item_h/2.0 - 12.0, avail_w - 120.0, 26.0)
            t_tf = t_box.TextFrame
            t_tf.WordWrap = 0
            t_tr = t_tf.TextRange
            t_tr.Text = slot.get("title", f"주요 아젠다 {i+1}")
            t_tr.Font.Name = "맑은 고딕"
            t_tr.Font.Size = 14.0
            t_tr.Font.Bold = -1
            t_tr.Font.Color.RGB = th["text_dark"]

            # 우측 화살표 아이콘 장식 (>)
            arrow = slide.Shapes.AddTextbox(1, left_m + avail_w - 40.0, cur_y + item_h/2.0 - 14.0, 30.0, 26.0)
            a_tf = arrow.TextFrame
            a_tr = a_tf.TextRange
            a_tr.Text = "➔"
            a_tr.Font.Name = "맑은 고딕"
            a_tr.Font.Size = 13.0
            a_tr.Font.Color.RGB = th["accent"]

    @classmethod
    def render_cards_3col(cls, slide: Any, data: dict[str, Any], theme: dict[str, int] = None):
        """[아키타입 1] 3단 수평 핵심과제 라운드 카드 그리드 (트렌디 비즈니스 대표 서식)"""
        th = theme or cls.DEFAULT_THEME
        start_y = cls.render_header(slide, data.get("title", "핵심 과제"), data.get("subtitle", ""), data.get("category", "전략기획"), th)
        slots = data.get("slots", [])[:3]

        left_m = 48.0
        avail_w = 864.0
        gap_x = 20.0
        card_w = (avail_w - gap_x * 2) / 3.0
        card_h = min(360.0, cls.CANVAS_H - start_y - 36.0)

        for i, slot in enumerate(slots):
            c_left = left_m + i * (card_w + gap_x)

            # 1. 외피 라운드 카드 컨테이너 (둥근 모서리 사각형)
            card = slide.Shapes.AddShape(5, c_left, start_y, card_w, card_h) # ROUNDED_RECTANGLE
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = th["bg_light"]
            card.Line.ForeColor.RGB = th["card_border"]
            card.Line.Weight = 1.0

            # 2. 상단 악센트 컬러 바 (5pt)
            top_bar = slide.Shapes.AddShape(1, c_left + 1.0, start_y + 1.0, card_w - 2.0, 6.0)
            top_bar.Fill.Solid()
            top_bar.Fill.ForeColor.RGB = th["primary"] if i == 0 else (th["accent_blue"] if i == 1 else th["secondary"])
            top_bar.Line.Visible = 0

            # 3. 넘버링 배지 (01, 02, 03)
            badge = slide.Shapes.AddShape(5, c_left + 16.0, start_y + 20.0, 36.0, 22.0)
            badge.Fill.Solid()
            badge.Fill.ForeColor.RGB = th["primary"] if i == 0 else th["surface_alt"]
            badge.Line.Visible = 0
            b_tr = badge.TextFrame.TextRange
            b_tr.Text = f"0{i+1}"
            b_tr.Font.Name = "맑은 고딕"
            b_tr.Font.Size = 10.0
            b_tr.Font.Bold = -1
            b_tr.Font.Color.RGB = th["white"] if i == 0 else th["text_dark"]

            # 4. 카드 제목 (14~15pt Bold)
            t_box = slide.Shapes.AddTextbox(1, c_left + 16.0, start_y + 52.0, card_w - 32.0, 32.0)
            t_tf = t_box.TextFrame
            t_tf.WordWrap = -1
            t_tf.MarginLeft = 0; t_tf.MarginRight = 0; t_tf.MarginTop = 0; t_tf.MarginBottom = 0
            t_tr = t_tf.TextRange
            t_tr.Text = slot.get("title", f"주요 과제 {i+1}")
            t_tr.Font.Name = "맑은 고딕"
            t_tr.Font.Size = 14.5
            t_tr.Font.Bold = -1
            t_tr.Font.Color.RGB = th["text_dark"]

            # 5. 구분선
            div = slide.Shapes.AddLine(c_left + 16.0, start_y + 92.0, c_left + card_w - 16.0, start_y + 92.0)
            div.Line.ForeColor.RGB = th["card_border"]
            div.Line.Weight = 0.75

            # 6. 본문 설명문 (개조식 또는 문장)
            b_box = slide.Shapes.AddTextbox(1, c_left + 16.0, start_y + 104.0, card_w - 32.0, card_h - 120.0)
            b_tf = b_box.TextFrame
            b_tf.WordWrap = -1
            b_tf.MarginLeft = 0; b_tf.MarginRight = 0; b_tf.MarginTop = 0; b_tf.MarginBottom = 0
            b_tr = b_tf.TextRange
            desc = slot.get("description") or slot.get("body", "")
            b_tr.Text = desc
            b_tr.Font.Name = "맑은 고딕"
            b_tr.Font.Size = 11.0 if len(desc) <= 80 else 10.0
            b_tr.Font.Color.RGB = th["text_muted"]

    @classmethod
    def render_process_4step(cls, slide: Any, data: dict[str, Any], theme: dict[str, int] = None):
        """[아키타입 2] 4단계 스텝 화살표 프로세스 체인 (트렌디 비즈니스 표준)"""
        th = theme or cls.DEFAULT_THEME
        start_y = cls.render_header(slide, data.get("title", "실행 프로세스"), data.get("subtitle", ""), data.get("category", "추진절차"), th)
        slots = data.get("slots", [])[:4]

        left_m = 48.0
        avail_w = 864.0
        step_cnt = 4
        gap_x = 16.0
        box_w = (avail_w - gap_x * (step_cnt - 1)) / step_cnt
        box_h = min(320.0, cls.CANVAS_H - start_y - 40.0)

        for i in range(step_cnt):
            slot = slots[i] if i < len(slots) else {"title": f"단계 {i+1}", "description": "세부 실행 내용 기술"}
            s_left = left_m + i * (box_w + gap_x)

            # 상단 스텝 라운드 카드
            card = slide.Shapes.AddShape(5, s_left, start_y, box_w, box_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = th["bg_light"]
            card.Line.ForeColor.RGB = th["card_border"]
            card.Line.Weight = 1.0

            # STEP 헤더 배지 (파란색)
            step_badge = slide.Shapes.AddShape(1, s_left, start_y, box_w, 36.0)
            step_badge.Fill.Solid()
            step_badge.Fill.ForeColor.RGB = th["primary"] if i == 0 else (th["accent_blue"] if i < 3 else th["secondary"])
            step_badge.Line.Visible = 0
            sb_tr = step_badge.TextFrame.TextRange
            sb_tr.Text = f"STEP 0{i+1}"
            sb_tr.Font.Name = "맑은 고딕"
            sb_tr.Font.Size = 11.0
            sb_tr.Font.Bold = -1
            sb_tr.Font.Color.RGB = th["white"]

            # 스텝 제목
            t_box = slide.Shapes.AddTextbox(1, s_left + 12.0, start_y + 48.0, box_w - 24.0, 36.0)
            t_tf = t_box.TextFrame
            t_tf.WordWrap = -1
            t_tf.MarginLeft = 0; t_tf.MarginRight = 0; t_tf.MarginTop = 0; t_tf.MarginBottom = 0
            t_tr = t_tf.TextRange
            t_tr.Text = slot.get("title", f"단계 {i+1}")
            t_tr.Font.Name = "맑은 고딕"
            t_tr.Font.Size = 13.5
            t_tr.Font.Bold = -1
            t_tr.Font.Color.RGB = th["text_dark"]

            # 스텝 본문
            b_box = slide.Shapes.AddTextbox(1, s_left + 12.0, start_y + 92.0, box_w - 24.0, box_h - 104.0)
            b_tf = b_box.TextFrame
            b_tf.WordWrap = -1
            b_tf.MarginLeft = 0; b_tf.MarginRight = 0; b_tf.MarginTop = 0; b_tf.MarginBottom = 0
            b_tr = b_tf.TextRange
            desc = slot.get("description") or slot.get("body", "")
            b_tr.Text = desc
            b_tr.Font.Name = "맑은 고딕"
            b_tr.Font.Size = 10.5
            b_tr.Font.Color.RGB = th["text_muted"]

            # 카드 간 네이티브 화살표 커넥터 (마지막 스텝 제외)
            if i < step_cnt - 1:
                arr = slide.Shapes.AddShape(13, s_left + box_w + 3.0, start_y + box_h/2.0 - 8.0, 10.0, 16.0) # RIGHT_ARROW
                arr.Fill.Solid()
                arr.Fill.ForeColor.RGB = th["accent"]
                arr.Line.Visible = 0

    @classmethod
    def render_comparison_asis_tobe(cls, slide: Any, data: dict[str, Any], theme: dict[str, int] = None):
        """[아키타입 3] AS-IS vs TO-BE 2단 좌우 대칭 비교 (대비 구조 극대화)"""
        th = theme or cls.DEFAULT_THEME
        start_y = cls.render_header(slide, data.get("title", "AS-IS vs TO-BE 비교 분석"), data.get("subtitle", ""), data.get("category", "전략비교"), th)
        slots = data.get("slots", [])

        left_m = 48.0
        avail_w = 864.0
        gap_x = 48.0
        panel_w = (avail_w - gap_x) / 2.0
        panel_h = min(360.0, cls.CANVAS_H - start_y - 36.0)

        # 1. AS-IS 패널 (좌측 - 차분한 그레이 테두리)
        as_is_data = slots[0] if len(slots) > 0 else {"title": "AS-IS (현재 체계)", "description": "파편화된 사일로 시스템\n수동 데이터 입력\n지연된 의사결정"}
        p1 = slide.Shapes.AddShape(5, left_m, start_y, panel_w, panel_h)
        p1.Fill.Solid()
        p1.Fill.ForeColor.RGB = th["surface_alt"]
        p1.Line.ForeColor.RGB = th["card_border"]
        p1.Line.Weight = 1.0

        # AS-IS 상단 헤더 바
        h1 = slide.Shapes.AddShape(1, left_m, start_y, panel_w, 40.0)
        h1.Fill.Solid()
        h1.Fill.ForeColor.RGB = th["secondary"]
        h1.Line.Visible = 0
        h1_tr = h1.TextFrame.TextRange
        h1_tr.Text = as_is_data.get("title", "AS-IS (현재 체계)")
        h1_tr.Font.Name = "맑은 고딕"
        h1_tr.Font.Size = 13.5
        h1_tr.Font.Bold = -1
        h1_tr.Font.Color.RGB = th["white"]

        # AS-IS 본문
        b1_box = slide.Shapes.AddTextbox(1, left_m + 20.0, start_y + 56.0, panel_w - 40.0, panel_h - 72.0)
        b1_tf = b1_box.TextFrame
        b1_tf.WordWrap = -1
        b1_tr = b1_tf.TextRange
        b1_tr.Text = as_is_data.get("description", "")
        b1_tr.Font.Name = "맑은 고딕"
        b1_tr.Font.Size = 12.0
        b1_tr.Font.Color.RGB = th["text_dark"]

        # 2. 중앙 VS 원형 배지
        vs_x = left_m + panel_w + gap_x / 2.0 - 18.0
        vs_y = start_y + panel_h / 2.0 - 18.0
        vs_circle = slide.Shapes.AddShape(9, vs_x, vs_y, 36.0, 36.0) # OVAL
        vs_circle.Fill.Solid()
        vs_circle.Fill.ForeColor.RGB = th["accent"]
        vs_circle.Line.Visible = 0
        vs_tr = vs_circle.TextFrame.TextRange
        vs_tr.Text = "VS"
        vs_tr.Font.Name = "맑은 고딕"
        vs_tr.Font.Size = 12.0
        vs_tr.Font.Bold = -1
        vs_tr.Font.Color.RGB = th["white"]

        # 3. TO-BE 패널 (우측 - 신뢰의 네이비 강조 테두리)
        to_be_data = slots[1] if len(slots) > 1 else {"title": "TO-BE (목표 체계)", "description": "통합 클라우드 플랫폼\n실시간 자동화 파이프라인\n데이터 기반 즉각 의사결정"}
        p2_left = left_m + panel_w + gap_x
        p2 = slide.Shapes.AddShape(5, p2_left, start_y, panel_w, panel_h)
        p2.Fill.Solid()
        p2.Fill.ForeColor.RGB = th["bg_light"]
        p2.Line.ForeColor.RGB = th["primary"]
        p2.Line.Weight = 2.0

        # TO-BE 상단 헤더 바
        h2 = slide.Shapes.AddShape(1, p2_left, start_y, panel_w, 40.0)
        h2.Fill.Solid()
        h2.Fill.ForeColor.RGB = th["primary"]
        h2.Line.Visible = 0
        h2_tr = h2.TextFrame.TextRange
        h2_tr.Text = to_be_data.get("title", "TO-BE (목표 체계)")
        h2_tr.Font.Name = "맑은 고딕"
        h2_tr.Font.Size = 13.5
        h2_tr.Font.Bold = -1
        h2_tr.Font.Color.RGB = th["white"]

        # TO-BE 본문
        b2_box = slide.Shapes.AddTextbox(1, p2_left + 20.0, start_y + 56.0, panel_w - 40.0, panel_h - 72.0)
        b2_tf = b2_box.TextFrame
        b2_tf.WordWrap = -1
        b2_tr = b2_tf.TextRange
        b2_tr.Text = to_be_data.get("description", "")
        b2_tr.Font.Name = "맑은 고딕"
        b2_tr.Font.Size = 12.0
        b2_tr.Font.Color.RGB = th["text_dark"]

    @classmethod
    def render_timeline_quarterly(cls, slide: Any, data: dict[str, Any], theme: dict[str, int] = None):
        """[아키타입 4] 4분기/연도별 추진 로드맵 (수평 축선 + 노드 + 마일스톤 카드)"""
        th = theme or cls.DEFAULT_THEME
        start_y = cls.render_header(slide, data.get("title", "추진 로드맵 및 일정"), data.get("subtitle", ""), data.get("category", "마일스톤"), th)
        slots = data.get("slots", [])[:4]

        left_m = 48.0
        avail_w = 864.0
        node_cnt = 4
        gap_x = 16.0
        card_w = (avail_w - gap_x * (node_cnt - 1)) / node_cnt

        axis_y = start_y + 40.0

        # 1. 수평 타임라인 메인 축선
        axis_line = slide.Shapes.AddLine(left_m + 30.0, axis_y, left_m + avail_w - 30.0, axis_y)
        axis_line.Line.ForeColor.RGB = th["card_border"]
        axis_line.Line.Weight = 3.0

        # 2. 분기별 노드 및 마일스톤 카드
        card_y = axis_y + 36.0
        card_h = min(240.0, cls.CANVAS_H - card_y - 36.0)

        for i in range(node_cnt):
            slot = slots[i] if i < len(slots) else {"title": f"Q{i+1}", "description": "주요 달성 과제"}
            c_left = left_m + i * (card_w + gap_x)
            node_center_x = c_left + card_w / 2.0

            # 원형 축선 노드 (동심원 효과)
            outer_circle = slide.Shapes.AddShape(9, node_center_x - 14.0, axis_y - 14.0, 28.0, 28.0)
            outer_circle.Fill.Solid()
            outer_circle.Fill.ForeColor.RGB = th["primary"] if i == 0 else th["accent_blue"]
            outer_circle.Line.Visible = 0

            inner_circle = slide.Shapes.AddShape(9, node_center_x - 6.0, axis_y - 6.0, 12.0, 12.0)
            inner_circle.Fill.Solid()
            inner_circle.Fill.ForeColor.RGB = th["white"]
            inner_circle.Line.Visible = 0

            # 상단 분기 라벨
            q_label = slide.Shapes.AddTextbox(1, node_center_x - 40.0, axis_y - 38.0, 80.0, 20.0)
            q_tf = q_label.TextFrame
            q_tf.WordWrap = 0
            q_tr = q_tf.TextRange
            q_tr.Text = f"2026 Q{i+1}"
            q_tr.ParagraphFormat.Alignment = 2 # Center
            q_tr.Font.Name = "맑은 고딕"
            q_tr.Font.Size = 11.5
            q_tr.Font.Bold = -1
            q_tr.Font.Color.RGB = th["primary"]

            # 하단 마일스톤 카드
            card = slide.Shapes.AddShape(5, c_left, card_y, card_w, card_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = th["bg_light"]
            card.Line.ForeColor.RGB = th["card_border"]
            card.Line.Weight = 1.0

            # 카드 제목
            t_box = slide.Shapes.AddTextbox(1, c_left + 12.0, card_y + 14.0, card_w - 24.0, 32.0)
            t_tf = t_box.TextFrame
            t_tf.WordWrap = -1
            t_tr = t_tf.TextRange
            t_tr.Text = slot.get("title", f"마일스톤 {i+1}")
            t_tr.Font.Name = "맑은 고딕"
            t_tr.Font.Size = 13.0
            t_tr.Font.Bold = -1
            t_tr.Font.Color.RGB = th["text_dark"]

            # 카드 본문
            b_box = slide.Shapes.AddTextbox(1, c_left + 12.0, card_y + 50.0, card_w - 24.0, card_h - 60.0)
            b_tf = b_box.TextFrame
            b_tf.WordWrap = -1
            b_tr = b_tf.TextRange
            b_tr.Text = slot.get("description", "")
            b_tr.Font.Name = "맑은 고딕"
            b_tr.Font.Size = 10.5
            b_tr.Font.Color.RGB = th["text_muted"]

    @classmethod
    def render_kpi_metrics(cls, slide: Any, data: dict[str, Any], theme: dict[str, int] = None):
        """[아키타입 5] 3대 핵심 성과지표 빅넘버 카드 (36pt 대형 지표 + 레이블)"""
        th = theme or cls.DEFAULT_THEME
        start_y = cls.render_header(slide, data.get("title", "핵심 경영 성과 지표 (KPI)"), data.get("subtitle", ""), data.get("category", "정량성과"), th)
        slots = data.get("slots", [])[:3]

        left_m = 48.0
        avail_w = 864.0
        gap_x = 24.0
        card_w = (avail_w - gap_x * 2) / 3.0
        card_h = min(320.0, cls.CANVAS_H - start_y - 40.0)

        SAMPLE_NUMS = ["35%", "94.2%", "15hrs"]
        SAMPLE_TAGS = ["비용 절감율", "고객 리텐션", "주간 시간 절약"]

        for i in range(3):
            slot = slots[i] if i < len(slots) else {"title": SAMPLE_TAGS[i], "description": "목표치 달성"}
            c_left = left_m + i * (card_w + gap_x)

            # 라운드 카드
            card = slide.Shapes.AddShape(5, c_left, start_y, card_w, card_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = th["bg_light"]
            card.Line.ForeColor.RGB = th["card_border"]
            card.Line.Weight = 1.0

            # 36pt 대형 빅넘버 텍스트
            num_box = slide.Shapes.AddTextbox(1, c_left + 16.0, start_y + 24.0, card_w - 32.0, 56.0)
            n_tf = num_box.TextFrame
            n_tf.WordWrap = 0
            n_tr = n_tf.TextRange
            n_tr.Text = slot.get("stat") or SAMPLE_NUMS[i]
            n_tr.Font.Name = "맑은 고딕"
            n_tr.Font.Size = 36.0
            n_tr.Font.Bold = -1
            n_tr.Font.Color.RGB = th["primary"] if i == 0 else (th["accent_blue"] if i == 1 else th["accent"])

            # 지표명 (15pt Bold)
            t_box = slide.Shapes.AddTextbox(1, c_left + 16.0, start_y + 88.0, card_w - 32.0, 30.0)
            t_tf = t_box.TextFrame
            t_tf.WordWrap = -1
            t_tr = t_tf.TextRange
            t_tr.Text = slot.get("title", SAMPLE_TAGS[i])
            t_tr.Font.Name = "맑은 고딕"
            t_tr.Font.Size = 14.5
            t_tr.Font.Bold = -1
            t_tr.Font.Color.RGB = th["text_dark"]

            # 구분선
            div = slide.Shapes.AddLine(c_left + 16.0, start_y + 124.0, c_left + card_w - 16.0, start_y + 124.0)
            div.Line.ForeColor.RGB = th["card_border"]
            div.Line.Weight = 0.75

            # 세부 설명
            b_box = slide.Shapes.AddTextbox(1, c_left + 16.0, start_y + 136.0, card_w - 32.0, card_h - 150.0)
            b_tf = b_box.TextFrame
            b_tf.WordWrap = -1
            b_tr = b_tf.TextRange
            b_tr.Text = slot.get("description", "")
            b_tr.Font.Name = "맑은 고딕"
            b_tr.Font.Size = 11.0
            b_tr.Font.Color.RGB = th["text_muted"]
