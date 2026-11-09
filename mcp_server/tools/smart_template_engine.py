# -*- coding: utf-8 -*-
"""
mcp_server/tools/smart_template_engine.py
'06_트렌디_비즈니스_템플릿' 실물 슬라이드를 InsertFromFile로 무손실 복제하고,
원본 폰트/스타일/자간을 100% 보존하면서 사용자 비즈니스 텍스트만 인플레이스로 치환하는
차세대 스마트 템플릿 카탈로그 엔진(STCE).
"""

from __future__ import annotations
import os
import json
import win32com.client
from typing import Any
from pathlib import Path

CATALOG_PATH = Path("data/trendy_smart_catalog.json")


class SmartTemplateEngine:
    _CATALOG_CACHE: dict[str, Any] = {}

    @classmethod
    def get_catalog(cls) -> dict[str, Any]:
        """스마트 카탈로그 메타데이터 로드"""
        if not cls._CATALOG_CACHE:
            if CATALOG_PATH.exists():
                with open(CATALOG_PATH, "r", encoding="utf-8") as f:
                    cls._CATALOG_CACHE = json.load(f)
        return cls._CATALOG_CACHE

    @classmethod
    def get_entry(cls, catalog_id: str) -> dict[str, Any] | None:
        """catalog_id로 카탈로그 항목 조회"""
        cat = cls.get_catalog()
        return cat.get(catalog_id)

    @classmethod
    def list_templates(cls, category: str | None = None) -> list[dict[str, Any]]:
        """카테고리별 스마트 템플릿 목록 조회"""
        cat = cls.get_catalog()
        res = []
        for k, v in cat.items():
            if not category or v.get("category") == category:
                res.append(v)
        return res

    @classmethod
    def clone_and_replace(
        cls,
        prs: Any,
        catalog_id: str,
        data: dict[str, Any],
        insert_index: int = 1
    ) -> Any:
        """
        지정된 PowerPoint Presentation(prs)에 template 슬라이드를 InsertFromFile로 복제하고,
        사용자 데이터를 인플레이스로 정밀 치환하여 반환.
        """
        entry = cls.get_entry(catalog_id)
        if not entry:
            raise ValueError(f"SmartTemplateEngine: 알 수 없는 catalog_id '{catalog_id}'")

        template_file = os.path.abspath(entry["source_file"])
        slide_idx = entry["slide_index"]  # 1-based index for PowerPoint COM

        if not os.path.exists(template_file):
            raise FileNotFoundError(f"템플릿 파일이 디스크에 없습니다: {template_file}")

        # 1. PowerPoint COM InsertFromFile 호출 (무손실 복제)
        current_count = prs.Slides.Count
        prs.Slides.InsertFromFile(template_file, insert_index - 1, slide_idx, slide_idx)
        new_count = prs.Slides.Count

        if new_count <= current_count:
            raise RuntimeError(f"슬라이드 복제 실패: {template_file} [Slide #{slide_idx}]")

        slide = prs.Slides(insert_index)

        # 2. 텍스트 치환 데이터 준비
        title = str(data.get("title", "")).strip()
        subtitle = str(data.get("subtitle", "")).strip()
        slots = data.get("slots") or data.get("items") or []

        # 표준 플레이스홀더 키워드 패턴
        PLACEHOLDER_KEYWORDS = (
            "제목을 입력하세요", "내용을 입력하세요", "텍스트를 입력하세요", "텍스트를입력하세요",
            "내용을입력하세요", "제목을입력하세요", "소제목", "설명", "part 1", "part 2", "part 3",
            "a table of contents", "목차"
        )

        # 3. 슬라이드 내 모든 텍스트 상자 수집
        text_shapes = []
        for s in slide.Shapes:
            if s.HasTextFrame and s.TextFrame.HasText:
                txt = s.TextFrame.TextRange.Text.strip()
                # 워터마크나 카피라이트는 절대 건드리지 않음
                if any(w in txt.lower() for w in ("copyright", "all rights reserved", "©", "ⓒ")):
                    continue
                text_shapes.append((s.Top, s.Left, s, txt))

        # 위치 순 정렬 (상->하, 좌->우)
        text_shapes.sort(key=lambda x: (x[0], x[1]))

        # 4. 정밀 슬롯 바인딩 및 텍스트 겹침 원천 차단 (Precision Slot Binding)
        slot_idx = 0
        title_done = False
        subtitle_done = False

        for top, left, sh, cur_txt in text_shapes:
            cur_clean = cur_txt.replace(" ", "").lower()
            h_pt = sh.Height
            w_pt = sh.Width

            # (1) 헤더 메인 타이틀 판별 (상단 110pt 이내, 큰 폰트 또는 '목차', 'part', 'title', 'viva')
            if not title_done and title:
                is_title_target = any(k in cur_clean for k in ("목차", "title", "part", "viva", "magenta", "contents")) or (top < 110.0 and h_pt < 45.0)
                if is_title_target:
                    # 1줄 제목만 정밀 주입
                    cls._replace_text_preserve_style(sh, title, is_single_line=True)
                    title_done = True
                    continue

            # (2) 헤더 서브타이틀 판별 (상단 140pt 이내, 영문 소제목 등)
            if not subtitle_done and top < 140.0:
                is_sub_target = any(k in cur_clean for k in ("atableofcontents", "소제목", "subtitle", "contents"))
                if is_sub_target:
                    if subtitle:
                        cls._replace_text_preserve_style(sh, subtitle, is_single_line=True)
                    # subtitle이 없으면 원본의 세련된 영문 데코레이션(a table of contents) 보존
                    subtitle_done = True
                    continue

            # (3) 슬롯 아이템 치환 (번호 '1', '2', '3' 등은 디자이너 원본 유지)
            if cur_txt.isdigit() and len(cur_txt) <= 2:
                continue

            # 플레이스홀더 텍스트 또는 일반 본문 영역
            is_placeholder = any(k in cur_clean for k in ("내용을입력하세요", "제목을입력하세요", "텍스트를입력하세요", "그림", "항목", "요소", "단계")) or len(cur_txt) >= 4

            if is_placeholder:
                if slot_idx < len(slots):
                    slot_item = slots[slot_idx]
                    s_title = ""
                    s_desc = ""
                    if isinstance(slot_item, dict):
                        s_title = slot_item.get("title", "")
                        s_desc = slot_item.get("description") or slot_item.get("body") or slot_item.get("points", "")
                        if isinstance(s_desc, list):
                            s_desc = " ".join(str(p) for p in s_desc)
                    else:
                        s_title = str(slot_item)

                    # 텍스트 박스 높이에 따른 스마트 분기
                    if h_pt < 45.0:
                        # 1줄짜리 슬롯 제목 전용 박스 -> 절대로 줄바꿈을 넣지 않고 제목만 주입!
                        inject = s_title if s_title else s_desc
                        cls._replace_text_preserve_style(sh, inject, is_single_line=True)
                    else:
                        # 다단 본문 설명 박스 -> 설명문 우선 주입
                        inject = s_desc if s_desc else s_title
                        cls._replace_text_preserve_style(sh, inject, is_single_line=False)

                    slot_idx += 1
                else:
                    # 더 이상 주입할 슬롯 데이터가 없는데 기본 안내 문구인 경우 -> 깨끗하게 공백 제거
                    if any(k in cur_clean for k in ("내용을입력하세요", "제목을입력하세요", "텍스트를입력하세요")):
                        sh.TextFrame.TextRange.Text = ""

        return slide

    @classmethod
    def _replace_text_preserve_style(cls, shape: Any, new_text: str, is_single_line: bool = False):
        """Shape의 원본 서식(Font, Size, Color, Alignment)을 보존하면서 텍스트만 치환"""
        tf = shape.TextFrame
        tr = tf.TextRange

        # 기존 스타일 속성 백업
        f_name = "맑은 고딕"
        f_size = 14.0
        f_bold = 0
        f_color = 0x222222
        align = 1

        try:
            f_name = tr.Font.Name
            f_size = tr.Font.Size
            f_bold = tr.Font.Bold
            f_color = tr.Font.Color.RGB
            align = tr.ParagraphFormat.Alignment
        except Exception:
            pass

        # 1줄 전용일 경우 개행문자 완전 제거 및 자동 너비 확보
        if is_single_line:
            new_text = " ".join(new_text.split())
            tf.WordWrap = 0  # 1줄 전용은 강제 줄바꿈 방지
            tf.MarginLeft = 0.0
            tf.MarginRight = 0.0
            # 텍스트 예상 폭 계산 및 shape 너비 안전 확장 (우측 침범 방지 한도 940pt)
            est_width = len(new_text) * (f_size * 1.1) + 20.0
            if shape.Width < est_width and (shape.Left + est_width) < 940.0:
                shape.Width = est_width
        else:
            tf.WordWrap = -1  # 다단 본문은 줄바꿈 허용

        # 텍스트 변경
        tr.Text = new_text

        # 텍스트 박스 자동 축소 (msoAutoSizeTextToFit = 2) 시도
        try:
            if hasattr(shape, "TextFrame2"):
                shape.TextFrame2.AutoSize = 2  # msoAutoSizeTextToFit
                shape.TextFrame2.WordWrap = -1
        except Exception:
            pass

        # 백업한 스타일 재적용
        try:
            tr.Font.Name = f_name
            # 텍스트 길이가 과도하게 길 경우에만 안전한 폰트 스케일링
            if len(new_text) > 80 and f_size > 11.0:
                tr.Font.Size = max(9.5, f_size * 0.85)
            elif len(new_text) > 50 and f_size > 13.0:
                tr.Font.Size = max(10.5, f_size * 0.9)
            elif is_single_line and len(new_text) > 20 and f_size > 14.0:
                tr.Font.Size = max(12.0, f_size * 0.9)
            else:
                tr.Font.Size = f_size

            tr.Font.Bold = f_bold
            tr.Font.Color.RGB = f_color
            tr.ParagraphFormat.Alignment = align
        except Exception:
            pass
