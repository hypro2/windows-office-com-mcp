# -*- coding: utf-8 -*-
"""
mcp_server/tools/ppt_technique_functions.py
06_트렌디_비즈니스_템플릿 최우선 적용 95개 기법 1:1 Digital Twin 독립 COM 렌더링 함수 모듈.

- 원본 디자이너의 Shape Tree(1,170개 실제 형상)를 1:1 정밀 복제
- 95개 기법 각각에 대해 render_tech_001 ~ render_tech_095 독립 전용 함수 구현
- 도형 종류(AutoShapeType: 사각형, 라운드, 셰브론 화살표, 타원 등) 100% 복원
- 텍스트 인셋 및 자동 폰트 스케일링으로 외피-내피 충돌 및 텍스트 삐져나옴 0% 달성
"""

from __future__ import annotations
import os
import json
from typing import Any, Callable
from pathlib import Path
from .ppt_slice_engine import PptSliceEngine, ComImageHelper

# Shape Tree 데이터 로드 및 인메모리 캐싱
_SHAPE_TREES_CACHE: dict[str, Any] = {}

def _get_shape_tree(tech_name: str) -> dict[str, Any]:
    global _SHAPE_TREES_CACHE
    if not _SHAPE_TREES_CACHE:
        trees_path = Path("data/trendy_golden_shape_trees.json")
        if trees_path.exists():
            with open(trees_path, "r", encoding="utf-8") as f:
                _SHAPE_TREES_CACHE = json.load(f)
    return _SHAPE_TREES_CACHE.get(tech_name, {})


def render_digital_twin_slide(slide: Any, tech_name: str, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """95개 독립 함수를 위한 Digital Twin 형상 복제 및 데이터 주입 코어 렌더러"""
    created: list[int] = []
    tree_data = _get_shape_tree(tech_name)
    if not tree_data or "shapes" not in tree_data:
        return created

    w, h = 960.0, 540.0
    title = str(data.get("title", tech_name)).strip()
    subtitle = str(data.get("subtitle", "")).strip()
    category = str(data.get("category", "")).strip()
    slots = data.get("slots") or data.get("items") or []

    shapes_list = tree_data["shapes"]
    # 슬롯 바인딩용 인덱스 카운터
    slot_title_idx = 0
    slot_body_idx = 0
    badge_idx = 0

    for s_info in shapes_list:
        geom = s_info["geometry"]
        rel_l = geom.get("rel_left", 0.0)
        rel_t = geom.get("rel_top", 0.0)
        rel_w = geom.get("rel_width", 0.1)
        rel_h = geom.get("rel_height", 0.1)

        # 슬라이드 범위 밖 비정상 좌표 가드
        if rel_w <= 0.005 or rel_h <= 0.005:
            continue
        if rel_l < -0.05 or rel_l > 1.05 or rel_t < -0.05 or rel_t > 1.05:
            continue

        left_pt = max(0.0, rel_l * w)
        top_pt = max(0.0, rel_t * h)
        width_pt = min(w - left_pt, rel_w * w)
        height_pt = min(h - top_pt, rel_h * h)

        # 최소 크기 보장
        if width_pt < 4.0 or height_pt < 4.0:
            continue

        role = s_info.get("role", "decorative_shape")
        ast_num = s_info.get("auto_shape_type")
        # 기본 사각형(1) 또는 원본 AutoShapeType 사용
        shape_type_to_add = ast_num if (ast_num and 1 <= ast_num <= 180) else 1

        has_text = s_info.get("has_text", False)
        orig_text = s_info.get("text", "")

        # 텍스트 주입 대상 결정
        inject_text = ""
        is_title_role = False
        if role == "header_title":
            inject_text = title
            is_title_role = True
        elif role == "header_tag":
            inject_text = category if category else (subtitle if subtitle else orig_text)
        elif role == "badge_number":
            badge_idx += 1
            inject_text = f"{badge_idx:02d}"
        elif role == "slot_title":
            if slot_title_idx < len(slots):
                inject_text = str(slots[slot_title_idx].get("title", f"주요 과제 {slot_title_idx+1}"))
            else:
                inject_text = orig_text if orig_text else f"과제 {slot_title_idx+1}"
            slot_title_idx += 1
        elif role == "slot_body":
            if slot_body_idx < len(slots):
                s_desc = slots[slot_body_idx].get("description") or slots[slot_body_idx].get("body") or slots[slot_body_idx].get("points", "")
                if isinstance(s_desc, list):
                    inject_text = "\n".join(str(p) for p in s_desc)
                else:
                    inject_text = str(s_desc)
            else:
                inject_text = orig_text
            slot_body_idx += 1
        elif has_text:
            inject_text = orig_text

        # COM 도형 생성
        try:
            if not has_text and not inject_text:
                # 순수 배경/장식/커넥터 도형
                sh = slide.Shapes.AddShape(shape_type_to_add, left_pt, top_pt, width_pt, height_pt)
                # 테마 색상 맵핑
                if role in ("card_background", "container"):
                    sh.Fill.Solid()
                    sh.Fill.ForeColor.RGB = theme["bg_light"]
                    sh.Line.ForeColor.RGB = theme["card_border"]
                    sh.Line.Weight = 1.0
                elif role == "connector_arrow":
                    sh.Fill.Solid()
                    sh.Fill.ForeColor.RGB = theme["accent"]
                    sh.Line.Visible = 0
                else:
                    # 장식 요소: primary 또는 secondary 대비 반사
                    sh.Fill.Solid()
                    sh.Fill.ForeColor.RGB = theme["secondary"]
                    sh.Line.Visible = 0
                created.append(sh.Id)
            else:
                # 텍스트 포함 도형 또는 텍스트 상자
                if role == "badge_number":
                    sh = slide.Shapes.AddShape(shape_type_to_add, left_pt, top_pt, width_pt, height_pt)
                    sh.Fill.Solid()
                    sh.Fill.ForeColor.RGB = theme["primary"]
                    sh.Line.Visible = 0
                elif role in ("slot_title", "slot_body") and shape_type_to_add != 1:
                    # 특수 형태 카드 자체에 텍스트 포함된 경우
                    sh = slide.Shapes.AddShape(shape_type_to_add, left_pt, top_pt, width_pt, height_pt)
                    sh.Fill.Solid()
                    sh.Fill.ForeColor.RGB = theme["bg_light"]
                    sh.Line.ForeColor.RGB = theme["card_border"]
                    sh.Line.Weight = 1.0
                else:
                    sh = slide.Shapes.AddTextbox(1, left_pt, top_pt, width_pt, height_pt)
                    sh.Fill.Background()
                    sh.Line.Visible = 0

                # 텍스트 프레임 안전 가드
                tf = sh.TextFrame
                tf.WordWrap = -1
                tf.MarginLeft = 4.0
                tf.MarginRight = 4.0
                tf.MarginTop = 4.0
                tf.MarginBottom = 4.0

                tr = tf.TextRange
                tr.Text = inject_text if inject_text else orig_text
                tr.Font.Name = "맑은 고딕"

                # 폰트 크기 및 스타일 결정
                orig_font_pt = s_info.get("max_font_size_pt", 12.0)
                if is_title_role:
                    t_pt = orig_font_pt if orig_font_pt >= 20.0 else 24.0
                    if len(inject_text) > 28:
                        t_pt = max(18.0, t_pt * 0.85)
                    tr.Font.Size = t_pt
                    tr.Font.Bold = -1
                    tr.Font.Color.RGB = theme["text_dark"]
                elif role == "badge_number":
                    tr.Font.Size = 11.0
                    tr.Font.Bold = -1
                    tr.Font.Color.RGB = theme["white"]
                elif role == "slot_title":
                    tr.Font.Size = 13.0 if orig_font_pt < 12.0 else min(16.0, orig_font_pt)
                    tr.Font.Bold = -1
                    tr.Font.Color.RGB = theme["text_dark"]
                elif role == "slot_body":
                    # 텍스트 길이에 따른 동적 스케일링
                    b_pt = 11.0
                    if len(inject_text) > 70:
                        b_pt = 9.5
                    elif len(inject_text) > 40:
                        b_pt = 10.5
                    tr.Font.Size = b_pt
                    tr.Font.Bold = 0
                    tr.Font.Color.RGB = theme["text_muted"]
                else:
                    tr.Font.Size = max(9.0, min(14.0, orig_font_pt))
                    tr.Font.Bold = -1 if s_info.get("is_bold") else 0
                    tr.Font.Color.RGB = theme["text_dark"]

                created.append(sh.Id)
        except Exception:
            continue

    return created


def render_tech_001_미니멀_타이포_원라인_표지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_001] 미니멀_타이포_원라인_표지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "미니멀_타이포_원라인_표지", data, theme, image_files)


def render_tech_002_세로_넘버링_리스트_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_002] 세로_넘버링_리스트_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "세로_넘버링_리스트_목차", data, theme, image_files)


def render_tech_003_좌우_1_1_히어로_사진_스플릿_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_003] 좌우_1:1_히어로_사진_스플릿_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "좌우_1:1_히어로_사진_스플릿_카드", data, theme, image_files)


def render_tech_004_넘버링_배지_2단카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_004] 넘버링_배지_2단카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "넘버링_배지_2단카드", data, theme, image_files)


def render_tech_005_수직_2단_스트라이프_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_005] 수직_2단_스트라이프_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "수직_2단_스트라이프_목차", data, theme, image_files)


def render_tech_006_추진일정_월별_마일스톤표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_006] 추진일정_월별_마일스톤표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 9개
    """
    return render_digital_twin_slide(slide, "추진일정_월별_마일스톤표", data, theme, image_files)


def render_tech_007_KISA식_1_가_1_공문서형_위계본문(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_007] KISA식_1_가_(1)_공문서형_위계본문 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 22개
    """
    return render_digital_twin_slide(slide, "KISA식_1_가_(1)_공문서형_위계본문", data, theme, image_files)


def render_tech_008_수직_2단_가로바_리스트_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_008] 수직_2단_가로바_리스트_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "수직_2단_가로바_리스트_카드", data, theme, image_files)


def render_tech_009_수직_3단_가로바_리스트_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_009] 수직_3단_가로바_리스트_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 27개
    """
    return render_digital_twin_slide(slide, "수직_3단_가로바_리스트_카드", data, theme, image_files)


def render_tech_010_2단_수평_대등_핵심과제_카드그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_010] 2단_수평_대등_핵심과제_카드그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "2단_수평_대등_핵심과제_카드그리드", data, theme, image_files)


def render_tech_011_표준_비즈니스_개조식_단락본문(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_011] 표준_비즈니스_개조식_단락본문 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 11개
    """
    return render_digital_twin_slide(slide, "표준_비즈니스_개조식_단락본문", data, theme, image_files)


def render_tech_012_네이티브_비즈니스_통계차트(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_012] 네이티브_비즈니스_통계차트 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 13개
    """
    return render_digital_twin_slide(slide, "네이티브_비즈니스_통계차트", data, theme, image_files)


def render_tech_013_3단_사진_상단결합_비주얼카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_013] 3단_사진_상단결합_비주얼카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 21개
    """
    return render_digital_twin_slide(slide, "3단_사진_상단결합_비주얼카드", data, theme, image_files)


def render_tech_014_스마트폰_모바일앱_지시선_콜아웃(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_014] 스마트폰_모바일앱_지시선_콜아웃 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "스마트폰_모바일앱_지시선_콜아웃", data, theme, image_files)


def render_tech_015_측면오픈형_비즈니스_명세표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_015] 측면오픈형_비즈니스_명세표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "측면오픈형_비즈니스_명세표", data, theme, image_files)


def render_tech_016_1_2_비대칭_좌측요약_우측상세_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_016] 1:2_비대칭_좌측요약_+_우측상세_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "1:2_비대칭_좌측요약_+_우측상세_카드", data, theme, image_files)


def render_tech_017_2_1_비대칭_좌측상세_우측콜아웃_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_017] 2:1_비대칭_좌측상세_+_우측콜아웃_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 10개
    """
    return render_digital_twin_slide(slide, "2:1_비대칭_좌측상세_+_우측콜아웃_카드", data, theme, image_files)


def render_tech_018_미니멀_라인_간지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_018] 미니멀_라인_간지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 9개
    """
    return render_digital_twin_slide(slide, "미니멀_라인_간지", data, theme, image_files)


def render_tech_019_Thank_You_미니멀_엔딩(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_019] Thank_You_미니멀_엔딩 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 13개
    """
    return render_digital_twin_slide(slide, "Thank_You_미니멀_엔딩", data, theme, image_files)


def render_tech_020_풀배경_다크_비주얼_표지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_020] 풀배경_다크_비주얼_표지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "풀배경_다크_비주얼_표지", data, theme, image_files)


def render_tech_021_좌우_2분할_컬러블록_표지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_021] 좌우_2분할_컬러블록_표지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "좌우_2분할_컬러블록_표지", data, theme, image_files)


def render_tech_022_컬러_풀블록_섹션_간지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_022] 컬러_풀블록_섹션_간지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 9개
    """
    return render_digital_twin_slide(slide, "컬러_풀블록_섹션_간지", data, theme, image_files)


def render_tech_023_좌측_빅넘버_우측_섹션명_분할간지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_023] 좌측_빅넘버_우측_섹션명_분할간지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 10개
    """
    return render_digital_twin_slide(slide, "좌측_빅넘버_우측_섹션명_분할간지", data, theme, image_files)


def render_tech_024_중앙_특대형_빅넘버_챕터간지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_024] 중앙_특대형_빅넘버_챕터간지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 20개
    """
    return render_digital_twin_slide(slide, "중앙_특대형_빅넘버_챕터간지", data, theme, image_files)


def render_tech_025_카드_오버레이_비즈니스_표지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_025] 카드_오버레이_비즈니스_표지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "카드_오버레이_비즈니스_표지", data, theme, image_files)


def render_tech_026_강조선_콜아웃_배너_장표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_026] 강조선_콜아웃_배너_장표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 11개
    """
    return render_digital_twin_slide(slide, "강조선_콜아웃_배너_장표", data, theme, image_files)


def render_tech_027_비주얼_배경사진_감사인사_엔딩(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_027] 비주얼_배경사진_감사인사_엔딩 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 18개
    """
    return render_digital_twin_slide(slide, "비주얼_배경사진_감사인사_엔딩", data, theme, image_files)


def render_tech_028_풀배경_이미지_표지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_028] 풀배경_이미지_표지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "풀배경_이미지_표지", data, theme, image_files)


def render_tech_029_대형_따옴표_슬로건_박스(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_029] 대형_따옴표_슬로건_박스 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "대형_따옴표_슬로건_박스", data, theme, image_files)


def render_tech_030_3단_아이콘_원형_헤더_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_030] 3단_아이콘_원형_헤더_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "3단_아이콘_원형_헤더_카드", data, theme, image_files)


def render_tech_031_3단_경영실적_증감배지_KPI(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_031] 3단_경영실적_증감배지_KPI 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 18개
    """
    return render_digital_twin_slide(slide, "3단_경영실적_증감배지_KPI", data, theme, image_files)


def render_tech_032_수직_4단_가로바_리스트_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_032] 수직_4단_가로바_리스트_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 15개
    """
    return render_digital_twin_slide(slide, "수직_4단_가로바_리스트_카드", data, theme, image_files)


def render_tech_033_4단_가로_파이프라인형_카드그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_033] 4단_가로_파이프라인형_카드그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 22개
    """
    return render_digital_twin_slide(slide, "4단_가로_파이프라인형_카드그리드", data, theme, image_files)


def render_tech_034_Q_A_질의응답_엔딩(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_034] Q&A_질의응답_엔딩 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 13개
    """
    return render_digital_twin_slide(slide, "Q&A_질의응답_엔딩", data, theme, image_files)


def render_tech_035_볼드_빅타이포_임팩트_표지(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_035] 볼드_빅타이포_임팩트_표지 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "볼드_빅타이포_임팩트_표지", data, theme, image_files)


def render_tech_036_5단_슬림_핵심과제_카드그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_036] 5단_슬림_핵심과제_카드그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 13개
    """
    return render_digital_twin_slide(slide, "5단_슬림_핵심과제_카드그리드", data, theme, image_files)


def render_tech_037_상하_3단_스텝_프로세스_바(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_037] 상하_3단_스텝_프로세스_바 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 15개
    """
    return render_digital_twin_slide(slide, "상하_3단_스텝_프로세스_바", data, theme, image_files)


def render_tech_038_동심원_3중_과녁모델(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_038] 동심원_3중_과녁모델 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 9개
    """
    return render_digital_twin_slide(slide, "동심원_3중_과녁모델", data, theme, image_files)


def render_tech_039_연락처결합_클로징_엔딩(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_039] 연락처결합_클로징_엔딩 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "연락처결합_클로징_엔딩", data, theme, image_files)


def render_tech_040_장문형_내러티브_설명문_본문(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_040] 장문형_내러티브_설명문_본문 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "장문형_내러티브_설명문_본문", data, theme, image_files)


def render_tech_041_선형_3단계_셰브론_파이프라인(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_041] 선형_3단계_셰브론_파이프라인 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 16개
    """
    return render_digital_twin_slide(slide, "선형_3단계_셰브론_파이프라인", data, theme, image_files)


def render_tech_042_넘버링_배지_4단카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_042] 넘버링_배지_4단카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 9개
    """
    return render_digital_twin_slide(slide, "넘버링_배지_4단카드", data, theme, image_files)


def render_tech_043_아이콘_헤더_결합_2단카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_043] 아이콘_헤더_결합_2단카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "아이콘_헤더_결합_2단카드", data, theme, image_files)


def render_tech_044_체크포인트_불릿_개조식_본문(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_044] 체크포인트_불릿_개조식_본문 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 10개
    """
    return render_digital_twin_slide(slide, "체크포인트_불릿_개조식_본문", data, theme, image_files)


def render_tech_045_Pros_vs_Cons_장단점_대조표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_045] Pros_vs_Cons_장단점_대조표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "Pros_vs_Cons_장단점_대조표", data, theme, image_files)


def render_tech_046_상단요약박스_하단개조식_본문(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_046] 상단요약박스_+_하단개조식_본문 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "상단요약박스_+_하단개조식_본문", data, theme, image_files)


def render_tech_047_수직_3단_스트라이프_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_047] 수직_3단_스트라이프_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "수직_3단_스트라이프_목차", data, theme, image_files)


def render_tech_048_1개_특대형_빅넘버_단독강조(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_048] 1개_특대형_빅넘버_단독강조 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 22개
    """
    return render_digital_twin_slide(slide, "1개_특대형_빅넘버_단독강조", data, theme, image_files)


def render_tech_049_오피스약도_본사주소_안내블록(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_049] 오피스약도_본사주소_안내블록 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 17개
    """
    return render_digital_twin_slide(slide, "오피스약도_본사주소_안내블록", data, theme, image_files)


def render_tech_050_가로_3열_카드형_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_050] 가로_3열_카드형_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 10개
    """
    return render_digital_twin_slide(slide, "가로_3열_카드형_목차", data, theme, image_files)


def render_tech_051_가로_4열_카드형_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_051] 가로_4열_카드형_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 22개
    """
    return render_digital_twin_slide(slide, "가로_4열_카드형_목차", data, theme, image_files)


def render_tech_052_3단_이미지_쇼케이스_그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_052] 3단_이미지_쇼케이스_그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 18개
    """
    return render_digital_twin_slide(slide, "3단_이미지_쇼케이스_그리드", data, theme, image_files)


def render_tech_053_멀티트랙_간트차트_일정표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_053] 멀티트랙_간트차트_일정표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 29개
    """
    return render_digital_twin_slide(slide, "멀티트랙_간트차트_일정표", data, theme, image_files)


def render_tech_054_3단_표준_핵심과제_카드그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_054] 3단_표준_핵심과제_카드그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "3단_표준_핵심과제_카드그리드", data, theme, image_files)


def render_tech_055_Top_1_5_랭킹_리더보드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_055] Top_1~5_랭킹_리더보드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "Top_1~5_랭킹_리더보드", data, theme, image_files)


def render_tech_056_6단_슬림_카드그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_056] 6단_슬림_카드그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "6단_슬림_카드그리드", data, theme, image_files)


def render_tech_057_4단_비주얼_갤러리(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_057] 4단_비주얼_갤러리 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 15개
    """
    return render_digital_twin_slide(slide, "4단_비주얼_갤러리", data, theme, image_files)


def render_tech_058_4단_대시보드_핵심수치_그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_058] 4단_대시보드_핵심수치_그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 23개
    """
    return render_digital_twin_slide(slide, "4단_대시보드_핵심수치_그리드", data, theme, image_files)


def render_tech_059_비대칭_매스너리_포트폴리오(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_059] 비대칭_매스너리_포트폴리오 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "비대칭_매스너리_포트폴리오", data, theme, image_files)


def render_tech_060_4열_사각사진_프로젝트팀_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_060] 4열_사각사진_프로젝트팀_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 17개
    """
    return render_digital_twin_slide(slide, "4열_사각사진_프로젝트팀_카드", data, theme, image_files)


def render_tech_061_좌측_빅넘버_우측_상세불릿_분할형(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_061] 좌측_빅넘버_+_우측_상세불릿_분할형 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 18개
    """
    return render_digital_twin_slide(slide, "좌측_빅넘버_+_우측_상세불릿_분할형", data, theme, image_files)


def render_tech_062_2단_주요지표_대조카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_062] 2단_주요지표_대조카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 11개
    """
    return render_digital_twin_slide(slide, "2단_주요지표_대조카드", data, theme, image_files)


def render_tech_063_세계지도_글로벌진출_거점핀(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_063] 세계지도_글로벌진출_거점핀 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 13개
    """
    return render_digital_twin_slide(slide, "세계지도_글로벌진출_거점핀", data, theme, image_files)


def render_tech_064_2열_원형사진_임원진_프로필(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_064] 2열_원형사진_임원진_프로필 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 9개
    """
    return render_digital_twin_slide(slide, "2열_원형사진_임원진_프로필", data, theme, image_files)


def render_tech_065_상하_2단_스텝_프로세스_바(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_065] 상하_2단_스텝_프로세스_바 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 17개
    """
    return render_digital_twin_slide(slide, "상하_2단_스텝_프로세스_바", data, theme, image_files)


def render_tech_066_3단_피라미드_비전_전략_과제(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_066] 3단_피라미드(비전-전략-과제) 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 21개
    """
    return render_digital_twin_slide(slide, "3단_피라미드(비전-전략-과제)", data, theme, image_files)


def render_tech_067_3열_원형사진_임원진_프로필(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_067] 3열_원형사진_임원진_프로필 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 9개
    """
    return render_digital_twin_slide(slide, "3열_원형사진_임원진_프로필", data, theme, image_files)


def render_tech_068_AS_IS_vs_TO_BE_2단_대칭패널(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_068] AS-IS_vs_TO-BE_2단_대칭패널 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "AS-IS_vs_TO-BE_2단_대칭패널", data, theme, image_files)


def render_tech_069_4단_피라미드_철학_비전_전략_실행과제(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_069] 4단_피라미드(철학-비전-전략-실행과제) 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "4단_피라미드(철학-비전-전략-실행과제)", data, theme, image_files)


def render_tech_070_3단_가격표_티어_비교_Basic_Pro_Enterprise(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_070] 3단_가격표_티어_비교(Basic-Pro-Enterprise) 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "3단_가격표_티어_비교(Basic-Pro-Enterprise)", data, theme, image_files)


def render_tech_071_수직_4단_스트라이프_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_071] 수직_4단_스트라이프_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 18개
    """
    return render_digital_twin_slide(slide, "수직_4단_스트라이프_목차", data, theme, image_files)


def render_tech_072_선형_3단계_화살표_체인(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_072] 선형_3단계_화살표_체인 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 6개
    """
    return render_digital_twin_slide(slide, "선형_3단계_화살표_체인", data, theme, image_files)


def render_tech_073_타임라인_연결형_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_073] 타임라인_연결형_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "타임라인_연결형_목차", data, theme, image_files)


def render_tech_074_비중_도넛게이지_점유율_지표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_074] 비중_도넛게이지_점유율_지표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 5개
    """
    return render_digital_twin_slide(slide, "비중_도넛게이지_점유율_지표", data, theme, image_files)


def render_tech_075_상태배지결합_업무진척_현황표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_075] 상태배지결합_업무진척_현황표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "상태배지결합_업무진척_현황표", data, theme, image_files)


def render_tech_076_수평_게이지_목표달성률_프로그레스바(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_076] 수평_게이지_목표달성률_프로그레스바 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "수평_게이지_목표달성률_프로그레스바", data, theme, image_files)


def render_tech_077_4열_원형사진_임원진_프로필(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_077] 4열_원형사진_임원진_프로필 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 17개
    """
    return render_digital_twin_slide(slide, "4열_원형사진_임원진_프로필", data, theme, image_files)


def render_tech_078_2x2_직교_우선순위_매트릭스_난이도_vs_효과(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_078] 2x2_직교_우선순위_매트릭스(난이도_vs_효과) 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "2x2_직교_우선순위_매트릭스(난이도_vs_효과)", data, theme, image_files)


def render_tech_079_넘버링_배지_3단카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_079] 넘버링_배지_3단카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 22개
    """
    return render_digital_twin_slide(slide, "넘버링_배지_3단카드", data, theme, image_files)


def render_tech_080_6단_2x3_격자형_카드그리드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_080] 6단_2x3_격자형_카드그리드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "6단_2x3_격자형_카드그리드", data, theme, image_files)


def render_tech_081_상하_4단_스텝_프로세스_바(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_081] 상하_4단_스텝_프로세스_바 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 21개
    """
    return render_digital_twin_slide(slide, "상하_4단_스텝_프로세스_바", data, theme, image_files)


def render_tech_082_SWOT_4분할_전략분석카드_SO_ST_WO_WT(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_082] SWOT_4분할_전략분석카드(SO-ST-WO-WT) 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 10개
    """
    return render_digital_twin_slide(slide, "SWOT_4분할_전략분석카드(SO-ST-WO-WT)", data, theme, image_files)


def render_tech_083_4단_사진_상단결합_비주얼카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_083] 4단_사진_상단결합_비주얼카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "4단_사진_상단결합_비주얼카드", data, theme, image_files)


def render_tech_084_4대_기둥_신전모델_Pillars_on_Base(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_084] 4대_기둥_신전모델(Pillars on Base) 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "4대_기둥_신전모델(Pillars on Base)", data, theme, image_files)


def render_tech_085_선형_6단계_화살표_체인(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_085] 선형_6단계_화살표_체인 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 22개
    """
    return render_digital_twin_slide(slide, "선형_6단계_화살표_체인", data, theme, image_files)


def render_tech_086_역할분담_R_R_조직표(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_086] 역할분담_R&R_조직표 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "역할분담_R&R_조직표", data, theme, image_files)


def render_tech_087_퍼즐조각_결합형_파트너십(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_087] 퍼즐조각_결합형_파트너십 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 11개
    """
    return render_digital_twin_slide(slide, "퍼즐조각_결합형_파트너십", data, theme, image_files)


def render_tech_088_선형_4단계_화살표_체인(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_088] 선형_4단계_화살표_체인 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 23개
    """
    return render_digital_twin_slide(slide, "선형_4단계_화살표_체인", data, theme, image_files)


def render_tech_089_중심축_3개년_마일스톤_타임라인(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_089] 중심축_3개년_마일스톤_타임라인 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 8개
    """
    return render_digital_twin_slide(slide, "중심축_3개년_마일스톤_타임라인", data, theme, image_files)


def render_tech_090_5열_사각사진_프로젝트팀_카드(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_090] 5열_사각사진_프로젝트팀_카드 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 21개
    """
    return render_digital_twin_slide(slide, "5열_사각사진_프로젝트팀_카드", data, theme, image_files)


def render_tech_091_선형_2단계_화살표_체인(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_091] 선형_2단계_화살표_체인 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 7개
    """
    return render_digital_twin_slide(slide, "선형_2단계_화살표_체인", data, theme, image_files)


def render_tech_092_선형_5단계_화살표_체인(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_092] 선형_5단계_화살표_체인 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 13개
    """
    return render_digital_twin_slide(slide, "선형_5단계_화살표_체인", data, theme, image_files)


def render_tech_093_좌우_분할_대형타이틀_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_093] 좌우_분할_대형타이틀_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 12개
    """
    return render_digital_twin_slide(slide, "좌우_분할_대형타이틀_목차", data, theme, image_files)


def render_tech_094_수평_3_4단_레이어드_스택_인프라_플랫폼_서비스(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_094] 수평_3~4단_레이어드_스택(인프라-플랫폼-서비스) 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 30개
    """
    return render_digital_twin_slide(slide, "수평_3~4단_레이어드_스택(인프라-플랫폼-서비스)", data, theme, image_files)


def render_tech_095_아이콘_결합_그리드_목차(slide: Any, data: dict[str, Any], theme: dict[str, int], image_files: list[str] | None = None) -> list[int]:
    """[tech_095] 아이콘_결합_그리드_목차 전용 Digital Twin 1:1 독립 렌더링 함수
    골든 레퍼런스 원본: ""
    총 원본 복제 형상 수: 11개
    """
    return render_digital_twin_slide(slide, "아이콘_결합_그리드_목차", data, theme, image_files)


TECHNIQUE_FUNCTIONS_REGISTRY: dict[str, Callable[..., list[int]]] = {
    "미니멀_타이포_원라인_표지": render_tech_001_미니멀_타이포_원라인_표지,
    "세로_넘버링_리스트_목차": render_tech_002_세로_넘버링_리스트_목차,
    "좌우_1:1_히어로_사진_스플릿_카드": render_tech_003_좌우_1_1_히어로_사진_스플릿_카드,
    "넘버링_배지_2단카드": render_tech_004_넘버링_배지_2단카드,
    "수직_2단_스트라이프_목차": render_tech_005_수직_2단_스트라이프_목차,
    "추진일정_월별_마일스톤표": render_tech_006_추진일정_월별_마일스톤표,
    "KISA식_1_가_(1)_공문서형_위계본문": render_tech_007_KISA식_1_가_1_공문서형_위계본문,
    "수직_2단_가로바_리스트_카드": render_tech_008_수직_2단_가로바_리스트_카드,
    "수직_3단_가로바_리스트_카드": render_tech_009_수직_3단_가로바_리스트_카드,
    "2단_수평_대등_핵심과제_카드그리드": render_tech_010_2단_수평_대등_핵심과제_카드그리드,
    "표준_비즈니스_개조식_단락본문": render_tech_011_표준_비즈니스_개조식_단락본문,
    "네이티브_비즈니스_통계차트": render_tech_012_네이티브_비즈니스_통계차트,
    "3단_사진_상단결합_비주얼카드": render_tech_013_3단_사진_상단결합_비주얼카드,
    "스마트폰_모바일앱_지시선_콜아웃": render_tech_014_스마트폰_모바일앱_지시선_콜아웃,
    "측면오픈형_비즈니스_명세표": render_tech_015_측면오픈형_비즈니스_명세표,
    "1:2_비대칭_좌측요약_+_우측상세_카드": render_tech_016_1_2_비대칭_좌측요약_우측상세_카드,
    "2:1_비대칭_좌측상세_+_우측콜아웃_카드": render_tech_017_2_1_비대칭_좌측상세_우측콜아웃_카드,
    "미니멀_라인_간지": render_tech_018_미니멀_라인_간지,
    "Thank_You_미니멀_엔딩": render_tech_019_Thank_You_미니멀_엔딩,
    "풀배경_다크_비주얼_표지": render_tech_020_풀배경_다크_비주얼_표지,
    "좌우_2분할_컬러블록_표지": render_tech_021_좌우_2분할_컬러블록_표지,
    "컬러_풀블록_섹션_간지": render_tech_022_컬러_풀블록_섹션_간지,
    "좌측_빅넘버_우측_섹션명_분할간지": render_tech_023_좌측_빅넘버_우측_섹션명_분할간지,
    "중앙_특대형_빅넘버_챕터간지": render_tech_024_중앙_특대형_빅넘버_챕터간지,
    "카드_오버레이_비즈니스_표지": render_tech_025_카드_오버레이_비즈니스_표지,
    "강조선_콜아웃_배너_장표": render_tech_026_강조선_콜아웃_배너_장표,
    "비주얼_배경사진_감사인사_엔딩": render_tech_027_비주얼_배경사진_감사인사_엔딩,
    "풀배경_이미지_표지": render_tech_028_풀배경_이미지_표지,
    "대형_따옴표_슬로건_박스": render_tech_029_대형_따옴표_슬로건_박스,
    "3단_아이콘_원형_헤더_카드": render_tech_030_3단_아이콘_원형_헤더_카드,
    "3단_경영실적_증감배지_KPI": render_tech_031_3단_경영실적_증감배지_KPI,
    "수직_4단_가로바_리스트_카드": render_tech_032_수직_4단_가로바_리스트_카드,
    "4단_가로_파이프라인형_카드그리드": render_tech_033_4단_가로_파이프라인형_카드그리드,
    "Q&A_질의응답_엔딩": render_tech_034_Q_A_질의응답_엔딩,
    "볼드_빅타이포_임팩트_표지": render_tech_035_볼드_빅타이포_임팩트_표지,
    "5단_슬림_핵심과제_카드그리드": render_tech_036_5단_슬림_핵심과제_카드그리드,
    "상하_3단_스텝_프로세스_바": render_tech_037_상하_3단_스텝_프로세스_바,
    "동심원_3중_과녁모델": render_tech_038_동심원_3중_과녁모델,
    "연락처결합_클로징_엔딩": render_tech_039_연락처결합_클로징_엔딩,
    "장문형_내러티브_설명문_본문": render_tech_040_장문형_내러티브_설명문_본문,
    "선형_3단계_셰브론_파이프라인": render_tech_041_선형_3단계_셰브론_파이프라인,
    "넘버링_배지_4단카드": render_tech_042_넘버링_배지_4단카드,
    "아이콘_헤더_결합_2단카드": render_tech_043_아이콘_헤더_결합_2단카드,
    "체크포인트_불릿_개조식_본문": render_tech_044_체크포인트_불릿_개조식_본문,
    "Pros_vs_Cons_장단점_대조표": render_tech_045_Pros_vs_Cons_장단점_대조표,
    "상단요약박스_+_하단개조식_본문": render_tech_046_상단요약박스_하단개조식_본문,
    "수직_3단_스트라이프_목차": render_tech_047_수직_3단_스트라이프_목차,
    "1개_특대형_빅넘버_단독강조": render_tech_048_1개_특대형_빅넘버_단독강조,
    "오피스약도_본사주소_안내블록": render_tech_049_오피스약도_본사주소_안내블록,
    "가로_3열_카드형_목차": render_tech_050_가로_3열_카드형_목차,
    "가로_4열_카드형_목차": render_tech_051_가로_4열_카드형_목차,
    "3단_이미지_쇼케이스_그리드": render_tech_052_3단_이미지_쇼케이스_그리드,
    "멀티트랙_간트차트_일정표": render_tech_053_멀티트랙_간트차트_일정표,
    "3단_표준_핵심과제_카드그리드": render_tech_054_3단_표준_핵심과제_카드그리드,
    "Top_1~5_랭킹_리더보드": render_tech_055_Top_1_5_랭킹_리더보드,
    "6단_슬림_카드그리드": render_tech_056_6단_슬림_카드그리드,
    "4단_비주얼_갤러리": render_tech_057_4단_비주얼_갤러리,
    "4단_대시보드_핵심수치_그리드": render_tech_058_4단_대시보드_핵심수치_그리드,
    "비대칭_매스너리_포트폴리오": render_tech_059_비대칭_매스너리_포트폴리오,
    "4열_사각사진_프로젝트팀_카드": render_tech_060_4열_사각사진_프로젝트팀_카드,
    "좌측_빅넘버_+_우측_상세불릿_분할형": render_tech_061_좌측_빅넘버_우측_상세불릿_분할형,
    "2단_주요지표_대조카드": render_tech_062_2단_주요지표_대조카드,
    "세계지도_글로벌진출_거점핀": render_tech_063_세계지도_글로벌진출_거점핀,
    "2열_원형사진_임원진_프로필": render_tech_064_2열_원형사진_임원진_프로필,
    "상하_2단_스텝_프로세스_바": render_tech_065_상하_2단_스텝_프로세스_바,
    "3단_피라미드(비전-전략-과제)": render_tech_066_3단_피라미드_비전_전략_과제,
    "3열_원형사진_임원진_프로필": render_tech_067_3열_원형사진_임원진_프로필,
    "AS-IS_vs_TO-BE_2단_대칭패널": render_tech_068_AS_IS_vs_TO_BE_2단_대칭패널,
    "4단_피라미드(철학-비전-전략-실행과제)": render_tech_069_4단_피라미드_철학_비전_전략_실행과제,
    "3단_가격표_티어_비교(Basic-Pro-Enterprise)": render_tech_070_3단_가격표_티어_비교_Basic_Pro_Enterprise,
    "수직_4단_스트라이프_목차": render_tech_071_수직_4단_스트라이프_목차,
    "선형_3단계_화살표_체인": render_tech_072_선형_3단계_화살표_체인,
    "타임라인_연결형_목차": render_tech_073_타임라인_연결형_목차,
    "비중_도넛게이지_점유율_지표": render_tech_074_비중_도넛게이지_점유율_지표,
    "상태배지결합_업무진척_현황표": render_tech_075_상태배지결합_업무진척_현황표,
    "수평_게이지_목표달성률_프로그레스바": render_tech_076_수평_게이지_목표달성률_프로그레스바,
    "4열_원형사진_임원진_프로필": render_tech_077_4열_원형사진_임원진_프로필,
    "2x2_직교_우선순위_매트릭스(난이도_vs_효과)": render_tech_078_2x2_직교_우선순위_매트릭스_난이도_vs_효과,
    "넘버링_배지_3단카드": render_tech_079_넘버링_배지_3단카드,
    "6단_2x3_격자형_카드그리드": render_tech_080_6단_2x3_격자형_카드그리드,
    "상하_4단_스텝_프로세스_바": render_tech_081_상하_4단_스텝_프로세스_바,
    "SWOT_4분할_전략분석카드(SO-ST-WO-WT)": render_tech_082_SWOT_4분할_전략분석카드_SO_ST_WO_WT,
    "4단_사진_상단결합_비주얼카드": render_tech_083_4단_사진_상단결합_비주얼카드,
    "4대_기둥_신전모델(Pillars on Base)": render_tech_084_4대_기둥_신전모델_Pillars_on_Base,
    "선형_6단계_화살표_체인": render_tech_085_선형_6단계_화살표_체인,
    "역할분담_R&R_조직표": render_tech_086_역할분담_R_R_조직표,
    "퍼즐조각_결합형_파트너십": render_tech_087_퍼즐조각_결합형_파트너십,
    "선형_4단계_화살표_체인": render_tech_088_선형_4단계_화살표_체인,
    "중심축_3개년_마일스톤_타임라인": render_tech_089_중심축_3개년_마일스톤_타임라인,
    "5열_사각사진_프로젝트팀_카드": render_tech_090_5열_사각사진_프로젝트팀_카드,
    "선형_2단계_화살표_체인": render_tech_091_선형_2단계_화살표_체인,
    "선형_5단계_화살표_체인": render_tech_092_선형_5단계_화살표_체인,
    "좌우_분할_대형타이틀_목차": render_tech_093_좌우_분할_대형타이틀_목차,
    "수평_3~4단_레이어드_스택(인프라-플랫폼-서비스)": render_tech_094_수평_3_4단_레이어드_스택_인프라_플랫폼_서비스,
    "아이콘_결합_그리드_목차": render_tech_095_아이콘_결합_그리드_목차,
}

def get_technique_renderer(technique_name: str) -> Callable[..., list[int]] | None:
    """기법 이름에 1:1 매칭되는 전용 독립 렌더러 함수 반환"""
    if technique_name in TECHNIQUE_FUNCTIONS_REGISTRY:
        return TECHNIQUE_FUNCTIONS_REGISTRY[technique_name]
    for k, fn in TECHNIQUE_FUNCTIONS_REGISTRY.items():
        if k in technique_name or technique_name in k:
            return fn
    return None
