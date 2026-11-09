# -*- coding: utf-8 -*-
"""
mcp_server/tools/folder_inspector.py
사용자 실무 문서 폴더 벌크 학습 및
1) 기존 70개 도구 미세조정(테마/폰트/여백 주입)
2) 사내 고유 서식 발견 시 신규 도구/레시피 자동 합성 엔진.
"""

from __future__ import annotations

import json
import logging
from collections import Counter
from pathlib import Path
from typing import Any

from .blueprint_analyzers import (
    detect_doc_type,
    get_analyzer_for_type,
)
from .ppt_slice_engine import DEFAULT_THEMES, CATALOG_DATA

logger = logging.getLogger("mcp_server.folder_inspector")

PROFILES_DIR = Path(__file__).resolve().parents[1] / "user_profiles"
CUSTOM_RECIPES_DIR = Path(__file__).resolve().parents[1] / "blueprints" / "custom_recipes"

PROFILES_DIR.mkdir(parents=True, exist_ok=True)
CUSTOM_RECIPES_DIR.mkdir(parents=True, exist_ok=True)

ACTIVE_PROFILE_FILE = PROFILES_DIR / "_active_profile.json"


def bgr_to_hex(bgr: int) -> str:
    """Windows BGR 정수를 #RRGGBB HEX 문자열로 변환"""
    r = bgr & 0xFF
    g = (bgr >> 8) & 0xFF
    b = (bgr >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


class FolderInspector:
    """사용자 실무 폴더 벌크 학습 및 서식 시스템 합성 엔진"""

    @classmethod
    def inspect_and_learn_folder(
        cls,
        folder_path: Path,
        profile_name: str = "my_office_profile",
        synthesize_new_tools: bool = True,
        recursive: bool = False,
    ) -> dict[str, Any]:
        """
        폴더 내 모든 오피스 문서(PPT, Excel, Word, HWP)를 순회하며
        자주 쓰이는 브랜드 컬러, 폰트, 여백, 모서리 라운딩 규격을 전수 분석하여
        1) 기존 도구 맞춤 프로필로 적재하고
        2) 고유 서식 발견 시 신규 슬라이스 레시피를 자동 합성합니다.
        """
        if not folder_path.exists() or not folder_path.is_dir():
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {folder_path}")

        pattern = "**/*" if recursive else "*"
        supported_exts = {".pptx", ".ppt", ".xlsx", ".xls", ".xlsm", ".docx", ".doc", ".hwp", ".hwpx"}
        target_files = [f for f in folder_path.glob(pattern) if f.is_file() and f.suffix.lower() in supported_exts]

        if not target_files:
            raise ValueError(f"폴더 내에 지원하는 오피스 문서가 없습니다: {folder_path}")

        logger.info(f"폴더 벌크 학습 시작: {folder_path} (총 {len(target_files)}개 문서)")

        files_by_type: dict[str, list[Path]] = {}
        for f in target_files:
            try:
                dt = detect_doc_type(f)
                files_by_type.setdefault(dt, []).append(f)
            except Exception:
                continue

        color_counter: Counter[int] = Counter()
        font_counter: Counter[str] = Counter()
        bg_colors: list[int] = []
        corner_adjustments: list[float] = []
        layout_candidates: list[dict[str, Any]] = []

        total_inspected = 0

        # 포맷별 분석기 가동 (DRM 친화적 배치 개방-확인-종료 루프)
        for dt, file_list in files_by_type.items():
            analyzer = get_analyzer_for_type(dt)
            try:
                docs = analyzer.analyze_files_batch(file_list, pause_sec=0.2)
                for doc in docs:
                    total_inspected += 1
                    # 폰트 집계
                    for fn in doc.get("design_tokens", {}).get("fonts", []):
                        if fn and fn != "None":
                            font_counter[fn] += 1

                    # 색상 집계
                    sample_colors = (
                        doc.get("design_tokens", {}).get("sample_colors_rgb", [])
                        or doc.get("design_tokens", {}).get("sample_colors_bgr", [])
                    )
                    for c in sample_colors:
                        if c not in (0, 16777215):  # 흑/백 제외
                            color_counter[c] += 1

                    # 슬라이드별 라운딩 곡률 및 고유 레이아웃 탐색 (PPT)
                    if dt == "ppt":
                        for page in doc.get("pages", []):
                            bg_rgb = page.get("background", {}).get("rgb")
                            if bg_rgb:
                                bg_colors.append(bg_rgb)

                            shapes = [el.get("shape_spec", {}) for el in page.get("elements", [])]
                            for sh in shapes:
                                for adj in sh.get("adjustments", []):
                                    if 0.01 <= adj <= 0.3:
                                        corner_adjustments.append(adj)

                            # 도형 4개 이상이고 텍스트가 풍부한 장표는 신규 레시피 후보로 수집
                            if len(shapes) >= 4:
                                layout_candidates.append({
                                    "source_file": doc.get("source_name"),
                                    "page_index": page.get("page_index"),
                                    "shapes_count": len(shapes),
                                    "shapes": shapes,
                                    "background": page.get("background", {}),
                                })

            except Exception as e:
                logger.warning(f"{dt.upper()} 문서군 분석 중 일부 오류 발생: {e}")

        # 디자인 시스템 토큰 도출
        most_common_fonts = [fn for fn, _ in font_counter.most_common(3)]
        title_font = most_common_fonts[0] if most_common_fonts else "맑은 고딕"
        body_font = most_common_fonts[1] if len(most_common_fonts) > 1 else title_font

        common_colors = [c for c, _ in color_counter.most_common(5)]
        primary_color = bgr_to_hex(common_colors[0]) if common_colors else "#1F3864"
        accent_color = bgr_to_hex(common_colors[1]) if len(common_colors) > 1 else "#2E608A"
        sub_accent = bgr_to_hex(common_colors[2]) if len(common_colors) > 2 else "#0284C7"

        avg_corner = round(sum(corner_adjustments) / len(corner_adjustments), 4) if corner_adjustments else 0.0622

        bg_hex = "#FFFFFF"
        if bg_colors:
            most_common_bg = Counter(bg_colors).most_common(1)[0][0]
            bg_hex = bgr_to_hex(most_common_bg)
        panel_hex = "#F2F4F7" if bg_hex.upper() in ("#FFFFFF", "#FFF") else "#FFFFFF"

        # 신규 도구/레시피 합성 (Synthesize New Tools)
        synthesized_tools: list[dict[str, Any]] = []
        if synthesize_new_tools and layout_candidates:
            synthesized_tools = cls._synthesize_custom_recipes(layout_candidates, profile_name)

        profile = {
            "profile_name": profile_name,
            "folder_path": str(folder_path.resolve()),
            "total_documents_analyzed": total_inspected,
            "fonts": {
                "title_font": title_font,
                "body_font": body_font,
                "detected_fonts": list(font_counter.keys())[:5],
            },
            "palette": {
                "primary": primary_color,
                "accent": accent_color,
                "sub": sub_accent,
                "bg": bg_hex,
                "panel": panel_hex,
                "text": "#1F2328",
            },
            "layout_tokens": {
                "corner_radius": avg_corner,
                "gap": 18.0,
                "margin_horizontal": 40.0,
            },
            "synthesized_tools": synthesized_tools,
        }

        # 프로필 영구 적재
        p_file = PROFILES_DIR / f"{profile_name}.json"
        p_file.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")

        # 활성 프로필로 등록 및 기존 PptSliceEngine 실시간 미세조정 동기화
        cls.set_active_profile(profile_name)

        return profile

    @classmethod
    def _synthesize_custom_recipes(
        cls,
        candidates: list[dict[str, Any]],
        profile_name: str,
        max_recipes: int = 3,
    ) -> list[dict[str, Any]]:
        """고유 서식 후보군으로부터 신규 슬라이스 레시피 합성 및 저장"""
        synthesized = []
        # 도형 수 기준 정렬하여 구조가 탄탄한 상위 장표 선별
        sorted_candidates = sorted(candidates, key=lambda c: c["shapes_count"], reverse=True)

        for idx, cand in enumerate(sorted_candidates[:max_recipes], start=1):
            recipe_id = f"custom_{profile_name}_{idx}"
            label = f"{cand['source_file']} 기반 고유 레이아웃 #{idx}"

            # 가변 슬롯 추출
            slots = []
            shapes = cand.get("shapes", [])
            for s_idx, sh in enumerate(shapes):
                if "text" in sh and sh["text"].get("content", "").strip():
                    slot_key = f"slot_{s_idx + 1}"
                    sh["slot_key"] = slot_key
                    slots.append({
                        "slot_key": slot_key,
                        "sample_text": sh["text"]["content"][:30],
                    })

            recipe_data = {
                "recipe_id": recipe_id,
                "name": label,
                "profile_name": profile_name,
                "source_file": cand["source_file"],
                "shapes_count": len(shapes),
                "slots_count": len(slots),
                "slots": slots,
                "background": cand.get("background", {}),
                "shapes": shapes,
            }

            r_file = CUSTOM_RECIPES_DIR / f"{recipe_id}.json"
            r_file.write_text(json.dumps(recipe_data, ensure_ascii=False, indent=2), encoding="utf-8")

            synthesized.append({
                "recipe_id": recipe_id,
                "tool_role": recipe_id,
                "name": label,
                "slots_count": len(slots),
                "usage_example": f"ppt_insert_slice(role='{recipe_id}', data={{...}})",
            })

            logger.info(f"신규 슬라이스 레시피 합성 완료: {recipe_id} ({len(slots)}개 슬롯)")

        return synthesized

    @classmethod
    def set_active_profile(cls, profile_name: str) -> dict[str, Any]:
        """특정 프로필을 모든 기존 도구의 기본 스타일로 활성화"""
        p_file = PROFILES_DIR / f"{profile_name}.json"
        if not p_file.exists():
            raise FileNotFoundError(f"프로필 '{profile_name}'이 존재하지 않습니다.")

        ACTIVE_PROFILE_FILE.write_text(json.dumps({"active_profile": profile_name}), encoding="utf-8")
        profile_data = json.loads(p_file.read_text(encoding="utf-8"))

        # 기존 도구(PptSliceEngine) 테마 딕셔너리 동적 주입
        pal = profile_data.get("palette", {})
        custom_theme = {
            "name": f"사용자 정의 서식 ({profile_name})",
            "primary": pal.get("primary", "#1F3864"),
            "secondary": pal.get("sub", "#0055AA"),
            "accent": pal.get("accent", "#2E608A"),
            "bg_light": pal.get("bg", "#FFFFFF"),
            "card_border": pal.get("panel", "#E2E8F0"),
            "text_dark": pal.get("text", "#1F2328"),
            "text_muted": "#475569",
            # 레거시 호환용 키
            "bg": pal.get("bg", "#FFFFFF"),
            "panel": pal.get("panel", "#F2F4F7"),
            "text": pal.get("text", "#1F2328"),
            "sub": pal.get("sub", "#0284C7"),
        }
        DEFAULT_THEMES[profile_name] = custom_theme
        CATALOG_DATA.setdefault("themes", {})[profile_name] = custom_theme

        # 합성된 신규 레시피들을 PptSliceEngine / PptRecipeEngine에 동적 등록
        for synth in profile_data.get("synthesized_tools", []):
            r_id = synth["recipe_id"]
            r_path = CUSTOM_RECIPES_DIR / f"{r_id}.json"
            if r_path.exists():
                r_content = json.loads(r_path.read_text(encoding="utf-8"))
                # PptSliceEngine의 CATALOG_DATA에 동적 등록
                CATALOG_DATA.setdefault("slice_roles", {})[r_id] = {
                    "name": synth["name"],
                    "description": f"사용자 폴더에서 자동 합성된 고유 서식 ({r_id})",
                    "slots": r_content.get("slots", []),
                    "recipe_data": r_content,
                }

        logger.info(f"활성 프로필 전환 및 기존 도구 동기화 완료: {profile_name}")
        return profile_data

    @classmethod
    def get_active_profile(cls) -> dict[str, Any] | None:
        """현재 활성화된 프로필 반환"""
        if ACTIVE_PROFILE_FILE.exists():
            try:
                active_name = json.loads(ACTIVE_PROFILE_FILE.read_text(encoding="utf-8")).get("active_profile")
                if active_name:
                    p_file = PROFILES_DIR / f"{active_name}.json"
                    if p_file.exists():
                        return json.loads(p_file.read_text(encoding="utf-8"))
            except Exception:
                pass
        return None

    @classmethod
    def list_all_profiles(cls) -> list[dict[str, Any]]:
        """저장된 전체 서식 프로필 및 합성 도구 목록 반환"""
        profiles = []
        active_info = cls.get_active_profile()
        active_name = active_info.get("profile_name") if active_info else None

        for f in PROFILES_DIR.glob("*.json"):
            if f.name.startswith("_"):
                continue
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
                profiles.append({
                    "profile_name": data.get("profile_name", f.stem),
                    "is_active": (data.get("profile_name") == active_name),
                    "total_documents": data.get("total_documents_analyzed", 0),
                    "palette": data.get("palette", {}),
                    "fonts": data.get("fonts", {}),
                    "synthesized_tools_count": len(data.get("synthesized_tools", [])),
                    "synthesized_tools": data.get("synthesized_tools", []),
                })
            except Exception:
                pass
        return profiles
