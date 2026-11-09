# -*- coding: utf-8 -*-
"""
mcp_server/tools/ppt_morph_engine.py
iOfficeAI(OfficeCLI) morph-ppt & officecli-pitch-deck 규격 기반 PowerPoint 실시간 COM 엔진:
1. 3-Prefix Morph 안무 (`!!scene-*`, `!!actor-*`, `#sN-*` + 캔버스 외곽 고스팅)
2. Morph 무결성 감사기 (`MorphAuditor`: 전환 누락, 유령 도형 잔존, M-2 누적, 중복 좌표, 변위 부족 점검)
3. Pitch Deck 단계별(5종) x 버티컬별(5종) 아키타입 플래너 (`PitchDeckPlanner` + C.1~C.5b 기하 규격)
"""

from __future__ import annotations

import math
import re
from typing import Any

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr
from .parsers import strip_markdown_syntax
from .ppt_tools import ALIGN_MAP, MSO_SHAPE_MAP, _resolve_target_slide, ppt_freeze_display

PP_EFFECT_MORPH = 3954  # PowerPoint ppEffectMorph 상수
CM_TO_PT = 72.0 / 2.54  # 1cm ≈ 28.3465pt
MIN_MORPH_DISPLACEMENT_PT = 5.0 * CM_TO_PT  # ≥ 5cm (≈ 141.73pt)
MIN_MORPH_ROTATION_DEG = 15.0
MIN_MORPH_SCALE_DELTA_RATIO = 0.30

_SLIDE_PREFIX_RE = re.compile(r"^#s(\d+)-")

STAGE_PROFILES: dict[str, dict[str, Any]] = {
    "Seed": {
        "revenue_band": "$0 – $1M ARR (또는 Pre-revenue)",
        "team_size": "2 – 8 FTE",
        "slide_range": (10, 12),
        "narrative_weight": "Problem (30%) + Solution (25%) + Team (15%) + Market (15%) + Traction (15%)",
        "must_have": ["창업자-시장 적합성(Founder-Market Fit)", "초기 파일럿/디자인 파트너 1~2곳", "Top-down TAM"],
        "red_flag": "초기 고객 10곳 수준을 '시장 검증 완료'로 과대 포장",
    },
    "Series A": {
        "revenue_band": "$1M – $5M ARR",
        "team_size": "10 – 25 FTE",
        "slide_range": (12, 16),
        "narrative_weight": "Problem (20%) + Solution (20%) + Market Why-Now (15%) + Product (15%) + Traction (20%) + Team (10%)",
        "must_have": ["PMF 입증(NRR > 110%, 저이탈률)", "Bottom-up TAM/SAM", "파일럿 유료 전환 파이프라인"],
        "red_flag": "인위적 Bottom-up TAM 산정 또는 초기 단계에서 무의미한 CAC 강조",
    },
    "Series B": {
        "revenue_band": "$5M – $30M ARR",
        "team_size": "30 – 100 FTE",
        "slide_range": (18, 22),
        "narrative_weight": "Traction + Unit Economics (30%) + Market + Product + Team + Financials & Ask",
        "must_have": ["0부터 시작하는 ARR 성장 곡선", "NRR/CAC/LTV/Payback(<18개월)", "코호트 리텐션 & 로고월"],
        "red_flag": "단위경제(Unit Econ) 슬라이드 누락, Use-of-Funds 비중(%) 누락",
    },
    "Series C": {
        "revenue_band": "$30M+ ARR",
        "team_size": "100+ FTE",
        "slide_range": (20, 24),
        "narrative_weight": "Financials + Scale + Moat (40%) + Market Expansion + Team Depth",
        "must_have": ["다개년 GAAP 재무추이", "Rule-of-40 / 매출총이익률 추세", "글로벌 확장 및 해자(Moat)"],
        "red_flag": "해자(Moat) 슬라이드 누락, 마진 스토리 없는 매출 성장 나열",
    },
    "Bridge": {
        "revenue_band": "단계 무관 (Bridge / SAFE)",
        "team_size": "현 체제 유지",
        "slide_range": (8, 10),
        "narrative_weight": "Specific Bridge Milestone + Runway Math + Investor Commitments",
        "must_have": ["직전 라운드 이후 핵심 성과", "브릿지 자금으로 달성할 단일 마일스톤", "확정 투자자 약정 규모"],
        "red_flag": "브릿지 라운드에 정규 Series A처럼 20장 이상 장황하게 구성",
    },
}

VERTICAL_PROFILES: dict[str, dict[str, Any]] = {
    "B2B SaaS": {
        "arc": ["Cover", "TL;DR", "Problem", "Solution", "Product Loop", "Market (TAM/SAM/SOM)", "Unit Economics (CAC/LTV/Payback)", "ARR Trajectory", "Retention Cohort", "Logo Wall", "Team", "Financials", "Ask & Use of Funds"],
        "must_have": ["Unit Economics", "Cohort Retention", "Logo Wall"],
        "unit_econ_equivalent": "CAC, LTV, Payback Period (<18개월), Gross Margin",
    },
    "Consumer": {
        "arc": ["Cover", "Hook Vision", "Problem", "Solution Screenshots", "Product Experience Flow", "Why Now Window", "Pre-orders / Early Sales", "Retention & Engagement (DAU/D30)", "Market", "Competitive Edge", "Founder Story & Team", "Financials", "Ask & Use of Funds"],
        "must_have": ["Product Visuals (>=3장)", "Why Now Window", "DAU/D30 Engagement"],
        "unit_econ_equivalent": "공헌이익(Contribution Margin), 재방문율(Repeat Rate), Blended CAC",
    },
    "DeepTech": {
        "arc": ["Cover", "Core Thesis", "Problem (SoTA 한계)", "Technical Solution", "Architecture", "Benchmarks vs SoTA", "Pipeline & TRL Level", "Long-tail Market", "Commercial Pilots / LOIs", "IP & Patents", "Research Team", "Financials", "Ask & Use of Funds"],
        "must_have": ["Benchmarks vs SoTA", "IP / Patents", "Technical Milestones / TRL"],
        "unit_econ_equivalent": "기술 마일스톤, TRL 단계, SoTA 대비 벤치마크 우위",
    },
    "Marketplace": {
        "arc": ["Cover", "Supply-Demand Friction", "Solution", "Two-Sided Product Demo", "Network Effects Diagram", "Liquidity Metrics (GMV / Match Time)", "Cohort Retention", "Category Expansion", "Take-Rate Model", "Team", "Financials", "Ask & Use of Funds"],
        "must_have": ["Liquidity Metrics (GMV/Take Rate)", "Cohort Retention", "Network Effects Diagram"],
        "unit_econ_equivalent": "GMV, Take Rate, 유동성/매칭 속도, 코호트 유지율",
    },
    "Bio": {
        "arc": ["Cover", "Unmet Medical Need", "Mechanism of Action", "Preclinical / Clinical Data (ORR)", "Pipeline Chart", "Standard-of-Care Differentiation", "IP & Exclusivity", "Regulatory Path (IND/FDA)", "Addressable Market", "Clinical Team", "Burn & Runway to Milestone", "Ask & Use of Funds"],
        "must_have": ["Pipeline Chart", "Clinical Data (ORR/Endpoints)", "Burn & Runway"],
        "unit_econ_equivalent": "월 Burn Rate + 차기 임상 마일스톤(IND/Ph1/BLA)까지 Runway",
    },
}

# Pitch Deck C.1 ~ C.5b 레이아웃 기하 규격 (16:9 33.87x19.05cm / 960x540pt 기준)
LAYOUT_CANON: dict[str, dict[str, Any]] = {
    "C.1": {
        "name": "Title / Cover (Dark Gradient)",
        "usage": "표지 슬라이드 (Cover 전용, 44pt+ 타이틀, 18~22pt 태그라인, 하단 브랜드 밴드)",
        "title_min_pt": 36,
        "elements_cm": {"title": (2.0, 5.0, 29.87, 4.0), "tagline": (2.0, 10.0, 29.87, 2.0), "meta": (2.0, 13.0, 29.87, 1.5), "band": (0.0, 18.5, 33.87, 0.55)},
    },
    "C.2": {
        "name": "3-Stat Callout Row",
        "usage": "문제 정의 / Why Now / 핵심 수치 3종 강조 (60~64pt 대형 숫자 + ≥16pt 레이블)",
        "title_min_pt": 36,
        "elements_cm": {"title": (1.5, 1.0, 30.87, 3.0), "stat_width": 9.0, "stat_x": [2.0, 12.5, 23.0], "stat_y": 5.0},
    },
    "C.3": {
        "name": "4-Stat Callout Row",
        "usage": "병렬 지표 4종 요약 (60pt 숫자 x 7cm 폭, ≥16pt 레이블; 5글자 이하 수치 권장)",
        "title_min_pt": 36,
        "elements_cm": {"title": (1.5, 1.0, 30.87, 3.0), "stat_width": 7.0, "stat_x": [1.5, 9.5, 17.5, 25.5], "stat_y": 5.0},
    },
    "C.4": {
        "name": "Chart + Context Split",
        "usage": "좌측 55% 차트 + 우측 2~3개 수치 콜아웃 (Traction, Financials, Market Sizing)",
        "title_min_pt": 36,
        "elements_cm": {"title": (2.0, 1.0, 29.87, 3.0), "chart": (2.0, 4.0, 17.0, 13.0), "context_x": 21.0, "context_width": 11.0},
    },
    "C.5": {
        "name": "Icon-in-Circle 3-Row Vertical",
        "usage": "솔루션 메커니즘 / 3대 가치 제안 (좌측 원형 아이콘 + 우측 제목/설명)",
        "title_min_pt": 36,
        "elements_cm": {"title": (2.0, 1.0, 29.87, 2.5), "icon_x": 2.0, "icon_y": [4.5, 8.5, 12.5], "text_x": 5.5, "text_width": 25.0},
    },
    "C.5b": {
        "name": "2x2 Feature Grid",
        "usage": "동등한 4대 제품/기능 축 (Strict Z-Order: bg -> ellipse -> title -> body)",
        "title_min_pt": 32,
        "elements_cm": {"title": (2.0, 1.0, 29.87, 2.5), "cards": [(1.5, 4.0, 14.5, 7.0), (17.5, 4.0, 14.5, 7.0), (1.5, 12.0, 14.5, 7.0), (17.5, 12.0, 14.5, 7.0)]},
    },
}


def morph_ghost_left_pt(slide_width_pt: float) -> float:
    """캔버스 크기(16:9 960pt 또는 A4 780pt)에 맞춰 안전한 우측 바깥 고스트 X 좌표(pt) 계산."""
    if float(slide_width_pt) >= 900.0:
        return round(36.0 * CM_TO_PT, 2)  # 1020.47pt (x=36cm)
    return round(float(slide_width_pt) + 60.0, 2)


def classify_morph_shape_prefix(name: str) -> str:
    """3-Prefix Morph 네이밍 규칙 분류 ('scene', 'actor', 'slide_content', 'other')."""
    if not name:
        return "other"
    if name.startswith("!!scene-"):
        return "scene"
    if name.startswith("!!actor-"):
        return "actor"
    if _SLIDE_PREFIX_RE.match(name):
        return "slide_content"
    return "other"


def extract_slide_prefix_num(name: str) -> int | None:
    """`#sN-...` 도형 이름에서 슬라이드 번호 N 추출."""
    m = _SLIDE_PREFIX_RE.match(name or "")
    return int(m.group(1)) if m else None


class PitchDeckPlanner:
    """iOfficeAI officecli-pitch-deck 규격 기반 투자유치(Fundraising) 덱 플래너."""

    @staticmethod
    def _normalize_stage(stage: str) -> str:
        raw = stage.strip().lower().replace("-", " ").replace("_", " ")
        mapping = {
            "seed": "Seed",
            "pre seed": "Seed",
            "series a": "Series A",
            "a": "Series A",
            "series b": "Series B",
            "b": "Series B",
            "series c": "Series C",
            "c": "Series C",
            "bridge": "Bridge",
            "safe": "Bridge",
            "convertible": "Bridge",
        }
        return mapping.get(raw, "Series A")

    @staticmethod
    def _normalize_vertical(vertical: str) -> str:
        raw = vertical.strip().lower().replace("-", " ").replace("_", " ")
        if "saas" in raw or "b2b" in raw or "enterprise" in raw:
            return "B2B SaaS"
        if "consumer" in raw or "b2c" in raw or "hardware" in raw or "d2c" in raw:
            return "Consumer"
        if "deep" in raw or "ai" in raw or "frontier" in raw or "robot" in raw or "quantum" in raw:
            return "DeepTech"
        if "market" in raw or "network" in raw or "platform" in raw:
            return "Marketplace"
        if "bio" in raw or "health" in raw or "clinical" in raw or "med" in raw:
            return "Bio"
        return "B2B SaaS"

    @classmethod
    def plan_deck(
        cls,
        stage: str,
        vertical: str,
        company: str = "Company",
        raise_amount: str = "",
        slides_override: int | None = None,
    ) -> dict[str, Any]:
        norm_stage = cls._normalize_stage(stage)
        norm_vertical = cls._normalize_vertical(vertical)
        stage_info = STAGE_PROFILES[norm_stage]
        vert_info = VERTICAL_PROFILES[norm_vertical]

        base_arc = list(vert_info["arc"])
        if slides_override and slides_override > 0:
            if slides_override < len(base_arc):
                # Cover와 Ask는 항상 보존
                arc = [base_arc[0]] + base_arc[1 : max(1, slides_override - 1)] + [base_arc[-1]]
                arc = arc[:slides_override]
            else:
                arc = base_arc
        else:
            arc = base_arc

        # 연속 동일 패턴 배제 원칙: C.1(Cover) 이후 C.2/C.4/C.5/C.5b/C.3 순환 배정
        rotation_patterns = ["C.2", "C.5", "C.4", "C.5b", "C.3"]
        slide_blueprints: list[dict[str, Any]] = []
        prev_pattern = None

        for idx, role in enumerate(arc, start=1):
            role_lower = role.lower()
            if idx == 1:
                pattern = "C.1"
            elif any(k in role_lower for k in ("arr", "chart", "financial", "trajectory", "cohort", "market", "ask", "burn", "liquidity", "benchmark")):
                pattern = "C.4" if prev_pattern != "C.4" else "C.2"
            elif any(k in role_lower for k in ("loop", "flow", "architecture", "demo", "grid")):
                pattern = "C.5b" if prev_pattern != "C.5b" else "C.5"
            elif any(k in role_lower for k in ("solution", "team", "mechanism", "ip", "strategy")):
                pattern = "C.5" if prev_pattern != "C.5" else "C.5b"
            else:
                candidate = rotation_patterns[idx % len(rotation_patterns)]
                pattern = candidate if candidate != prev_pattern else rotation_patterns[(idx + 1) % len(rotation_patterns)]

            prev_pattern = pattern
            is_cover_or_close = idx == 1
            slide_blueprints.append(
                {
                    "slide_index": idx,
                    "role": role,
                    "layout_pattern": pattern,
                    "layout_spec": LAYOUT_CANON[pattern],
                    "requires_speaker_notes": not is_cover_or_close,
                }
            )

        return {
            "success": True,
            "company": company,
            "raise_amount": raise_amount or "TBD",
            "stage": norm_stage,
            "vertical": norm_vertical,
            "stage_profile": stage_info,
            "vertical_requirements": {
                "must_have_slides": vert_info["must_have"],
                "unit_econ_equivalent": vert_info["unit_econ_equivalent"],
            },
            "hard_rules": [
                "Title >= 36pt / Body >= 18pt / Sub-label >= 16pt (H4 Floor)",
                "연속 2개 슬라이드에 동일 레이아웃 패턴(C.1~C.5b) 재사용 금지",
                "트랙션 차트 Y축은 반드시 0부터 시작 (y_min = 0)",
                "모든 콘텐츠 슬라이드(Cover 제외)에 발표자 노트(Speaker Notes) 필수 포함",
                "Use-of-Funds는 4대 영역 비율(%)이 합계 100%가 되도록 명시",
            ],
            "planned_slides_count": len(slide_blueprints),
            "slides": slide_blueprints,
        }


class MorphAuditor:
    """morph-helpers.py의 Morph 검증 알고리즘을 Python/COM 스냅샷용으로 정식 이식한 감사기."""

    _SCENE_KEYWORDS = (
        "ring", "dot", "line", "circle", "rect", "slash",
        "accent", "actor", "star", "triangle", "diamond",
    )

    @classmethod
    def audit_deck(cls, slides_data: list[dict[str, Any]]) -> dict[str, Any]:
        total_slides = len(slides_data)
        if total_slides == 0:
            return {
                "passed": False,
                "slides_checked": 0,
                "issues": [{"code": "empty_deck", "message": "슬라이드가 1장도 없습니다."}],
                "warnings": [],
                "ghost_count": 0,
            }

        issues: list[dict[str, Any]] = []
        warnings: list[dict[str, Any]] = []
        ghost_count = 0
        in_morph_sequence = False

        for slide in slides_data:
            s_idx = int(slide.get("slide_index", 0))
            slide_w = float(slide.get("slide_width_pt", 960.0))
            slide_h = float(slide.get("slide_height_pt", 540.0))
            ghost_threshold_x = slide_w - 5.0
            shapes = slide.get("shapes", [])

            # 1) Ghost 누적 카운트 및 가시 도형 경계 클리핑 점검
            for sh in shapes:
                sh_left = float(sh.get("left", 0.0))
                sh_top = float(sh.get("top", 0.0))
                sh_h = float(sh.get("height", 0.0))
                sh_name = str(sh.get("name", ""))
                if sh_left >= ghost_threshold_x:
                    ghost_count += 1
                elif classify_morph_shape_prefix(sh_name) != "scene":
                    if sh_left < -2.0 or sh_top < -2.0 or (sh_top + sh_h) > slide_h + 2.0:
                        warnings.append(
                            {
                                "slide": s_idx,
                                "code": "visible_shape_clipped",
                                "shape": sh_name,
                                "message": f"슬라이드 {s_idx}의 가시 도형 '{sh_name}'이 캔버스 상하/좌측 경계를 벗어났습니다.",
                            }
                        )

            eff = slide.get("transition_effect")
            has_morph_eff = eff in (PP_EFFECT_MORPH, "morph")
            has_morph_pref = any(classify_morph_shape_prefix(str(s.get("name", ""))) != "other" for s in shapes)

            if s_idx >= 2:
                # 2) Morph 시퀀스 개시 이후의 Morph 전환 효과 누락 탐지 (혼합 덱의 선행 일반 슬라이드 허용)
                if (in_morph_sequence or has_morph_pref) and not has_morph_eff:
                    issues.append(
                        {
                            "slide": s_idx,
                            "code": "missing_morph_transition",
                            "message": f"슬라이드 {s_idx}에 Morph 전환(ppEffectMorph={PP_EFFECT_MORPH})이 설정되지 않았습니다.",
                        }
                    )

            if has_morph_eff or has_morph_pref:
                in_morph_sequence = True

            if s_idx >= 2:

                # 3) 이전 슬라이드 전용 도형(#s{prev}-*) 잔존 (Unghosted Content) 탐지
                prev_prefix = f"#s{s_idx - 1}-"
                for sh in shapes:
                    sh_name = str(sh.get("name", ""))
                    sh_left = float(sh.get("left", 0.0))
                    if sh_name.startswith(prev_prefix) and sh_left < slide_w:
                        issues.append(
                            {
                                "slide": s_idx,
                                "code": "unghosted_prev_content",
                                "shape": sh_name,
                                "message": (
                                    f"슬라이드 {s_idx}에 이전 슬라이드 도형 '{sh_name}'이 가시 영역"
                                    f"(left={sh_left:.1f}pt < {slide_w:.1f}pt)에 남아 있습니다. 우측 바깥으로 고스팅하세요."
                                ),
                            }
                        )

        # 4) M-2 Ghost Accumulation 상한 초과 검사
        expected_max_ghosts = max(50, total_slides * 4)
        if ghost_count > expected_max_ghosts:
            issues.append(
                {
                    "code": "ghost_accumulation_overflow",
                    "ghost_count": ghost_count,
                    "limit": expected_max_ghosts,
                    "message": f"M-2 Ghost Accumulation 초과: 오프캔버스 고스트 도형 {ghost_count}개 (상한 {expected_max_ghosts}개).",
                }
            )

        # 5) 인접 슬라이드 간 동일 텍스트+좌표 중복 및 Actor 최소 변위 검사
        for i in range(1, total_slides):
            prev_s = slides_data[i - 1]
            curr_s = slides_data[i]
            curr_idx = int(curr_s.get("slide_index", i + 1))
            curr_w = float(curr_s.get("slide_width_pt", 960.0))

            prev_shapes = prev_s.get("shapes", [])
            curr_shapes = curr_s.get("shapes", [])

            prev_actors: dict[str, dict[str, Any]] = {}
            prev_boxes: list[dict[str, Any]] = []
            for sh in prev_shapes:
                name = str(sh.get("name", ""))
                role = classify_morph_shape_prefix(name)
                if role in ("actor", "scene"):
                    prev_actors[name] = sh
                text = str(sh.get("text", "")).strip()
                if len(text) >= 6 and float(sh.get("left", 0.0)) < curr_w:
                    is_scene_kw = any(kw in name.lower() for kw in cls._SCENE_KEYWORDS)
                    if role == "slide_content" or not is_scene_kw:
                        prev_boxes.append(sh)

            for sh in curr_shapes:
                name = str(sh.get("name", ""))
                role = classify_morph_shape_prefix(name)
                sh_left = float(sh.get("left", 0.0))
                sh_top = float(sh.get("top", 0.0))

                # 5a) Actor 최소 변위 검증 (Δ < 5cm & 회전 < 15° & 크기 < 30%)
                if role == "actor" and name in prev_actors and sh_left < curr_w:
                    p = prev_actors[name]
                    dist_pt = math.hypot(sh_left - float(p.get("left", 0.0)), sh_top - float(p.get("top", 0.0)))
                    rot_delta = abs(float(sh.get("rotation", 0.0)) - float(p.get("rotation", 0.0))) % 360.0
                    pw = max(1.0, float(p.get("width", 1.0)))
                    ph = max(1.0, float(p.get("height", 1.0)))
                    w_ratio = abs(float(sh.get("width", pw)) - pw) / pw
                    h_ratio = abs(float(sh.get("height", ph)) - ph) / ph
                    if (
                        dist_pt < MIN_MORPH_DISPLACEMENT_PT
                        and rot_delta < MIN_MORPH_ROTATION_DEG
                        and max(w_ratio, h_ratio) < MIN_MORPH_SCALE_DELTA_RATIO
                    ):
                        warnings.append(
                            {
                                "slide": curr_idx,
                                "code": "insufficient_morph_delta",
                                "shape": name,
                                "message": (
                                    f"슬라이드 {curr_idx}의 '{name}' 변화량 미달 "
                                    f"(이동 {dist_pt/CM_TO_PT:.1f}cm < 5cm, 회전 {rot_delta:.1f}° < 15°, 크기 변화 {max(w_ratio,h_ratio)*100:.0f}% < 30%)."
                                ),
                            }
                        )

                # 5b) 동일 텍스트 + 동일 좌표 중복(고스팅 누락) 검증
                text = str(sh.get("text", "")).strip()
                if len(text) >= 6 and sh_left < curr_w and role != "scene":
                    for pb in prev_boxes:
                        if (
                            str(pb.get("text", "")).strip()[:50] == text[:50]
                            and abs(float(pb.get("left", 0.0)) - sh_left) < 1.0
                            and abs(float(pb.get("top", 0.0)) - sh_top) < 1.0
                        ):
                            warnings.append(
                                {
                                    "slide": curr_idx,
                                    "code": "duplicate_static_text",
                                    "shape": name,
                                    "message": f"슬라이드 {curr_idx - 1}→{curr_idx}에 동일 텍스트·동일 좌표 도형('{text[:25]}...')이 중복 존재합니다.",
                                }
                            )
                            break

        return {
            "passed": len(issues) == 0,
            "slides_checked": total_slides,
            "ghost_count": ghost_count,
            "ghost_limit": expected_max_ghosts,
            "issues_count": len(issues),
            "warnings_count": len(warnings),
            "issues": issues,
            "warnings": warnings,
        }


def _hex_luminance(hex_color: str) -> float:
    """HEX 색상(#RRGGBB) 상대 밝기(0.0~1.0) 계산."""
    try:
        clean = hex_color.strip().lstrip("#")
        if len(clean) == 3:
            clean = "".join(ch * 2 for ch in clean)
        r, g, b = int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16)
        return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0
    except Exception:
        return 1.0


def _read_presentation_morph_snapshot(prs: Any) -> list[dict[str, Any]]:
    """PowerPoint COM 프레젠테이션 객체에서 MorphAuditor용 경량 스냅샷 추출."""
    slide_w = 960.0
    slide_h = 540.0
    try:
        slide_w = float(prs.PageSetup.SlideWidth)
        slide_h = float(prs.PageSetup.SlideHeight)
    except Exception:
        pass

    slides_data: list[dict[str, Any]] = []
    for s_idx in range(1, prs.Slides.Count + 1):
        sl = prs.Slides(s_idx)
        eff = None
        try:
            eff = sl.SlideShowTransition.EntryEffect
        except Exception:
            pass

        shapes_list: list[dict[str, Any]] = []
        for sh_idx in range(1, sl.Shapes.Count + 1):
            try:
                sh = sl.Shapes(sh_idx)
                text = ""
                if getattr(sh, "HasTextFrame", 0) and getattr(sh.TextFrame, "HasText", 0):
                    text = (sh.TextFrame.TextRange.Text or "").replace("\r", " ").strip()
                shapes_list.append(
                    {
                        "name": str(getattr(sh, "Name", "")),
                        "left": float(getattr(sh, "Left", 0.0)),
                        "top": float(getattr(sh, "Top", 0.0)),
                        "width": float(getattr(sh, "Width", 0.0)),
                        "height": float(getattr(sh, "Height", 0.0)),
                        "rotation": float(getattr(sh, "Rotation", 0.0)),
                        "text": text,
                    }
                )
            except Exception:
                continue

        slides_data.append(
            {
                "slide_index": s_idx,
                "slide_width_pt": slide_w,
                "slide_height_pt": slide_h,
                "transition_effect": eff,
                "shapes": shapes_list,
            }
        )
    return slides_data


def register_ppt_morph_tools(mcp: MCPServer) -> None:
    """PowerPoint Morph 안무 및 Pitch Deck 플래너 MCP 도구 등록."""

    @mcp.tool()
    def ppt_clone_and_morph_slide(
        from_slide: int,
        ghost_prev_content: bool = True,
        duration_sec: float = 1.0,
        ghost_legacy_shapes: bool = False,
    ) -> dict[str, Any]:
        """
        지정한 PowerPoint 슬라이드를 복제한 뒤, Morph 전환(ppEffectMorph=3954)을 자동 적용하고
        이전 슬라이드의 고유 도형(`#s{from_slide}-*`)을 삭제 대신 캔버스 우측 바깥으로 고스팅합니다.

        Args:
            from_slide: 복제할 원본 슬라이드 번호 (1-based)
            ghost_prev_content: True이면 `#s{from_slide}-*` 도형을 우측 바깥(`1020.5pt` 또는 `SlideWidth+60pt`)으로 이동
            duration_sec: Morph 전환 재생 시간(초, 기본 1.0)
            ghost_legacy_shapes: True이면 일반 레거시 도구 도형(`!!scene-`/`!!actor-`/`#sN-` 미지정)도 우측 바깥으로 고스팅/정리
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        if from_slide < 1 or from_slide > prs.Slides.Count:
            return {"success": False, "error": f"유효하지 않은 슬라이드 번호입니다: {from_slide} (전체 {prs.Slides.Count}장)"}

        try:
            with ppt_freeze_display(ppt_app):
                src_slide = prs.Slides(from_slide)
                dup_range = src_slide.Duplicate()
                to_slide_idx = from_slide + 1
                try:
                    new_slide = dup_range.Item(1)
                except Exception:
                    new_slide = prs.Slides(to_slide_idx)
                try:
                    new_slide.MoveTo(to_slide_idx)
                except Exception:
                    pass

                # Morph 전환 적용
                trans = new_slide.SlideShowTransition
                trans.EntryEffect = PP_EFFECT_MORPH
                trans.Duration = float(duration_sec)

                slide_w = float(getattr(prs.PageSetup, "SlideWidth", 960.0))
                ghost_x = morph_ghost_left_pt(slide_w)
                ghosted_shapes: list[str] = []
                pruned_stale_ghosts: list[str] = []

                # 역순 순회: 이미 과거 슬라이드(K < from_slide)에서 오프캔버스 퇴장 완료된 유령 도형은 정리(O(N^2) 누적 방지)
                for i in range(new_slide.Shapes.Count, 0, -1):
                    sh = new_slide.Shapes(i)
                    sh_name = str(getattr(sh, "Name", ""))
                    pref_num = extract_slide_prefix_num(sh_name)
                    sh_role = classify_morph_shape_prefix(sh_name)
                    sh_left = float(getattr(sh, "Left", 0.0))
                    if pref_num is not None and pref_num < from_slide:
                        try:
                            if sh_left >= slide_w - 5.0:
                                sh.Delete()
                                pruned_stale_ghosts.append(sh_name)
                        except Exception:
                            pass
                    elif sh_role == "other" and sh_left >= slide_w - 5.0:
                        if ghost_legacy_shapes:
                            try:
                                sh.Delete()
                                pruned_stale_ghosts.append(sh_name)
                            except Exception:
                                pass
                    elif ghost_prev_content and (
                        sh_name.startswith(f"#s{from_slide}-")
                        or (ghost_legacy_shapes and sh_role == "other")
                    ):
                        sh.Left = ghost_x
                        ghosted_shapes.append(sh_name)

                ghosted_shapes.reverse()
                pruned_stale_ghosts.reverse()

                try:
                    ppt_app.ActiveWindow.View.GotoSlide(to_slide_idx)
                except Exception:
                    pass
                com_manager.activate_app("ppt")

            return {
                "success": True,
                "from_slide": from_slide,
                "to_slide": to_slide_idx,
                "transition": "morph",
                "duration_sec": float(duration_sec),
                "ghost_x_pt": ghost_x,
                "ghosted_shapes": ghosted_shapes,
                "pruned_stale_ghosts": pruned_stale_ghosts,
            }
        except Exception as e:
            return {"success": False, "error": f"Morph 슬라이드 복제 실패: {e}"}

    @mcp.tool()
    def ppt_upsert_morph_actor(
        slide_index: int,
        actor_name: str,
        shape_type: str = "rounded_rect",
        left_pt: float = 100.0,
        top_pt: float = 100.0,
        width_pt: float = 150.0,
        height_pt: float = 80.0,
        text: str = "",
        fill_hex: str = "",
        font_color: str = "",
        border_hex: str = "none",
        font_size_pt: float | None = None,
        bold: bool | None = None,
        font_name: str = "",
        text_align: str = "",
        z_order: str = "",
        opacity: float = 1.0,
        rotation: float = 0.0,
        ghost: bool = False,
    ) -> dict[str, Any]:
        """
        슬라이드 내 Morph 액터(`!!scene-*`, `!!actor-*`, `#sN-*`)를 이름 기준으로 찾아
        있으면 위치/크기/회전/투명도/텍스트를 갱신하고, 없으면 새로 생성합니다.
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        try:
            slide, resolved_idx = _resolve_target_slide(ppt_app, prs, slide_index=slide_index, create_new_slide=False)
            slide_w = float(getattr(prs.PageSetup, "SlideWidth", 960.0))
            target_left = morph_ghost_left_pt(slide_w) if ghost else float(left_pt)

            target_shape = None
            action = "updated"
            for i in range(1, slide.Shapes.Count + 1):
                sh = slide.Shapes(i)
                if getattr(sh, "Name", "") == actor_name:
                    target_shape = sh
                    break

            with ppt_freeze_display(ppt_app):
                if target_shape is None:
                    action = "created"
                    st_lower = shape_type.lower().strip()
                    if st_lower == "textbox":
                        target_shape = slide.Shapes.AddTextbox(1, target_left, float(top_pt), float(width_pt), float(height_pt))
                    else:
                        mso_id = MSO_SHAPE_MAP.get(st_lower, 5)
                        target_shape = slide.Shapes.AddShape(mso_id, target_left, float(top_pt), float(width_pt), float(height_pt))
                    target_shape.Name = actor_name
                else:
                    target_shape.Left = target_left
                    target_shape.Top = float(top_pt)
                    target_shape.Width = float(width_pt)
                    target_shape.Height = float(height_pt)

                target_shape.Rotation = float(rotation)

                clamped_opacity = max(0.0, min(1.0, float(opacity)))
                if fill_hex:
                    if fill_hex.strip().lower() in ("none", "transparent"):
                        target_shape.Fill.Visible = 0
                    else:
                        target_shape.Fill.Visible = -1
                        target_shape.Fill.Solid()
                        target_shape.Fill.ForeColor.RGB = hex_to_bgr(fill_hex)
                        target_shape.Fill.Transparency = 1.0 - clamped_opacity
                elif opacity != 1.0:
                    try:
                        target_shape.Fill.Transparency = 1.0 - clamped_opacity
                    except Exception:
                        pass

                if text:
                    target_shape.TextFrame.TextRange.Text = strip_markdown_syntax(text).replace("\n", "\r")

                # 외곽선 제어 (기본값 "none": AddShape 기본 진한 테두리 제거, #HEX 시 정교한 0.75pt 헤어라인)
                b_norm = (border_hex or "none").strip().lower()
                try:
                    if b_norm in ("none", "transparent"):
                        target_shape.Line.Visible = 0
                    elif b_norm != "keep":
                        target_shape.Line.Visible = -1
                        target_shape.Line.ForeColor.RGB = hex_to_bgr(border_hex)
                        target_shape.Line.Weight = 0.75
                except Exception:
                    pass

                # 둥근 사각형 모서리 반경 및 카드 내부 여백 · 한글 줄바꿈 제어 (생성/리사이즈 모두 재계산)
                if shape_type.lower().strip() != "textbox":
                    min_dim = max(1.0, min(float(width_pt), float(height_pt)))
                    if shape_type.lower().strip() == "rounded_rect":
                        try:
                            if getattr(target_shape.Adjustments, "Count", 0) >= 1:
                                target_shape.Adjustments[1] = round(min(0.45, 10.0 / min_dim), 4)
                        except Exception:
                            pass
                    try:
                        tf = target_shape.TextFrame
                        h_pad = 18.0 if min_dim > 100.0 else 14.0
                        v_pad = 12.0 if min_dim > 100.0 else 10.0
                        tf.MarginLeft = h_pad
                        tf.MarginRight = h_pad
                        tf.MarginTop = v_pad
                        tf.MarginBottom = v_pad
                        tf.WordWrap = -1
                    except Exception:
                        pass

                resolved_fc = font_color.strip()
                if not resolved_fc and fill_hex and fill_hex.strip().lower() not in ("none", "transparent"):
                    resolved_fc = "#1F2937" if _hex_luminance(fill_hex) > 0.55 else "#FFFFFF"
                try:
                    tr = target_shape.TextFrame.TextRange
                    if resolved_fc:
                        tr.Font.Color.RGB = hex_to_bgr(resolved_fc)
                    if font_size_pt is not None:
                        tr.Font.Size = float(font_size_pt)
                    if bold is not None:
                        tr.Font.Bold = -1 if bold else 0
                    if font_name.strip():
                        fn_clean = font_name.strip()
                        tr.Font.Name = fn_clean
                        try:
                            tr.Font.NameFarEast = fn_clean
                        except Exception:
                            pass
                    al_key = text_align.strip().lower()
                    if al_key in ALIGN_MAP:
                        tr.ParagraphFormat.Alignment = ALIGN_MAP[al_key]
                    try:
                        tr.ParagraphFormat.FarEastLineBreakControl = -1
                        tr.ParagraphFormat.LineRuleWithin = -1
                        tr.ParagraphFormat.SpaceWithin = 1.15
                    except Exception:
                        pass
                except Exception:
                    pass

                z_key = (z_order or "").strip().lower()
                if z_key in ("front", "bring_to_front"):
                    try:
                        target_shape.ZOrder(0)
                    except Exception:
                        pass
                elif z_key in ("back", "send_to_back"):
                    try:
                        target_shape.ZOrder(1)
                    except Exception:
                        pass

            com_manager.activate_app("ppt")
            return {
                "success": True,
                "action": action,
                "slide_index": resolved_idx,
                "actor_name": actor_name,
                "prefix_role": classify_morph_shape_prefix(actor_name),
                "ghosted": ghost,
                "left_pt": target_left,
            }
        except Exception as e:
            return {"success": False, "error": f"Morph 액터 Upsert 실패: {e}"}

    @mcp.tool()
    def ppt_verify_morph_deck() -> dict[str, Any]:
        """
        현재 활성화된 PowerPoint 프레젠테이션 전체의 Morph 설정(전환 누락, `#sN-` 미고스팅 잔존,
        M-2 고스트 과누적, 인접 슬라이드 동일 텍스트 중복, 액터 최소 변위 5cm 미달)을 검증합니다.
        """
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if ppt_app is None or prs is None:
            return {"success": False, "error": "열려 있는 PowerPoint 프레젠테이션이 없습니다."}

        try:
            snapshot = _read_presentation_morph_snapshot(prs)
            report = MorphAuditor.audit_deck(snapshot)
            report["success"] = True
            report["presentation_name"] = prs.Name
            return report
        except Exception as e:
            return {"success": False, "error": f"Morph 덱 검증 실패: {e}"}

    @mcp.tool()
    def ppt_plan_pitch_deck(
        stage: str,
        vertical: str,
        company: str = "Company",
        raise_amount: str = "",
        slides_count: int | None = None,
    ) -> dict[str, Any]:
        """
        투자 단계(Seed, Series A/B/C, Bridge) 및 산업 버티컬(B2B SaaS, Consumer, DeepTech, Marketplace, Bio)에 맞춰
        iOfficeAI Pitch Deck 표준 아크 및 C.1~C.5b 레이아웃 청사진을 생성합니다.
        """
        return PitchDeckPlanner.plan_deck(
            stage=stage,
            vertical=vertical,
            company=company,
            raise_amount=raise_amount,
            slides_override=slides_count,
        )
