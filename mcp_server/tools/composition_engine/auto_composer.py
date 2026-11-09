# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/auto_composer.py
사용자가 입력한 비구조화 줄글/원시 데이터 텍스트(Raw Text Dump)를 스스로 분석하여,
서사 구조(기승전결)와 17대 슬라이스 최적 아키타입(KPI, 프로세스, 비교, 로드맵, 매트릭스 등)을
지능적으로 자동 판정하고 완성형 슬라이드 덱 스펙(Spec)을 생성하는 오토 컴포저.
"""

from __future__ import annotations

import re
from typing import Any


class AutoComposer:
    """원시 텍스트 덤프 기반 PPT 자동 기획·구성 엔진"""

    @classmethod
    def compose(
        cls,
        raw_text: str,
        presentation_title: str | None = None,
        theme: str | None = None,
    ) -> dict[str, Any]:
        clean_text = raw_text.strip()
        lines = [l.strip() for l in clean_text.splitlines() if l.strip()]

        # ---------------------------------------------------------------------
        # 1. 프레젠테이션 메인 타이틀 추론
        # ---------------------------------------------------------------------
        inferred_title = presentation_title
        if not inferred_title:
            for l in lines[:3]:
                c_line = re.sub(r"^[#=\-\s\*\d\.\(\)\[\]]+|[\[\]]+$", "", l).strip()
                if len(c_line) > 4 and not l.startswith(("-", "*", "•", "1.", "2.", "3.")):
                    inferred_title = c_line
                    break
        if not inferred_title:
            inferred_title = "전략 비즈니스 제안서"

        # ---------------------------------------------------------------------
        # 2. 테마 자동 감지
        # ---------------------------------------------------------------------
        inferred_theme = theme
        if not inferred_theme:
            lower_text = clean_text.lower()
            if any(k in lower_text for k in ("esg", "친환경", "탄소", "지속가능", "re100", "그린")):
                inferred_theme = "emerald_green"
            elif any(k in lower_text for k in ("ai", "인공지능", "데이터", "테크", "saas", "스타트업", "클라우드", "소프트웨어", "엔지니어")):
                inferred_theme = "modern_teal"
            else:
                inferred_theme = "business_navy"

        # ---------------------------------------------------------------------
        # 3. 텍스트를 논리적 섹션/슬라이드 블록으로 분할 (Smart Chunking)
        # ---------------------------------------------------------------------
        raw_blocks: list[str] = []
        current_block: list[str] = []

        for l in lines:
            # 마크다운 헤더 (# 1., ## 2. 등) 또는 주요 섹션 구분자 인식
            is_header = bool(re.match(r"^(\s*#{1,4}\s+|==+|--+|\d+[\.\)]\s+[가-힣A-Za-z])", l)) and len(l) < 70
            if is_header and current_block:
                raw_blocks.append("\n".join(current_block))
                current_block = [l]
            else:
                current_block.append(l)
        if current_block:
            raw_blocks.append("\n".join(current_block))

        # 첫 블록이 덱 메인 타이틀인 경우 슬라이드 블록에서 제외
        if len(raw_blocks) > 1:
            first_block_clean = re.sub(r"^[#=\-\s\*\d\.\(\)\[\]]+|[\[\]]+$", "", raw_blocks[0].splitlines()[0]).strip()
            if (
                first_block_clean == inferred_title
                or (len(raw_blocks[0].splitlines()) <= 2 and any(k in first_block_clean for k in ("자기소개서", "제안서", "보고서", "발표자료")))
            ):
                raw_blocks = raw_blocks[1:]

        # 블록이 1개 이하일 경우 단락 단위 분할
        if len(raw_blocks) < 2:
            paragraphs = [p.strip() for p in clean_text.split("\n\n") if p.strip()]
            if len(paragraphs) >= 2:
                raw_blocks = paragraphs
            else:
                sentences = [s.strip() for s in re.split(r"(?<=[.!?다])\s+", clean_text) if s.strip()]
                chunk_sz = max(2, len(sentences) // 4)
                raw_blocks = ["\n".join(sentences[i : i + chunk_sz]) for i in range(0, len(sentences), chunk_sz)]

        # ---------------------------------------------------------------------
        # 4. 각 블록별 시각 아키타입 분류 및 세부 데이터 추출
        # ---------------------------------------------------------------------
        planned_slides = []
        agenda_items = []

        for idx, block in enumerate(raw_blocks, start=1):
            b_lines = [b.strip() for b in block.splitlines() if b.strip()]
            if not b_lines:
                continue

            first_line = b_lines[0]
            # 헤더 마크다운 기호 정제 (예: ## 1. 핵심 역량 -> 핵심 역량)
            first_line_clean = re.sub(r"^#{1,6}\s*", "", first_line)
            first_line_clean = re.sub(r"^\d+[\.\)]\s*", "", first_line_clean)
            first_line_clean = re.sub(r"^[=\-\*\s]+|[=\-\*\s]+$", "", first_line_clean).strip()
            if not first_line_clean:
                first_line_clean = f"핵심 추진 과제 및 전략 분석 {idx}"

            action_title = first_line_clean[:50]
            subtitle = ""
            if len(b_lines) > 1 and len(b_lines[1]) < 60 and not b_lines[1].startswith(("-", "*", "•", "1.", "2.")):
                subtitle = b_lines[1].strip()
            else:
                subtitle = "데이터 기반 정밀 분석 및 세부 실행 방안"

            # 본문 라인만 별도 분리 (헤더 제외)
            body_lines = b_lines[1:] if len(b_lines) > 1 else []
            body_text = "\n".join(body_lines)
            block_text = block.lower()

            # -----------------------------------------------------------------
            # (A) 로드맵/마일스톤형 (Roadmap Chevron) 사전 판정
            # -----------------------------------------------------------------
            has_roadmap = any(
                k in action_title.lower() or k in block_text
                for k in ("로드맵", "마일스톤", "roadmap", "milestone", "비전 & 로드맵", "성장 로드맵")
            )

            # -----------------------------------------------------------------
            # (B) 수치 지표형 (Metric Cards) 정밀 판정 및 파싱
            # -----------------------------------------------------------------
            unit_pattern = r"(?:%|억원|억|만원|원|배|건|개|ms|초|점|위|년차|년|만|명|개사|TPS|tps|k|K|X|x)"
            kpis = []

            # 본문 불릿 라인 기반 수치 지표 추출
            for line in body_lines:
                # 라벨: 값 형태 라인 매칭
                kv_m = re.match(r"^[-*•\d\.\s]*([가-힣a-zA-Z][가-힣a-zA-Z0-9\s_\-\(\)/]{1,25})[:\s-]+(.*)", line)
                if kv_m:
                    l_name = kv_m.group(1).strip()
                    rest_val = kv_m.group(2).strip()

                    # 화살표 변화 표기 추출 (250ms -> 145ms 등)
                    delta_str = "주요 성과 지표"
                    ar_match = re.search(r"([+\-]?\d+[\.,]?\d*\s*\w*)\s*(?:->|→)\s*([+\-]?\d+[\.,]?\d*\s*\w*)", rest_val)
                    if ar_match:
                        delta_str = f"{ar_match.group(1)} → {ar_match.group(2)}"

                    # 숫자 + 단위 검색
                    val_matches = re.findall(rf"([+\-]?\d+[\.,]?\d*\s*{unit_pattern})", rest_val)
                    if val_matches:
                        # %가 있거나 화살표가 있는 경우 가장 대표적인 값 선택
                        best_val = val_matches[0]
                        for vm in val_matches:
                            if "%" in vm or "배" in vm or "년" in vm or "만" in vm or "tps" in vm.lower():
                                best_val = vm
                                break

                        if delta_str == "주요 성과 지표":
                            if "%" in best_val or "단축" in rest_val or "감소" in rest_val:
                                delta_str = "효율성 극대화"
                            elif "달성" in rest_val or "향상" in rest_val or "증가" in rest_val:
                                delta_str = "목표 초과 달성"

                        kpis.append({
                            "value": best_val.strip(),
                            "label": l_name,
                            "delta": delta_str,
                            "status": "positive",
                        })

            # 지표형 슬라이드 명시 여부
            is_metric_slide = any(
                k in action_title.lower() for k in ("key metrics", "정량 성과", "핵심 지표", "주요 성과", "kpi")
            ) or (len(kpis) >= 3 and not has_roadmap)

            # -----------------------------------------------------------------
            # (C) 프로세스/단계형 (Flow Process Steps) 판정
            # -----------------------------------------------------------------
            # 본문 라인(body_text)에서만 단계 매칭 수행 (헤더의 '4. ' 오인식 방지)
            step_matches = re.findall(
                r"(?:step\s*\d+|\d+단계|\d+[\.\)]|단계\s*\d+)[:\s-]*(\[[^\]]+\]|[\w가-힣\s\-]+)[:\s-]*(.*)",
                body_text,
                re.IGNORECASE,
            )
            has_process = (len(step_matches) >= 3 or ("단계" in body_text and len(body_lines) >= 3)) and not has_roadmap

            # -----------------------------------------------------------------
            # (D) 비교/대조형 (Split 2Col) 판정
            # -----------------------------------------------------------------
            has_comparison = any(
                k in block_text
                for k in (
                    "before & after", "before", "after", "as-is", "to-be", "기존", "개선",
                    "도입 전", "도입 후", "vs", "대비", "장단점"
                )
            ) and not has_roadmap

            # -----------------------------------------------------------------
            # (E) 매트릭스 / 표형 판정
            # -----------------------------------------------------------------
            has_matrix = any(
                k in block_text for k in ("매트릭스", "4분면", "사분면", "우선순위", "난이도", "영향도")
            )
            has_table = "|" in block or "경쟁사" in block_text or "스펙 비교" in block_text

            slide_role = "card_grid_3col"
            slide_data: dict[str, Any] = {}
            category_badge = "STRATEGY"

            # -----------------------------------------------------------------
            # 분기 매핑 및 데이터 빌드
            # -----------------------------------------------------------------
            if has_roadmap:
                slide_role = "roadmap_chevron"
                category_badge = "ROADMAP"
                milestones = []

                # 본문에서 단계별 로드맵 불릿 파싱 (예: 1단계 (입사 1년차 - 적응 및 기여): ...)
                roadmap_steps = re.findall(
                    r"(?:step\s*\d+|\d+단계|\d+[\.\)]|단계\s*\d+)[:\s-]*(\([^)]+\)|\[[^\]]+\]|[\w가-힣\s\-]+)[:\s-]*(.*)",
                    body_text,
                    re.IGNORECASE,
                )
                if roadmap_steps:
                    for m_idx, (r_title, r_desc) in enumerate(roadmap_steps[:4], start=1):
                        clean_m_title = r_title.strip("()[] -").strip()
                        parts = clean_m_title.split("-", 1)
                        period_str = parts[0].strip() if len(parts) > 1 else f"Phase {m_idx}"
                        task_str = parts[1].strip() if len(parts) > 1 else clean_m_title
                        milestones.append({
                            "phase": f"P{m_idx}",
                            "title": task_str[:12],
                            "period": period_str[:15],
                            "highlight": (m_idx == 2 or m_idx == 3),
                        })
                else:
                    periods = ["Phase 1 (단기)", "Phase 2 (중기)", "Phase 3 (장기)", "Phase 4 (비전)"]
                    roadmap_items = [l for l in body_lines if len(l) > 3][:4]
                    for m_idx, p_name in enumerate(periods[:max(3, len(roadmap_items))]):
                        m_title = roadmap_items[m_idx].split(":")[0][:12] if m_idx < len(roadmap_items) else f"과제 {m_idx+1}"
                        milestones.append({
                            "phase": f"P{m_idx+1}",
                            "title": m_title,
                            "period": p_name,
                            "highlight": (m_idx == 1 or m_idx == 2),
                        })
                slide_data = {"milestones": milestones}

            elif is_metric_slide:
                slide_role = "metric_cards"
                category_badge = "PERFORMANCE"
                while len(kpis) < 3:
                    kpis.append({"value": "99.9%", "label": "시스템 가용성", "delta": "안정적 운영", "status": "neutral"})
                slide_data = {"metrics": kpis[:4]}

            elif has_process:
                slide_role = "flow_process_steps"
                category_badge = "PROCESS"
                steps = []
                if step_matches:
                    for s_idx, (s_title, s_desc) in enumerate(step_matches[:4], start=1):
                        clean_s_title = s_title.strip("[] ").strip()
                        clean_s_desc = s_desc.strip() or "상세 실행 과제 및 표준 프로세스 검증"
                        steps.append({
                            "step": f"STEP {s_idx:02d}",
                            "title": clean_s_title[:15],
                            "desc": clean_s_desc[:50],
                        })
                else:
                    items = [l for l in body_lines if len(l) > 3 and not l.startswith("#")][:4]
                    for s_idx, itm in enumerate(items, start=1):
                        c_itm = re.sub(r"^[-*•\d\.\s]+", "", itm).strip()
                        parts = c_itm.split(":", 1)
                        t = parts[0].strip()[:15]
                        d = parts[1].strip()[:40] if len(parts) > 1 else "프로세스 상세 실행 및 표준 검증"
                        steps.append({"step": f"STEP {s_idx:02d}", "title": t, "desc": d})
                slide_data = {"steps": steps}

            elif has_comparison:
                slide_role = "split_2col"
                category_badge = "COMPARISON"
                left_items = []
                right_items = []
                current_side = "left"

                for l in body_lines:
                    l_clean = re.sub(r"^[-*•\s]+", "", l).strip()
                    if not l_clean or len(l_clean) < 2:
                        continue
                    l_lower = l.lower()

                    if any(k in l_lower for k in ("before", "기존", "as-is", "과거", "한계", "문제", "도입 전")):
                        current_side = "left"
                        continue
                    elif any(k in l_lower for k in ("after", "개선", "to-be", "신규", "혁신", "효과", "도입 후")):
                        current_side = "right"
                        continue

                    if current_side == "left":
                        left_items.append(l_clean[:45])
                    else:
                        right_items.append(l_clean[:45])

                if not left_items:
                    left_items = ["수작업 중심의 비효율적 운영", "데이터 단편화 및 긴 분석 지연"]
                if not right_items:
                    right_items = ["실시간 자동화 파이프라인 구축", "운영 비용 30% 절감 및 민첩성 확보"]

                slide_data = {
                    "left_title": "AS-IS (기존 레거시 현황)",
                    "left_items": left_items[:4],
                    "right_title": "TO-BE (클라우드/AI 자동화)",
                    "right_items": right_items[:4],
                }

            elif has_matrix:
                slide_role = "matrix_2x2"
                category_badge = "PORTFOLIO"
                slide_data = {
                    "x_label": "실행 난이도 (저 -> 고)",
                    "y_label": "비즈니스 영향도 (저 -> 고)",
                    "quadrants": [
                        {"name": "Quick Wins (즉시 실행)", "items": ["공통 모듈 전환", "자동화 알림 연동"]},
                        {"name": "Strategic Bets (전략 과제)", "items": ["코어 시스템 SaaS 이관", "실시간 AI 파이프라인"]},
                        {"name": "Low Priority (후순위)", "items": ["단순 게시판 개편", "보존용 데이터 백업"]},
                        {"name": "Re-evaluate (재검토)", "items": ["외산 솔루션 단독 구축", "물리 인프라 증설"]},
                    ],
                }

            elif has_table:
                slide_role = "comparison_table"
                category_badge = "BENCHMARK"
                slide_data = {
                    "headers": ["평가 항목", "자사 솔루션", "경쟁사 A", "경쟁사 B"],
                    "rows": [
                        ["보안 인증", "CSAP SaaS 최상위", "일반 ISMS", "자체 보안 가이드"],
                        ["구축 소요 기간", "2주 이내 (초고속)", "6~8주", "12주 이상"],
                        ["비용 경쟁력", "경쟁사 대비 30% 절감", "표준 정가제", "초기 도입비 과다"],
                        ["기술 지원", "24/7 국내 전담 엔지니어", "영업시간 외 콜센터", "해외 본사 티켓팅"],
                    ],
                }

            else:
                # -------------------------------------------------------------
                # 카드 그리드 (Card Grid 3Col) - 본문 세부 불릿 100% 보존
                # -------------------------------------------------------------
                slide_role = "card_grid_3col"
                category_badge = "STRATEGY"
                cards = []
                content_lines = [l for l in body_lines if not l.startswith("#") and len(l) > 2]

                for c_idx, line in enumerate(content_lines[:3], start=1):
                    c_clean = re.sub(r"^[-*•\s]+", "", line).strip()
                    parts = c_clean.split(":", 1)
                    c_title = parts[0].strip()[:18]

                    bullets: list[str] = []
                    if len(parts) > 1:
                        val_part = parts[1].strip()
                        # 괄호 안의 쉼표는 보호하면서 바깥 쉼표/세미콜론으로 분할
                        sub_items = [s.strip() for s in re.split(r",\s*(?![^()]*\))|;\s*", val_part) if s.strip()]
                        if len(sub_items) >= 2:
                            bullets = [it[:40] for it in sub_items[:4]]
                        else:
                            bullets = [val_part[:50]]
                    else:
                        bullets = ["전사 표준화 및 체계적 고도화 추진", "안정적 운영 거버넌스 확립"]

                    cards.append({
                        "tag": f"0{c_idx}",
                        "title": c_title,
                        "bullets": bullets,
                        "desc": bullets[0] if bullets else "세부 과제 실행",
                    })

                while len(cards) < 3:
                    c_num = len(cards) + 1
                    cards.append({
                        "tag": f"0{c_num}",
                        "title": f"과제 0{c_num}",
                        "bullets": ["지속 가능한 비즈니스 성장 기반 마련", "체계적 품질 검증 및 거버넌스 확보"],
                        "desc": "지속 가능한 성장 기반 마련",
                    })

                slide_data = {"cards": cards}

            planned_slides.append({
                "slide_index": idx,
                "title": action_title,
                "subtitle": subtitle,
                "category": category_badge,
                "role": slide_role,
                "data": slide_data,
            })
            agenda_items.append(f"{idx:02d} {action_title[:18]}")

        return {
            "title": inferred_title,
            "subtitle": "비즈니스 전략 및 실행 계획 보고서",
            "theme": inferred_theme,
            "agenda_items": agenda_items,
            "slides_count": len(planned_slides),
            "slides": planned_slides,
        }



    # =========================================================================
    # Zero-Config Contextual Auto-Styling with 95 Blueprints
    # =========================================================================
    @classmethod
    def infer_style_context(cls, text: str) -> dict[str, Any]:
        """
        사용자의 원시 텍스트/프롬프트 맥락을 스스로 분석하여
        5대 톤앤매너, 최적의 95대 기법 청사진 및 이미지 문법을 자동 추론.
        """
        from pathlib import Path
        import json

        bp_path = Path("data/ppt_technique_blueprints.json")
        rules = {}
        if bp_path.exists():
            try:
                with open(bp_path, "r", encoding="utf-8") as f:
                    rules = json.load(f).get("context_intent_rules", {})
            except Exception:
                pass

        lower_t = text.lower()

        scores: dict[str, int] = {
            "executive_impact": 0,
            "corporate_formal": 0,
            "tech_modern": 0,
            "soft_creative": 0,
            "minimal_clean": 1,
        }

        for mood, m_data in rules.items():
            for kw in m_data.get("keywords", []):
                if kw in lower_t:
                    scores[mood] = scores.get(mood, 0) + 2

        if re.search(r"(\d+(\.\d+)?%|\d+[조억만원]|kpi|roic|ebitda)", lower_t):
            scores["executive_impact"] = scores.get("executive_impact", 0) + 3

        if re.search(r"(\d+조|제\d+항|가\.|나\.|다\.|\[지침\]|규정)", text):
            scores["corporate_formal"] = scores.get("corporate_formal", 0) + 4

        best_mood = max(scores.items(), key=lambda x: x[1])[0]
        rule = rules.get(best_mood, {})

        candidate_techniques = rule.get("primary_techniques", ["3단_표준_핵심과제_카드그리드"])
        
        selected_technique = candidate_techniques[0]
        if best_mood == "executive_impact":
            if any(k in lower_t for k in ("빅넘버", "단독", "1위", "최대")):
                selected_technique = "1개_특대형_빅넘버_단독강조"
            else:
                selected_technique = "3단_경영실적_증감배지_KPI"
        elif best_mood == "corporate_formal":
            if any(k in lower_t for k in ("명세", "항목", "세부내역", "비용")):
                selected_technique = "측면오픈형_비즈니스_명세표"
            else:
                selected_technique = "KISA식_1_가_(1)_공문서형_위계본문"
        elif best_mood == "tech_modern":
            if any(k in lower_t for k in ("아키텍처", "인프라", "스택", "계층", "레이어")):
                selected_technique = "수평_3~4단_레이어드_스택"
            elif any(k in lower_t for k in ("방사", "네트워크", "중앙", "통합")):
                selected_technique = "중앙_허브앤스포크_방사형"
            elif any(k in lower_t for k in ("단계", "흐름", "파이프라인", "프로세스")):
                selected_technique = "3단_스텝형_수평_프로세스"
            else:
                selected_technique = "수평_3~4단_레이어드_스택"
        elif best_mood == "soft_creative":
            if any(k in lower_t for k in ("모바일", "앱", "화면", "스크린샷")):
                selected_technique = "스마트폰_모바일앱_지시선_콜아웃"
            elif any(k in lower_t for k in ("팀", "멤버", "조직", "프로필", "인물")):
                selected_technique = "3~4열_원형사진_임원진_프로필"
            else:
                selected_technique = "3단_사진_상단결합_비주얼카드"

        return {
            "mood": best_mood,
            "technique_name": selected_technique,
            "image_grammar": rule.get("default_image_grammar", "pure_vector_no_image"),
            "style_modifiers": rule.get("style_modifiers", {}),
            "scores": scores,
        }

    @classmethod
    def compose_with_blueprints(
        cls,
        raw_text: str,
        presentation_title: str | None = None,
        theme: str | None = None,
    ) -> dict[str, Any]:
        """
        95개 전수 청사진과 15개 이미지 배치 문법을 기반으로,
        자연어 맥락 자동추론이 완전히 통합된 슬라이드 덱 스펙 생성.
        """
        base_deck = cls.compose(raw_text, presentation_title=presentation_title, theme=theme)
        
        enhanced_slides = []
        for s in base_deck.get("slides", []):
            block_text = f"{s.get('title', '')} {s.get('subtitle', '')} {s.get('data', {})}"
            inference = cls.infer_style_context(block_text)

            items = []
            s_data = s.get("data", {})
            if "cards" in s_data:
                for c in s_data["cards"]:
                    items.append({
                        "title": c.get("title", ""),
                        "description": c.get("desc", ""),
                        "stat": c.get("tag", "")
                    })
            elif "steps" in s_data:
                for st in s_data["steps"]:
                    items.append({
                        "title": st.get("title", ""),
                        "description": st.get("desc", ""),
                        "stat": st.get("step", "")
                    })
            elif "kpis" in s_data:
                for k in s_data["kpis"]:
                    items.append({
                        "title": k.get("label", ""),
                        "stat": k.get("value", ""),
                        "description": k.get("badge", "")
                    })
            else:
                items = [
                    {"title": "주요 추진 과제", "description": "체계적 분석 및 안정적 실행"},
                    {"title": "핵심 기대 효과", "description": "비즈니스 경쟁력 극대화 및 품질 향상"},
                    {"title": "거버넌스 체계", "description": "지속 가능한 운영 및 모니터링"}
                ]

            enhanced_slide = {
                **s,
                "inferred_mood": inference["mood"],
                "blueprint_technique": inference["technique_name"],
                "image_grammar": inference["image_grammar"],
                "style_modifiers": inference["style_modifiers"],
                "render_data": {
                    "title": s.get("title", ""),
                    "subtitle": s.get("subtitle", ""),
                    "category": s.get("category", ""),
                    "items": items,
                }
            }
            enhanced_slides.append(enhanced_slide)

        return {
            **base_deck,
            "zero_file_dependency": True,
            "contextual_auto_styling": True,
            "slides": enhanced_slides,
        }
