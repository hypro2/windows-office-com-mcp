# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/auto_storyboarder.py
분해된 시맨틱 성분(Metrics, Chronology, Duality, Directives)을 바탕으로
4대 오피스별 완결된 스토리라인과 페이지/슬라이드 구성을 자율 빌드하는 스토리보더.
"""

from __future__ import annotations

import re
from typing import Any

from .narrative_planner import NarrativePlanner


class AutoStoryboarder:
    """원시 덤프 기반 오피스 문서 자율 스토리보더"""

    # -------------------------------------------------------------------------
    # 1. PPT 자율 스토리보딩 (3~5장 덱 빌드)
    # -------------------------------------------------------------------------
    @classmethod
    def storyboard_ppt(cls, topic: str, decomposed: dict[str, Any]) -> dict[str, Any]:
        """분해된 성분으로 기-승-전-결 PPT 슬라이드 덱 자동 합성"""
        slides = []
        metrics = decomposed.get("metrics", [])
        chronology = decomposed.get("chronology", [])
        duality = decomposed.get("duality", [])
        directives = decomposed.get("directives", [])
        sentences = decomposed.get("raw_sentences", [])

        # Slide 1: 핵심 성과 및 지표 요약 (kpi_grid_3)
        lead_metric_title = f"{topic}의 주요 핵심 지표는 목표 대비 초과 달성함"
        if directives:
            first_dir = directives[0]
            if len(first_dir) > 10:
                lead_metric_title = first_dir[:60]

        kpi_items = []
        if len(metrics) >= 3:
            for m in metrics[:3]:
                kpi_items.append({
                    "label": m["label"],
                    "value": m["value"],
                    "desc": m.get("context", "목표 대비 정상 추진 중")[:40]
                })
        else:
            # 수치가 부족한 경우 텍스트에서 3개 요약 포인트 생성
            for i in range(3):
                s_text = sentences[i] if i < len(sentences) else f"핵심 전략 과제 0{i+1}"
                kpi_items.append({
                    "label": f"주요 성과 0{i+1}",
                    "value": f"{100 + (i*15)}%",
                    "desc": s_text[:40]
                })

        slides.append({
            "archetype": "kpi_grid_3",
            "category": "경영 실적",
            "action_title": lead_metric_title,
            "items": kpi_items,
            "footnote": "자료: 사내 내부 집계 잠정치 (2026)",
        })

        # Slide 2: 현행 대비 전략 전환 대조 (comparison_2col)
        if duality or len(sentences) >= 4:
            contrast_title = f"{topic} 추진을 통해 현행 프로세스의 비효율을 획기적으로 개선함"
            as_is_desc = "단일 채널 의존 및 수작업 반복으로 인한 리드타임 지연 및 효율 저하"
            to_be_desc = "다변화 채널 및 자동화 파이프라인 도입을 통한 생산성 85% 극대화"
            if duality:
                to_be_desc = duality[0]["sentence"][:60]

            slides.append({
                "archetype": "comparison_2col",
                "category": "전략 전환",
                "action_title": contrast_title,
                "items": [
                    {
                        "title": "현행 방식 (As-Is)",
                        "desc": as_is_desc,
                        "badge": "한계점"
                    },
                    {
                        "title": "개선 방향 (To-Be)",
                        "desc": to_be_desc,
                        "badge": "혁신안"
                    }
                ],
                "footnote": "전략기획팀 프로세스 혁신안 기준",
            })

        # Slide 3: 세부 추진 로드맵 (process_horizontal_4)
        if chronology or len(sentences) >= 3:
            roadmap_title = f"{topic}의 연간 마일스톤을 4단계로 수립하여 차질 없이 집행함"
            step_items = []
            if len(chronology) >= 4:
                for c in chronology[:4]:
                    step_items.append({
                        "title": c.get("tag", "단계"),
                        "desc": c.get("desc", "세부 추진 계획 수립")[:45]
                    })
            else:
                default_steps = [
                    ("10월 기획", "요구사항 분석 및 환경 검증"),
                    ("11월 구축", "핵심 모듈 개발 및 A/B 테스트"),
                    ("12월 전개", "전사 확대 적용 및 교육"),
                    ("01월 고도화", "안정화 및 성과 모니터링"),
                ]
                for st, sd in default_steps:
                    step_items.append({"title": st, "desc": sd})

            slides.append({
                "archetype": "process_horizontal_4",
                "category": "실행 계획",
                "action_title": roadmap_title,
                "items": step_items,
                "footnote": "세부 일정은 유관부서 협의에 따라 조정 가능함",
            })

        return NarrativePlanner.plan_ppt(topic=topic, slides_input=slides)

    # -------------------------------------------------------------------------
    # 2. Excel 자율 스토리보딩 (2D 대시보드 토폴로지)
    # -------------------------------------------------------------------------
    @classmethod
    def storyboard_excel(cls, topic: str, decomposed: dict[str, Any]) -> dict[str, Any]:
        """비정형 데이터에서 상단 KPI 타일과 정규 그리드 자동 추출"""
        metrics = decomposed.get("metrics", [])
        sentences = decomposed.get("raw_sentences", [])

        # 상단 KPI 타일 3개
        kpi_tiles = []
        if metrics:
            for m in metrics[:3]:
                kpi_tiles.append({"label": m["label"], "value": m["value"]})
        while len(kpi_tiles) < 3:
            kpi_tiles.append({"label": f"지표 0{len(kpi_tiles)+1}", "value": "100.0%"})

        # 테이블 헤더 및 데이터 행 파싱
        headers = ["구분", "목표액", "실적액", "달성률"]
        data_rows = []

        # 문장에서 팀/부서별 수치 추출 시도
        team_pattern = re.compile(r"(\d팀|[가-힣A-Za-z]+팀|[가-힣]+부)\s*.*?(\d+(?:,\d+)*(?:\.\d+)?).*?(\d+(?:,\d+)*(?:\.\d+)?)")
        found_rows = False
        for s in sentences:
            m = team_pattern.search(s)
            if m:
                team_name = m.group(1)
                try:
                    target_val = float(m.group(2).replace(",", ""))
                    actual_val = float(m.group(3).replace(",", ""))
                    rate = round(actual_val / target_val, 2) if target_val > 0 else 1.0
                    data_rows.append([team_name, target_val, actual_val, rate])
                    found_rows = True
                except Exception:
                    pass

        if not found_rows or len(data_rows) == 0:
            data_rows = [
                ["사업1본부", 5000, 5400, 1.08],
                ["사업2본부", 4000, 3800, 0.95],
                ["사업3본부", 3000, 3300, 1.10],
            ]

        summary_formulas = {
            "합계": f"=SUM(C8:C{7 + len(data_rows)})"
        }

        return NarrativePlanner.plan_excel(
            title=f"{topic} 집계 대시보드",
            kpi_tiles=kpi_tiles,
            headers=headers,
            data_rows=data_rows,
            summary_formulas=summary_formulas,
            unit="단위: 백만원, %",
        )

    # -------------------------------------------------------------------------
    # 3. HWP 자율 스토리보딩 (A4 1장 개조식 요약)
    # -------------------------------------------------------------------------
    @classmethod
    def storyboard_hwp(cls, topic: str, decomposed: dict[str, Any]) -> dict[str, Any]:
        """A4 1장 28줄 상한제에 맞춘 표준 공문서 개조식 스토리보드"""
        directives = decomposed.get("directives", [])
        sentences = decomposed.get("raw_sentences", [])

        # 상단 요약 상자 (최대 2줄)
        sum_box = []
        if directives:
            sum_box.append(directives[0][:50])
        if len(directives) > 1:
            sum_box.append(directives[1][:50])
        elif len(sentences) > 1:
            sum_box.append(sentences[1][:50])
        else:
            sum_box.append(f"{topic} 추진 안건 검토 및 예산 승인 요청 건임.")

        # 본문 개조식 3대 섹션
        sections = [
            {
                "title": "추진 배경 및 필요성",
                "items": [
                    {"level": 2, "symbol": "○", "text": sentences[0][:40] if sentences else "시장 환경 대응을 위한 선제적 체계 개편 필요"},
                    {"level": 3, "symbol": "-", "text": "현행 수작업 방식의 비효율 및 오류 지속 발생에 따른 자동화 시급"},
                ]
            },
            {
                "title": "주요 추진 현황 및 대책",
                "items": [
                    {"level": 2, "symbol": "○", "text": sentences[1][:40] if len(sentences) > 1 else "핵심 KPI 목표 초과 달성 및 채널 다변화 추진"},
                    {"level": 3, "symbol": "-", "text": "세부 일정별 단계적 검증 및 유관부서 협업 체계 가동"},
                ]
            },
            {
                "title": "향후 실행 계획",
                "items": [
                    {"level": 2, "symbol": "○", "text": "연내 파일럿 론칭 및 정량적 성과 모니터링 지속"},
                ]
            }
        ]

        return NarrativePlanner.plan_hwp_one_page(
            title=topic,
            executive_summary=sum_box,
            sections=sections,
        )

    # -------------------------------------------------------------------------
    # 4. DOCX 자율 스토리보딩 (4대 비즈니스 블록)
    # -------------------------------------------------------------------------
    @classmethod
    def storyboard_docx(cls, topic: str, decomposed: dict[str, Any]) -> dict[str, Any]:
        """워드 품의·기획 4대 표준 블록 스토리보드"""
        directives = decomposed.get("directives", [])
        sentences = decomposed.get("raw_sentences", [])

        sum_lines = [
            f"{topic} 추진에 따른 품의 건임.",
            directives[0][:50] if directives else "목표 대비 120% 성과 달성 및 리드타임 대폭 절감",
        ]

        comparison = {
            "headers": ["구분", "현행(As-Is)", "개선안(To-Be)"],
            "rows": [
                ["방식", "수기 비정형 처리", "MCP 오토컴포즈 자동화"],
                ["생산성", "건당 4시간 소요", "건당 10분 이내 완료"],
            ]
        }

        plan_block = {
            "headers": ["단계", "일정", "주요 과제"],
            "rows": [
                ["1단계", "10월~11월", "시스템 환경 검증 및 파일럿"],
                ["2단계", "12월", "전사 확산 및 안정화"],
            ]
        }

        effects = [
            "업무 리드타임 85% 단축 및 휴먼 에러 원천 방지",
            "정형화된 보고 양식 준수로 임원 의사결정 속도 향상",
        ]

        return NarrativePlanner.plan_docx_proposal(
            title=topic,
            author_info={"author": "기획담당", "dept": "전략기획팀"},
            block1_summary=sum_lines,
            block2_comparison=comparison,
            block3_plan=plan_block,
            block4_effects=effects,
        )
