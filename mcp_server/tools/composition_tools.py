# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_tools.py
4대 오피스(HWP, DOCX, PPT, Excel) 문서 구성 지능(Composition Engine) MCP 도구 등록 모듈.
실무자 고충(공간 예산, 도해 변환, 2D 그리드 토폴로지, 개조식 위계)을 해결하는 고수준 구성 API를 제공합니다.
"""

from __future__ import annotations

import json
import logging

from mcp_server.compat import MCPServer

from .composition_engine import (
    NarrativePlanner,
    SpatialBudgeter,
    CompositionLinter,
    CompositionOrchestrator,
    SemanticDecomposer,
    SmartRouter,
    AutoStoryboarder,
)

logger = logging.getLogger("mcp_server.composition_tools")


def register_composition_tools(mcp: MCPServer) -> None:
    """문서 구성 능력 향상(Composition Engine) 도구군 등록"""

    # -------------------------------------------------------------------------
    # 1. 문서 구성 기획 (Plan Composition)
    # -------------------------------------------------------------------------
    @mcp.tool()
    def office_plan_composition(
        doc_type: str,
        topic: str,
        raw_content: str,
        archetype: str = "auto",
    ) -> str:
        """
        4대 오피스(ppt, excel, hwp, docx) 문서의 목적과 거친 원시 내용을 바탕으로,
        실무 표준 정보 아키텍처(서사, 시각 컴포넌트, 공간 예산, 린터)가 적용된
        고품질 '문서 구성 청사진(Composition Plan)'을 자동 기획합니다.

        Args:
            doc_type: 문서 종류 ('ppt', 'excel', 'hwp', 'docx')
            topic: 문서 주제 (예: '2026년 3분기 경영 실적 보고', '신규 AI 도입 기획안')
            raw_content: 구조화할 원시 텍스트나 JSON 문자열 (슬라이드/데이터/문단 내용)
            archetype: 'auto' 또는 특정 아키타입 ('kpi_grid_3', 'process_horizontal_4', 'executive_dashboard', 'hwp_one_page', 'docx_4block')

        Returns:
            JSON 형식의 사전 검증 완료된 구성 설계 청사진 (Composition Blueprint)
        """
        doc_type = doc_type.lower().strip()

        # 파싱 시도 (JSON 또는 일반 텍스트)
        try:
            parsed = json.loads(raw_content)
        except Exception:
            parsed = raw_content

        if doc_type in ("ppt", "powerpoint"):
            # PPT 슬라이드 구성안 생성
            if isinstance(parsed, list):
                slides_input = parsed
            elif isinstance(parsed, dict) and "slides" in parsed:
                slides_input = parsed["slides"]
            else:
                # 일반 텍스트인 경우 기본 3단 KPI 카드 슬라이드 1장 구성
                arch = "kpi_grid_3" if archetype == "auto" else archetype
                slides_input = [
                    {
                        "archetype": arch,
                        "action_title": f"{topic}의 핵심 지표는 목표 대비 초과 달성함",
                        "items": [
                            {"label": "핵심 성과", "value": "120%", "desc": str(raw_content)[:40]},
                            {"label": "비용 절감", "value": "35%", "desc": "프로세스 자동화 효율화"},
                            {"label": "만족도", "value": "94점", "desc": "사내 실무자 정량 평가"},
                        ],
                        "footnote": "사내 내부 집계 기준 (2026)",
                    }
                ]
            res = NarrativePlanner.plan_ppt(topic=topic, slides_input=slides_input)
            return json.dumps(res, ensure_ascii=False, indent=2)

        elif doc_type in ("excel", "xlsx", "sheet"):
            # Excel 대시보드 토폴로지 구성
            if isinstance(parsed, dict) and "headers" in parsed and "data_rows" in parsed:
                headers = parsed["headers"]
                data_rows = parsed["data_rows"]
                kpi_tiles = parsed.get("kpi_tiles", [
                    {"label": "총 매출", "value": "12,450백만"},
                    {"label": "달성률", "value": "104.2%"},
                    {"label": "전년비", "value": "+18.5%"},
                ])
                unit = parsed.get("unit", "단위: 백만원, %")
                summary_formulas = parsed.get("summary_formulas")
                if not summary_formulas and len(data_rows) > 0:
                    summary_formulas = {"합계": f"=SUM(C8:C{7 + len(data_rows)})"}
            else:
                headers = ["부서명", "목표액", "집행액", "달성률"]
                data_rows = [
                    ["영업1팀", 5000, 5200, 1.04],
                    ["영업2팀", 4000, 4150, 1.03],
                    ["기획팀", 3000, 3100, 1.03],
                ]
                kpi_tiles = [
                    {"label": "총 목표액", "value": "12,000백만"},
                    {"label": "총 집행액", "value": "12,450백만"},
                    {"label": "평균 달성률", "value": "103.7%"},
                ]
                unit = "단위: 백만원, %"
                summary_formulas = {"합계": "=SUM(C8:C10)"}

            res = NarrativePlanner.plan_excel(
                title=topic,
                kpi_tiles=kpi_tiles,
                headers=headers,
                data_rows=data_rows,
                summary_formulas=summary_formulas,
                unit=unit,
            )
            return json.dumps(res, ensure_ascii=False, indent=2)

        elif doc_type in ("hwp", "hwpx", "hancom"):
            # 한글 A4 1장 표준 개조식 보고서 기획
            if isinstance(parsed, dict) and "sections" in parsed:
                exec_sum = parsed.get("executive_summary", [f"{topic} 추진 필요성 및 조기 승인 요청"])
                sections = parsed["sections"]
            else:
                exec_sum = [
                    f"{topic} 관련 핵심 안건 보고 및 예산 집행 승인 건임.",
                    "추진 타당성 분석 결과 ROI 180% 달성 예상되며 위험 요인 해소 완료됨.",
                ]
                sections = [
                    {
                        "title": "추진 배경 및 필요성",
                        "items": [
                            {"level": 2, "symbol": "○", "text": "시장 환경 변화에 따른 업무 자동화 도입 시급성 증대"},
                            {"level": 3, "symbol": "-", "text": "현행 수작업 방식의 비효율 및 입력 오류 발생 지속"},
                        ],
                    },
                    {
                        "title": "주요 추진 계획",
                        "items": [
                            {"level": 2, "symbol": "○", "text": "전사 MCP 기반 오피스 자동화 파이프라인 단계적 적용"},
                            {"level": 3, "symbol": "-", "text": "4대 도구별 시각적 아키타입 표준화 및 실시간 COM 제어 구축"},
                        ],
                    },
                ]

            res = NarrativePlanner.plan_hwp_one_page(
                title=topic,
                executive_summary=exec_sum,
                sections=sections,
            )
            return json.dumps(res, ensure_ascii=False, indent=2)

        elif doc_type in ("docx", "word"):
            # 워드 4대 표준 블록 품의서 기획
            if isinstance(parsed, dict) and "blocks" in parsed:
                blocks = parsed["blocks"]
                sum_lines = blocks.get("block1_summary", [f"{topic} 핵심 요약"])
                comp = blocks.get("block2_comparison", {})
                plan_block = blocks.get("block3_plan", {})
                effects = blocks.get("block4_effects", ["업무 생산성 40% 향상"])
            else:
                sum_lines = [
                    f"{topic}에 대한 전사 적용 품의 건임.",
                    "초기 구축 비용 대비 연간 3.5억 원의 리드타임 절감 효과 확인.",
                ]
                comp = {
                    "headers": ["구분", "현행(As-Is)", "개선안(To-Be)"],
                    "rows": [
                        ["작성 방식", "개별 수기 작성 및 서식 노가다", "MCP 아키타입 자동 조립"],
                        ["소요 시간", "평균 4.5시간/건", "평균 15분/건 (85% 단축)"],
                    ],
                }
                plan_block = {"headers": ["단계", "일정", "담당"], "rows": [["1단계", "11월", "IT기획"]]}
                effects = [
                    "문서 작성 표준화율 100% 달성 및 서식 왜곡 제거",
                    "임원 보고용 시각적 가독성 대폭 향상",
                ]

            res = NarrativePlanner.plan_docx_proposal(
                title=topic,
                author_info={"author": "기획담당", "dept": "전략기획팀"},
                block1_summary=sum_lines,
                block2_comparison=comp,
                block3_plan=plan_block,
                block4_effects=effects,
            )
            return json.dumps(res, ensure_ascii=False, indent=2)

        else:
            return json.dumps(
                {"error": f"지원하지 않는 문서 종류입니다: {doc_type} (지원: ppt, excel, hwp, docx)"},
                ensure_ascii=False,
            )

    # -------------------------------------------------------------------------
    # 2. 문서 구성 실시간 렌더링 (Render Composition)
    # -------------------------------------------------------------------------
    @mcp.tool()
    def office_render_composition(
        doc_type: str,
        composition_plan: str,
        new_window: bool = True,
    ) -> str:
        """
        office_plan_composition으로 기획/검증된 구성 청사진을
        현재 활성화된 또는 신규 데스크톱 오피스 창(PPT, Excel, HWP, Word)에
        시각적 아키타입(도해, 대시보드 타일, 개조식 위계, 요약 박스)으로 실시간 렌더링합니다.

        Args:
            doc_type: 문서 종류 ('ppt', 'excel', 'hwp', 'docx')
            composition_plan: office_plan_composition이 출력한 JSON 문자열
            new_window: 신규 창/문서로 생성할지 여부 (기본 True)

        Returns:
            실시간 렌더링 결과 JSON
        """
        doc_type = doc_type.lower().strip()
        try:
            plan = json.loads(composition_plan)
        except Exception as e:
            return json.dumps({"error": f"구성 청사진 JSON 파싱 실패: {e}"}, ensure_ascii=False)

        try:
            if doc_type in ("ppt", "powerpoint"):
                res = CompositionOrchestrator.render_ppt_composition(plan, new_presentation=new_window)
            elif doc_type in ("excel", "xlsx", "sheet"):
                res = CompositionOrchestrator.render_excel_composition(plan, new_workbook=new_window)
            elif doc_type in ("hwp", "hwpx", "hancom"):
                res = CompositionOrchestrator.render_hwp_composition(plan)
            elif doc_type in ("docx", "word"):
                res = CompositionOrchestrator.render_docx_composition(plan, new_document=new_window)
            else:
                return json.dumps({"error": f"지원하지 않는 doc_type: {doc_type}"}, ensure_ascii=False)

            return json.dumps(res, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.exception("실시간 문서 구성 렌더링 중 오류 발생")
            return json.dumps({"status": "error", "message": f"렌더링 실패: {e}"}, ensure_ascii=False)

    # -------------------------------------------------------------------------
    # 3. 공간 예산 및 품질 감사 (Audit Composition Budget)
    # -------------------------------------------------------------------------
    @mcp.tool()
    def office_audit_composition_budget(
        doc_type: str,
        spec: str,
    ) -> str:
        """
        작성 중인 문서의 공간 예산(Spatial Budget) 및 구성 품질을 감사(Audit)합니다:
        - HWP: A4 1페이지 28줄 상한 초과 여부 및 장평/자간/행간 압축 제안
        - PPT: 카드당 본문 45자 초과 여부 및 결론형 Action Title 적합성
        - Excel: 가로 1페이지 인쇄 규격(12열 이내) 및 KPI 카드/수식 누락 점검
        - DOCX: 4대 핵심 블록 누락 검사

        Args:
            doc_type: 문서 종류 ('ppt', 'excel', 'hwp', 'docx')
            spec: 검사할 구성 스펙 JSON 문자열

        Returns:
            품질 감사 결과 및 권고사항 JSON
        """
        doc_type = doc_type.lower().strip()
        try:
            data = json.loads(spec)
        except Exception as e:
            return json.dumps({"error": f"JSON 파싱 실패: {e}"}, ensure_ascii=False)

        if doc_type in ("ppt", "powerpoint"):
            budget = SpatialBudgeter.audit_ppt_slide(data)
            lint = CompositionLinter.lint_ppt_slide(data)
            return json.dumps({"budget": budget, "lint": lint}, ensure_ascii=False, indent=2)
        elif doc_type in ("excel", "xlsx"):
            budget = SpatialBudgeter.audit_excel_topology(data)
            lint = CompositionLinter.lint_excel_topology(data)
            return json.dumps({"budget": budget, "lint": lint}, ensure_ascii=False, indent=2)
        elif doc_type in ("hwp", "hwpx"):
            budget = SpatialBudgeter.audit_hwp_one_page(data)
            lint = CompositionLinter.lint_hwp_one_page(data)
            return json.dumps({"budget": budget, "lint": lint}, ensure_ascii=False, indent=2)
        elif doc_type in ("docx", "word"):
            lint = CompositionLinter.lint_docx_proposal(data)
            return json.dumps({"lint": lint}, ensure_ascii=False, indent=2)
        else:
            return json.dumps({"error": f"지원하지 않는 doc_type: {doc_type}"}, ensure_ascii=False)

    # -------------------------------------------------------------------------
    # 4. 원시 덤프 자율 구성 마스터 툴 (Auto-Compose Raw Dump)
    # -------------------------------------------------------------------------
    @mcp.tool()
    def office_autocompose_raw_dump(
        raw_text: str,
        topic: str = "auto",
        doc_type: str = "auto",
        open_in_desktop: bool = False,
    ) -> str:
        """
        [원샷 자율 구성 마스터 도구]
        정리되지 않은 날것의 텍스트(회의록 녹취, 메모 뭉치, 뉴스 기사, 비정형 데이터)를
        한 번에 털어 넣으면, AI가 시맨틱 성분(수치, 일정, 비교, 결론)을 자동 분해하고
        최적의 오피스 도구(PPT, Excel, HWP, DOCX)를 자율 결정하여
        실무 표준 아키타입으로 완성된 문서를 기획 및 렌더링합니다.

        Args:
            raw_text: 사용자가 입력한 날것의 비정형 텍스트 덤프
            topic: 문서 주제 ('auto' 입력 시 텍스트 첫 줄이나 핵심 키워드에서 자동 추출)
            doc_type: 대상 오피스 도구 ('auto', 'ppt', 'excel', 'hwp', 'docx')
            open_in_desktop: 데스크톱 오피스 활성 창에 실시간 생성할지 여부 (기본 False)

        Returns:
            자동 추출된 성분 요약, 선정된 문서 포맷, 생성된 구성 청사진 및 품질 상태 JSON
        """
        raw_text = raw_text.strip()
        if not raw_text:
            return json.dumps({"error": "입력된 텍스트가 비어 있습니다."}, ensure_ascii=False)

        # 1. 시맨틱 성분 분해 (Metrics, Chronology, Duality, Directives)
        decomposed = SemanticDecomposer.decompose(raw_text)

        # 2. 문서 주제(Topic) 결정
        resolved_topic = topic
        if not resolved_topic or resolved_topic == "auto":
            first_sentence = decomposed.get("raw_sentences", ["비즈니스 업무 보고"])[0]
            # 첫 문장이 너무 길면 30자로 자름
            resolved_topic = first_sentence[:30].strip()
            if resolved_topic.endswith("."):
                resolved_topic = resolved_topic[:-1]

        # 3. 문서 포맷(Doc Type) 결정
        resolved_doc_type = doc_type.lower().strip()
        if not resolved_doc_type or resolved_doc_type == "auto":
            resolved_doc_type = SmartRouter.route_format(raw_text, decomposed)

        # 4. 자율 스토리보딩 및 청사진 빌드
        if resolved_doc_type in ("ppt", "powerpoint"):
            plan = AutoStoryboarder.storyboard_ppt(resolved_topic, decomposed)
        elif resolved_doc_type in ("excel", "xlsx", "sheet"):
            plan = AutoStoryboarder.storyboard_excel(resolved_topic, decomposed)
        elif resolved_doc_type in ("hwp", "hwpx", "hancom"):
            plan = AutoStoryboarder.storyboard_hwp(resolved_topic, decomposed)
        elif resolved_doc_type in ("docx", "word"):
            plan = AutoStoryboarder.storyboard_docx(resolved_topic, decomposed)
        else:
            return json.dumps({"error": f"지원하지 않는 doc_type: {resolved_doc_type}"}, ensure_ascii=False)

        # 5. 데스크톱 오피스 실시간 렌더링 (옵션)
        render_result = None
        if open_in_desktop:
            try:
                if resolved_doc_type in ("ppt", "powerpoint"):
                    render_result = CompositionOrchestrator.render_ppt_composition(plan, new_presentation=True)
                elif resolved_doc_type in ("excel", "xlsx", "sheet"):
                    render_result = CompositionOrchestrator.render_excel_composition(plan, new_workbook=True)
                elif resolved_doc_type in ("hwp", "hwpx", "hancom"):
                    render_result = CompositionOrchestrator.render_hwp_composition(plan)
                elif resolved_doc_type in ("docx", "word"):
                    render_result = CompositionOrchestrator.render_docx_composition(plan, new_document=True)
            except Exception as e:
                logger.warning(f"데스크톱 렌더링 중 오류 발생 (무시하고 계획 반환): {e}")
                render_result = {"status": "warning", "message": f"데스크톱 렌더링 실패: {e}"}

        response = {
            "status": "success",
            "autocomposed_summary": {
                "resolved_topic": resolved_topic,
                "selected_doc_type": resolved_doc_type,
                "extracted_metrics_count": len(decomposed.get("metrics", [])),
                "extracted_chronology_count": len(decomposed.get("chronology", [])),
                "extracted_duality_count": len(decomposed.get("duality", [])),
                "quality_status": plan.get("quality_status", "EXCELLENT"),
            },
            "plan": plan,
            "desktop_render": render_result,
        }

        return json.dumps(response, ensure_ascii=False, indent=2)
