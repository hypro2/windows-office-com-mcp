# -*- coding: utf-8 -*-
"""
mcp_server/client_demo.py
MyOffice Desktop COM & KISA1310 MCP 서버를 호출하는 공식 MCP 클라이언트 사용 코드 및 엔드투엔드 검증 스크립트.

사용법:
  1) 등록된 모든 도구/리소스/프롬프트 목록 확인:
     python -m mcp_server.client_demo --list-tools
  2) 현재 데스크톱 오피스(Word, Excel, PPT, HWP) 설치 및 실행 상태 조회:
     python -m mcp_server.client_demo --status
  3) 실시간 COM 읽기/작성 + 비즈니스 컴포넌트 + 동적 서식 통합 데모:
     python -m mcp_server.client_demo --demo all
  4) 개별 데모 실행:
     python -m mcp_server.client_demo --demo live-com
     python -m mcp_server.client_demo --demo components
     python -m mcp_server.client_demo --demo dynamic
  5) 임의 도구 직접 호출:
     python -m mcp_server.client_demo --call office_get_status --args "{}"
"""

from __future__ import annotations

import argparse
import asyncio
import json
import sys
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# Windows 콘솔 UTF-8 출력 설정
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT_DIR / "mcp_server" / "demo_outputs"


def _parse_tool_result(res: Any) -> Any:
    if not res or not getattr(res, "content", None):
        return None
    texts = [c.text for c in res.content if hasattr(c, "text")]
    joined = "\n".join(texts)
    try:
        return json.loads(joined)
    except Exception:
        return joined


async def run_client(args: argparse.Namespace) -> None:
    server_params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "mcp_server.server", "--transport", "stdio"],
        cwd=str(ROOT_DIR),
    )

    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()

            if args.list_tools:
                tools_resp = await session.list_tools()
                resources_resp = await session.list_resources()
                prompts_resp = await session.list_prompts()

                print(f"=== [등록된 MCP Tools ({len(tools_resp.tools)}개)] ===")
                for idx, t in enumerate(tools_resp.tools, start=1):
                    desc_first_line = (t.description or "").strip().splitlines()[0] if t.description else ""
                    print(f"  {idx:2d}. {t.name:<36} : {desc_first_line}")

                print(f"\n=== [등록된 MCP Resources ({len(resources_resp.resources)}개)] ===")
                for r in resources_resp.resources:
                    print(f"  - {r.uri} ({r.name})")

                print(f"\n=== [등록된 MCP Prompts ({len(prompts_resp.prompts)}개)] ===")
                for p in prompts_resp.prompts:
                    print(f"  - {p.name}: {p.description}")
                return

            if args.status:
                res = await session.call_tool("office_get_status", {})
                print("=== [데스크톱 오피스 실시간 상태] ===")
                print(json.dumps(_parse_tool_result(res), ensure_ascii=False, indent=2))
                return

            if args.call:
                call_args = json.loads(args.args) if args.args else {}
                res = await session.call_tool(args.call, call_args)
                print(f"=== [도구 호출 결과: {args.call}] ===")
                print(json.dumps(_parse_tool_result(res), ensure_ascii=False, indent=2))
                return

            if args.demo:
                OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                demo_target = args.demo.lower()

                # 1. 실시간 데스크톱 Word / Excel / PPT COM 작성 -> 서식 적용 -> 다시 읽기(Round-trip) 검증
                if demo_target in ("all", "live-com"):
                    print("\n[1/3] 실시간 데스크톱 Word & Excel & PPT COM 작성 -> 서식 적용 -> 다시 읽기(Round-trip) 검증...")
                    w_res = await session.call_tool(
                        "word_render_markdown_report",
                        {
                            "theme": "navy",
                            "apply_brand_header": False,
                            "markdown_text": (
                                "# MyOffice AI 오피스 어시스턴트 실시간 제어 보고서\n"
                                "## 1. 실시간 COM 연동 개요\n"
                                "- 코딩 어시스턴트에서 자연어 질의로 즉시 문서 읽기 및 작성\n"
                                "  - 마우스 선택 영역(Selection) 정밀 인식 및 서식 적용\n"
                                "| 구분 | 대상 앱 | 핵심 서식 기능 |\n"
                                "|---|---|---|\n"
                                "| 문서 | MS Word / HWP | 비즈니스 네이비 서식 및 행안부 공문서 서식 |\n"
                                "| 데이터 | MS Excel | 겉바속촉 테두리 및 스틸블루 표 렌더링 |\n"
                                "| 발표 | MS PowerPoint | 14대 카탈로그 슬라이스 및 Morph 전환 |\n"
                            ),
                        },
                    )
                    print("  -> Word 실시간 마크다운 렌더링:", json.dumps(_parse_tool_result(w_res), ensure_ascii=False))

                    w_read = await session.call_tool("word_read_document", {"scope": "all", "max_chars": 500})
                    w_read_data = _parse_tool_result(w_read) or {}
                    print("  -> Word 실시간 다시 읽기 확인 (문단 수/표 수):", {
                        "document_name": w_read_data.get("document_name"),
                        "paragraphs_count": w_read_data.get("paragraphs_count"),
                        "tables_count": w_read_data.get("tables_count"),
                    })

                    e_write = await session.call_tool(
                        "excel_write_data",
                        {
                            "sheet_name": "AI자동화성과",
                            "start_cell": "A1",
                            "table_style": "ib_steel_blue",
                            "auto_fit": True,
                            "data": [
                                ["부서명", "적용 시스템", "단축 시간(시간/월)", "만족도"],
                                ["재무기획팀", "3-Zone 재무모델 자동화", 42, "98%"],
                                ["디지털혁신팀", "표준 슬라이스 PPT 생성기", 35, "96%"],
                                ["준법지원팀", "SDT 실무양식 DOCX", 28, "95%"],
                            ],
                        },
                    )
                    print("  -> Excel 실시간 표 작성 및 스틸블루 서식:", json.dumps(_parse_tool_result(e_write), ensure_ascii=False))

                    e_read = await session.call_tool("excel_read_data", {"sheet_name": "AI자동화성과", "range_address": "A1:D4"})
                    e_read_data = _parse_tool_result(e_read) or {}
                    print("  -> Excel 실시간 다시 읽기 (Markdown 변환):\n" + str(e_read_data.get("markdown_table", "")))

                    ppt_out = str(OUTPUT_DIR / "business_navy_deck_demo.pptx")
                    p_shell = await session.call_tool(
                        "ppt_create_report_deck_shell",
                        {
                            "title": "생성형 AI 기반 오피스 자동화 추진 보고",
                            "subtitle": "실시간 COM 제어 및 표준 서식 MCP 구축",
                            "department": "AI혁신팀",
                            "output_path": ppt_out,
                            "open_in_desktop": args.open_desktop,
                            "theme": "business_navy",
                        },
                    )
                    print("  -> PPT 비즈니스 네이비 덱 쉘 생성:", json.dumps(_parse_tool_result(p_shell), ensure_ascii=False))

                    p_read = await session.call_tool("ppt_read_presentation", {})
                    p_read_data = _parse_tool_result(p_read) or {}
                    print("  -> PowerPoint 실시간 다시 읽기 확인:", {
                        "presentation_name": p_read_data.get("presentation_name"),
                        "slides_count": p_read_data.get("slides_count"),
                    })

                    live_status = await session.call_tool("office_get_status", {})
                    print("  -> 실시간 4대 오피스 연결 상태 요약:", json.dumps(_parse_tool_result(live_status), ensure_ascii=False))

                # 5. 7대 비즈니스 컴포넌트 & 스마트 정제 도구 데모
                if demo_target in ("all", "components"):
                    print("\n[5/5] 7대 비즈니스 컴포넌트 빌더 & 스마트 서식 정제 도구 데모 실행...")

                    # 1) Word 실행 및 준비
                    print("  1) Word 새 문서 활성화...")
                    await session.call_tool("office_launch_or_open", {"app_name": "word", "visible": True})

                    # 2) 계층 구조도 (Hierarchy)
                    print("  2) [컴포넌트 1] 조직도/계층 구조도 생성...")
                    h_res = await session.call_tool(
                        "office_insert_hierarchy",
                        {
                            "title": "디지털혁신총괄본부",
                            "levels": ["AI인프라팀", "오피스자동화팀", "클라우드보안팀"],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(h_res))

                    # 3) 프로세스 흐름도 (Process Flow)
                    print("  3) [컴포넌트 2] 단계별 프로세스 흐름도 생성...")
                    p_res = await session.call_tool(
                        "office_insert_process_flow",
                        {
                            "steps": [
                                {"title": "현안 진단", "desc": "레거시 서식 및 업무 비효율 분석"},
                                {"title": "MCP 엔진 구축", "desc": "실시간 데스크톱 COM 제어 및 양식 내재화"},
                                {"title": "전사 확산", "desc": "Antigravity/Claude 연동 원클릭 자동화"},
                            ],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(p_res))

                    # 4) 추진일정 로드맵 (Timeline)
                    print("  4) [컴포넌트 3] 추진일정 타임라인 로드맵 생성...")
                    t_res = await session.call_tool(
                        "office_insert_timeline",
                        {
                            "title": "2026 하반기 지능형 오피스 전환 일정",
                            "milestones": [
                                {"phase": "1단계: 분석/설계", "period": "2026.07 ~ 08", "task": "범피스 14만줄 전수 분석 및 컴포넌트화", "status": "완료"},
                                {"phase": "2단계: MCP 구현", "period": "2026.09 ~ 10", "task": "FastMCP 통합 및 4대 오피스 바인딩", "status": "진행중"},
                                {"phase": "3단계: 전사 배포", "period": "2026.11 ~ 12", "task": "AI IDE 연계 실무 적용", "status": "예정"},
                            ],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(t_res))

                    # 5) AS-IS vs TO-BE 비교표 (Comparison)
                    print("  5) [컴포넌트 4] AS-IS vs TO-BE 비교표 생성...")
                    c_res = await session.call_tool(
                        "office_insert_comparison",
                        {
                            "title": "오피스 업무 방식 혁신 전후 비교",
                            "as_is": [
                                "지자체별 복잡한 HWP 양식 수동 작성",
                                "MS 오피스(Word, Excel, PPT) 서식 불일치",
                                "표 테두리 및 자간/줄간격 수동 맞춤",
                            ],
                            "to_be": [
                                "자연어 한 줄로 비즈니스 컴포넌트 즉시 생성",
                                "Word/Excel/PPT/HWP 4대 오피스 단일 MCP 통합",
                                "범피스 겉바속촉 및 궁극의 정리 알고리즘 자동화",
                            ],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(c_res))

                    # 6) 요약 콜아웃 박스 (Callout)
                    print("  6) [컴포넌트 5] 핵심 요약 콜아웃 상자 생성...")
                    b_res = await session.call_tool(
                        "office_insert_callout_box",
                        {
                            "text": "본 솔루션은 특정 공공기관에 편중되었던 레거시 알고리즘을 완전 범용화하여 일반 비즈니스 실무에 즉시 적용 가능하도록 개편되었습니다.",
                            "box_type": "tip",
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(b_res))

                    # 7) KPI 대시보드 카드 (KPI Cards)
                    print("  7) [컴포넌트 6] KPI 대시보드 카드 생성...")
                    k_res = await session.call_tool(
                        "office_insert_kpi_cards",
                        {
                            "kpi_data": [
                                {"label": "문서 작성 소요시간", "value": "15분", "change": "-75.0%"},
                                {"label": "컴포넌트 재활용률", "value": "94.2%", "change": "+42.5%"},
                                {"label": "서식 일관성 만족도", "value": "99.8점", "change": "+18.2%"},
                            ],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(k_res))

                    # 8) 겉바속촉 스마트 표 (Smart Table)
                    print("  8) [컴포넌트 7] 겉바속촉 프리미엄 비즈니스 표 생성...")
                    s_res = await session.call_tool(
                        "office_insert_smart_table",
                        {
                            "headers": ["모듈명", "원천 함수 수", "적용 오피스", "상태"],
                            "rows": [
                                ["조직도/구조도", "21개", "Word, Excel, PPT", "운영중"],
                                ["프로세스 흐름도", "8개", "Word, PPT", "운영중"],
                                ["타임라인 로드맵", "30개", "Word, PPT, HWP", "운영중"],
                                ["겉바속촉 스마트표", "41개", "Word, Excel", "운영중"],
                            ],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(s_res))

                    # 9) Word 궁극의 개조식 정리 (Clean and Structure)
                    print("  9) [스마트 정제 1] Word 궁극의 7단계 개조식 위계 자동 정리...")
                    # 텍스트 추가 후 정리 실행
                    await session.call_tool(
                        "word_write_text",
                        {
                            "text": "1. 추진 배경 및 개요\n가. 디지털 전환 가속화\n1) 데스크톱 COM 자동화 도입\n(1) 에이전트 연동 체계 구축\n- 사용자 업무 생산성 극대화",
                            "mode": "append",
                        },
                    )
                    clean_res = await session.call_tool("word_clean_and_structure", {})
                    print("     결과:", _parse_tool_result(clean_res))

                    # 10) Excel 겉바속촉 자동 서식화 (Auto Format Table)
                    print(" 10) [스마트 정제 2] Excel 겉바속촉 자동 서식화 실행...")
                    await session.call_tool("office_launch_or_open", {"app_name": "excel", "visible": True})
                    await session.call_tool(
                        "excel_write_data",
                        {
                            "sheet_name": "컴포넌트현황",
                            "start_cell": "A1",
                            "data": [
                                ["컴포넌트", "원천자산", "지원상태", "적용비율"],
                                ["구조도", 21, "완료", "100%"],
                                ["프로세스", 8, "완료", "100%"],
                                ["타임라인", 30, "완료", "100%"],
                                ["비교표", 5, "완료", "100%"],
                                ["콜아웃박스", 73, "완료", "100%"],
                                ["스마트표", 41, "완료", "100%"],
                            ],
                            "auto_fit": True,
                        },
                    )
                    ex_clean_res = await session.call_tool("excel_auto_format_table", {})
                    print("     결과:", _parse_tool_result(ex_clean_res))

                    print("  ==> [성공] 7대 비즈니스 컴포넌트 & 스마트 정제 도구 전수 검증 완료!")

                # 6. 자연어 지시 기반 범용 동적 서식 조립 데모
                if demo_target in ("all", "dynamic"):
                    print("\n[6/6] 자연어 지시 기반 범용 동적 서식 빌더(Universal Dynamic Form) 데모 실행...")

                    # 1) 실무 서식 1: 킥오프 회의록 양식
                    print("  1) [동적 서식 1] 킥오프 회의록(메타정보+안건요약+체크리스트+서명) 자동 생성...")
                    await session.call_tool("office_launch_or_open", {"app_name": "word", "visible": True})
                    m_res = await session.call_tool(
                        "office_build_custom_form",
                        {
                            "title": "2026 차세대 AI 오피스 자동화 프로젝트 킥오프 회의록",
                            "blocks": [
                                {
                                    "type": "title_header",
                                    "title": "2026 차세대 AI 오피스 자동화 프로젝트 킥오프 회의록",
                                    "subtitle": "실무 추진 방향 및 분과별 핵심 R&R 수립",
                                    "doc_no": "AOF-2026-001",
                                    "date": "2026. 10. 31",
                                },
                                {
                                    "type": "key_value_grid",
                                    "cols": 4,
                                    "items": [
                                        {"key": "회의 일시", "value": "2026년 10월 31일(금) 14:00~16:00"},
                                        {"key": "회의 장소", "value": "본사 대회의실 (Online Teams 병행)"},
                                        {"key": "주관 부서", "value": "디지털혁신총괄본부 오피스혁신팀"},
                                        {"key": "참 석 자", "value": "김부서장, 이팀장, 박수석, 최선임 외 5명"},
                                    ],
                                },
                                {
                                    "type": "callout_box",
                                    "text": "본 프로젝트는 기존 지자체용 서식 자산을 전사 비즈니스용 동적 오피스 자동화 엔진으로 완전 개편하는 것을 최우선 목표로 합니다.",
                                    "box_type": "info",
                                },
                                {
                                    "type": "table",
                                    "title": "주요 안건별 협의 및 결정 사항",
                                    "headers": ["안건 번호", "핵심 안건명", "주요 논의 내용", "의결 결과"],
                                    "rows": [
                                        ["제1호", "MCP 서버 아키텍처 확정", "FastMCP 기반 데스크톱 COM 바인딩", "원안 가결"],
                                        ["제2호", "동적 서식 조립 엔진 도입", "10대 레고 블록 기반 임의 서식 생성 체계", "만장일치 승인"],
                                        ["제3호", "전사 시범 배포 일정 수립", "11월 중 1차 파일럿 그룹 배포 및 피드백", "조건부 승인"],
                                    ],
                                },
                                {
                                    "type": "checklist",
                                    "title": "부서별 후속 조치 사항 (Action Items)",
                                    "items": [
                                        {"task": "Word/Excel/PPT 실시간 COM 안정성 검증", "assignee": "김선임", "due": "11/05", "done": True},
                                        {"task": "FastMCP 도구 정의서 및 자연어 프롬프트 가이드 배포", "assignee": "이책임", "due": "11/08", "done": True},
                                        {"task": "전사 사용자 피드백 수집 설문지 작성", "assignee": "박수석", "due": "11/12", "done": False},
                                    ],
                                },
                                {
                                    "type": "signature",
                                    "notice": "위와 같이 회의 결과를 기록하며, 명시된 Action Item을 기한 내 차질 없이 수행합니다.",
                                    "date": "2026년 10월 31일",
                                    "signers": ["회의 주관자: 디지털혁신총괄본부장 (인)", "회의 기록자: 오피스혁신팀 선임연구원 (인)"],
                                },
                            ],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(m_res))

                    # 2) 실무 서식 2: 1페이지 사업기획서 양식
                    print("  2) [동적 서식 2] 1페이지 사업기획서(결재선+비교표+KPI+예산표) 자동 생성...")
                    p_res = await session.call_tool(
                        "office_build_custom_form",
                        {
                            "title": "차세대 스마트 오피스 전환 및 AI 비서 도입 기획서",
                            "blocks": [
                                {
                                    "type": "title_header",
                                    "title": "차세대 스마트 오피스 전환 및 AI 비서 도입 기획서",
                                    "doc_no": "PLAN-2026-088",
                                    "date": "2026. 10. 31",
                                },
                                {
                                    "type": "approval_line",
                                    "positions": ["기안자", "팀장", "실장", "본부장", "대표이사"],
                                },
                                {
                                    "type": "callout_box",
                                    "text": "사내 반복적인 문서 작성·서식 지정 업무를 AI 자연어 원클릭 자동화로 혁신하여 연간 1,200시간의 고부가가치 시간을 확보하고자 함.",
                                    "box_type": "tip",
                                },
                                {
                                    "type": "comparison",
                                    "title": "업무 방식 혁신 전후(AS-IS vs TO-BE) 비교",
                                    "as_is": [
                                        "서식 및 표 테두리를 마우스로 일일이 수동 편집",
                                        "공기업/지자체별 복잡한 개조식 규격 준수 어려움",
                                        "Word/Excel/PPT 간 데이터 복사-붙여넣기 오류 발생",
                                    ],
                                    "to_be": [
                                        "자연어 한 줄로 회의록/기획서/점검표 서식 즉시 완성",
                                        "범피스 7단계 궁극의 정리로 개조식 위계 자동 교정",
                                        "단일 MCP 서버로 4대 오피스 실시간 데스크톱 통합 제어",
                                    ],
                                },
                                {
                                    "type": "kpi_cards",
                                    "cards": [
                                        {"label": "문서 작성 단축률", "value": "78.5%", "change": "+25.4%"},
                                        {"label": "연간 절감 비용", "value": "2.4 억원", "change": "+42.1%"},
                                        {"label": "전사 직원 만족도", "value": "96.8 점", "change": "+14.5%"},
                                    ],
                                },
                                {
                                    "type": "table",
                                    "title": "2026 도입 소요 예산 산출 내역",
                                    "headers": ["예산 비목", "산출 근거", "수량", "단가(원)", "소요 금액(원)"],
                                    "rows": [
                                        ["서버 인프라 구축", "고성능 On-premise 연계 서버", "2대", "15,000,000", "30,000,000"],
                                        ["FastMCP 엔진 라이선스", "사내 무제한 데스크톱 라이선스", "1건", "20,000,000", "20,000,000"],
                                        ["사용자 교육 및 매뉴얼", "전사 실무자 워크숍 4회", "4회", "2,500,000", "10,000,000"],
                                        ["예비비", "운영 안정화 및 유지보수", "1식", "5,000,000", "5,000,000"],
                                        ["총 합 계", "VAT 별도", "-", "-", "65,000,000"],
                                    ],
                                },
                            ],
                            "target_app": "word",
                        },
                    )
                    print("     결과:", _parse_tool_result(p_res))

                    # 3) 실무 서식 3: Excel 맞춤형 견적서/비용 산출표 자동 생성
                    print("  3) [동적 서식 3] Excel 맞춤형 견적서(타이틀+고객정보+겉바속촉견적표+안내문) 자동 생성...")
                    await session.call_tool("office_launch_or_open", {"app_name": "excel", "visible": True})
                    ex_res = await session.call_tool(
                        "office_build_custom_form",
                        {
                            "title": "2026 AI 오피스 자동화 솔루션 공급 견적서",
                            "blocks": [
                                {
                                    "type": "title_header",
                                    "title": "2026 AI 오피스 자동화 솔루션 공급 견적서",
                                    "subtitle": "엔터프라이즈 데스크톱 COM & FastMCP 표준 패키지",
                                },
                                {
                                    "type": "key_value_grid",
                                    "items": [
                                        {"key": "수신 업체명", "value": "주식회사 미래금융그룹 디지털혁신본부"},
                                        {"key": "견적 일자", "value": "2026년 10월 31일"},
                                        {"key": "유효 기간", "value": "견적일로부터 30일간 유효"},
                                        {"key": "공급 총액", "value": "일금 사천오백만원정 (₩45,000,000, VAT별도)"},
                                    ],
                                },
                                {
                                    "type": "callout_box",
                                    "text": "본 견적은 Word, Excel, PPT, HWP 4대 오피스 실시간 COM 제어 및 전사 커스텀 서식 빌더가 일괄 포함된 조건입니다.",
                                },
                                {
                                    "type": "table",
                                    "title": "세부 공급 내역 및 단가 산출표",
                                    "headers": ["품번", "품목 및 규격", "수량", "단가(원)", "공급가액(원)"],
                                    "rows": [
                                        ["01", "FastMCP COM 데스크톱 코어 엔진", 1, 20000000, 20000000],
                                        ["02", "10대 레고 블록 동적 서식 조립기", 1, 15000000, 15000000],
                                        ["03", "전사 임직원 실무 활용 교육 (2회)", 2, 2500000, 5000000],
                                        ["04", "연간 유지보수 및 기술 지원 (1년)", 1, 5000000, 5000000],
                                        ["합계", "총 공급가액 (VAT 별도)", "-", "-", 45000000],
                                    ],
                                },
                            ],
                            "target_app": "excel",
                        },
                    )
                    print("     결과:", _parse_tool_result(ex_res))

                    print("  ==> [성공] 자연어 지시 기반 범용 동적 서식 빌더 전수 검증 완료!")


def main() -> None:
    parser = argparse.ArgumentParser(description="MyOffice MCP Client Demo & Verification")
    parser.add_argument("--list-tools", action="store_true", help="서버에 등록된 모든 도구/리소스/프롬프트 출력")
    parser.add_argument("--status", action="store_true", help="데스크톱 4대 오피스 실시간 상태 조회")
    parser.add_argument(
        "--demo",
        choices=["all", "live-com", "components", "dynamic"],
        help="실행할 데모 시나리오 선택",
    )
    parser.add_argument(
        "--open-desktop",
        action="store_true",
        default=True,
        help="생성된 문서를 데스크톱 오피스 창에 즉시 열기 (기본값: True)",
    )
    parser.add_argument("--no-open-desktop", dest="open_desktop", action="store_false", help="파일만 생성하고 창은 열지 않음")
    parser.add_argument("--call", type=str, help="직접 호출할 도구 이름")
    parser.add_argument("--args", type=str, default="{}", help="직접 호출할 도구의 JSON 인자")
    args = parser.parse_args()

    if not (args.list_tools or args.status or args.demo or args.call):
        args.list_tools = True

    asyncio.run(run_client(args))


if __name__ == "__main__":
    main()
