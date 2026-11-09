# -*- coding: utf-8 -*-
"""
mcp_server/server.py
실시간 데스크톱 COM 제어(MS Word, Excel, PowerPoint, 한컴 HWP) 및
범용 비즈니스 · KISA 범피스 1310 공문서 서식 통합 MCPServer (mcp 2.3.0).

실행 방법:
  1) stdio 모드 (Claude Desktop / Antigravity IDE / Cursor 기본 연동):
     python -m mcp_server.server
  2) SSE 모드 (웹/원격 에이전트 연동):
     python -m mcp_server.server --transport sse --host 127.0.0.1 --port 8765
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가하여 mcp_server 임포트 보장
ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from mcp_server.compat import MCPServer

from mcp_server.com_manager import com_manager
from mcp_server.tools.chart_tools import register_chart_tools
from mcp_server.tools.component_tools import register_component_tools
from mcp_server.tools.document_forms import register_document_form_tools
from mcp_server.tools.dynamic_form_tools import register_dynamic_form_tools
from mcp_server.tools.excel_tools import register_excel_tools
from mcp_server.tools.hwp_tools import KISA1310_OFFICIAL_REPORT_GUIDE, register_hwp_tools
from mcp_server.tools.ppt_tools import register_ppt_tools
from mcp_server.tools.system_tools import register_system_tools
from mcp_server.tools.folder_tools import register_folder_tools
from mcp_server.tools.user_template_tools import register_user_template_tools
from mcp_server.tools.word_tools import register_word_tools
from mcp_server.tools.composition_tools import register_composition_tools
from mcp_server.tools.ppt_morph_engine import register_ppt_morph_tools
from mcp_server.tools.word_form_and_academic_tools import register_word_form_academic_tools
from mcp_server.tools.excel_modeling_tools import register_excel_modeling_tools
from mcp_server.tools.delivery_gate_and_merge_tools import register_delivery_gate_tools

mcp = MCPServer(
    "MyOffice-Desktop-COM-MCP",
    instructions=(
        "[MyOffice-Desktop-COM-MCP 핵심 작업 지침]\n"
        "1. [무조건 최초 1단계: 편집기 탐색 및 내용 먼저 확인]\n"
        "   - 절대 한 번에 전체 문서를 쫙 만들거나 임의로 새 문서를 생성하지 마십시오.\n"
        "   - 사용자가 요청하면 무조건 1단계로 현재 실행 중인 편집기를 찾고 내용을 먼저 확인합니다:\n"
        "     * 전체 실행 상태: office_get_status\n"
        "     * Word 선택영역/내용: word_read_document(scope='selection' 또는 'all')\n"
        "     * Excel 선택영역/내용: excel_read_data(use_selection=True)\n"
        "     * HWP 선택영역/내용: hwp_read_document(only_selection=True)\n"
        "   - 확인된 문서명과 선택 내용을 사용자에게 간결히 인지시킵니다.\n\n"
        "2. [2단계: 참조 양식 베이스 커스텀 양식 우선 원칙 (Reference-Based Customization First)]\n"
        "   - 문서에 표, 박스, 제목 바, 일정표, 결재선, 커스텀 양식을 만들 때는 절대 백지에서 임의의 서식이나 스크립트를 먼저 만들지 마십시오!\n"
        "   - 반드시 검증된 표준 참조 서식(283종)을 먼저 검색하고, 이를 베이스(Base)로 삼아 커스텀 데이터를 주입·확장하는 순서를 따릅니다:\n"
        "     * 1) 참조 양식 RAG 검색: office_list_forms(query='...')로 목적에 맞는 표준 양식(소제목, 시간계획, 결재선, 점선박스, 기본표 등)의 구조(cols, rows, has_merged_cells, literals)를 먼저 확인\n"
        "     * 2) [1순위] 참조 양식 위에서 즉시 커스텀 삽입: office_insert_form(form_name='...', args=[...], headers=[...], rows=[[...]], replacements={...})를 호출하여, 참조 양식의 원본 여백·테두리·이중선·병합 레이아웃을 100% 보존한 채 헤더/행 데이터 주입(행 자동 추가/정리) 및 고정 문구 치환을 한 번에 수행\n"
        "     * 3) [2순위 폴백] 열 개수 변경 또는 다중 블록 복합 양식 조립: 단일 참조 양식과 열 개수가 달라야 하거나 여러 블록을 조립해야 할 때만 office_build_custom_form(..., base_form='검색된참조양식')을 호출하여 참조 양식의 실제 AST 배경색·테두리 스타일을 상속받아 생성 (블록 내 {'type': 'reference_form', ...} 직접 포함 가능)\n"
        "     * 4) 미세 보정 및 톤앤매너 일괄 통일 (마무리): 필요한 경우 word_format_selection, word_format_table로 다듬고, 마지막에 office_harmonize_document를 호출하여 글꼴, 표 너비(본문 100% AutoFit), 폰트 크기 계층, 테마 색상을 한 번에 일괄 통일\n"
        "     * 조직도/계층도는 office_insert_hierarchy 한 번으로 3단 이상 완성되므로 직접 셀을 나누거나 덧붙이지 마십시오.\n\n"
        "3. [3단계: 오피스 4대 문서(PPT/Word/Excel/HWP) 단계적 분할 작성 원칙 (Chunked Generation)]\n"
        "   - 절대 한 번의 호출로 전체 문서(슬라이드 5~15장, 워드 수십 페이지, 엑셀 대형 데이터 등)를 한 번에 일괄 생성(One-Shot Monolithic Generation)하지 마십시오!\n"
        "   - 10분간 아무런 조작 없이 멈춰있는 현상은 사용자 경험을 심각하게 훼손하므로 엄격히 금지됩니다.\n"
        "   - 반드시 30초~1분 단위로 쪼개어 단계별로 실행하십시오:\n"
        "     * [PPT]: 먼저 목차 제안 -> 표지/목차 기본 덱 생성(1단계) -> 슬라이드 1장씩 순차 추가(ppt_compose_slide 또는 ppt_split_slide_layout)\n"
        "     * [Word]: 머리글/바닥글 및 제목 정돈(1단계) -> 본문 섹션별 작성(2단계) -> 표/콜아웃 박스 삽입(3단계)\n"
        "     * [Excel]: 시트 및 결재란 생성(1단계) -> 데이터 블록 주입(2단계) -> 서식/너비 자동 맞춤(3단계)\n"
        "     * [HWP]: 공문서 개요 작성(1단계) -> 조항/본문군 주입(2단계) -> 표 삽입(3단계)\n"
        "   - 호출 1회에는 최대 1~2개 단위만 처리하십시오. 단일 도구 호출에 3단위를 초과하여 밀어 넣지 마십시오.\n"
        "   - office_run_com_script는 50줄을 넘거나, Paragraphs/Shapes/Slides 루프 안에 if/elif가 3개 이상이면 실행되지 않고 거절됩니다(blocked=True).\n"
        "   - Word 문단은 word_format_selection(paragraph_pattern=정규식)으로 요소 1종씩 다듬고, 선택영역(Selection)과 표는 word_format_selection, word_write_text, word_format_table을 사용하십시오.\n\n"
        "4. [4단계: 매 단위 실행 직후 실시간 화면 갱신 및 진행률 보고]\n"
        "   - 각 단위를 실행할 때마다 창을 즉시 활성화(Activate)하여 사용자 눈앞에 결과가 추가되는 것을 보여주십시오.\n"
        "   - 채팅창에 '[2/5 완료] \"Ⅱ. 추진 배경\" 슬라이드가 적용되었습니다.' 형태로 명확한 진행률을 즉시 보고하십시오.\n\n"
        "5. [절대 금지: 마크다운 서식 기호(**, ##, __ 등) 입력 엄격 금지]\n"
        "   - Word, Excel, HWP 문서에 입력하는 text에는 '**굵게**'나 '### 제목', '- [x]' 같은 마크다운 기호를 절대 섞지 마십시오.\n"
        "   - 오피스 문서는 마크다운 뷰어가 아니므로 별표(**)가 문서에 그대로 타이핑됩니다. 반드시 순수 일반 텍스트(Plain Text)로 전달하십시오.\n\n"
        "6. [절대 금지: 소스코드 파일(*.py) 열람 금지]\n"
        "   - mcp_server/ 및 *.py 파일을 read_file로 뒤적거리지 말고, 등록된 MCP 도구를 즉시 호출하십시오.\n\n"
        "7. [핵심 원칙: 파이썬 스크립트 작성 우회 금지 -> 등록된 공식 MCP 도구 최우선 호출]\n"
        "   - 문서를 생성하거나 편집할 때 임의의 파이썬 스크립트(python -c, .py 파일 생성 후 실행 등)를 작성하여 우회하지 마십시오.\n"
        "   - 반드시 등록된 공식 MCP 도구(ppt_create_report_deck_shell, ppt_compose_slide, word_write_text, excel_write_data, hwp_write_styled_text, office_insert_form, ppt_clone_and_morph_slide, word_insert_sdt_control, excel_scaffold_financial_model 등)를 순차 직접 호출하여 실행하십시오.\n\n"
        "8. [5단계 납품 품질 감사 (5-Gate Delivery Gate) & 시각 스냅샷 검증]\n"
        "   - 문서 완성 직후 또는 납품 전 반드시 `office_run_delivery_gate`를 호출하여 토큰 누수(`{{...}}`, `**`), Excel `###`/`#REF!`, 저대비(Dark-on-Dark), 재무 불균형, Morph 유령 도형 잔존, Word `____` 가짜 양식을 차단하십시오.\n"
        "   - 템플릿 대량 치환은 `office_merge_template_tokens`, 승인 전 마킹은 `office_mark_review_item`, 시각 검증은 `office_export_visual_snapshot`을 활용하십시오."
    ),
)

# 도구 모듈 일괄 등록
register_system_tools(mcp)
register_word_tools(mcp)
register_excel_tools(mcp)
register_ppt_tools(mcp)
register_hwp_tools(mcp)
register_component_tools(mcp)
register_chart_tools(mcp)
register_dynamic_form_tools(mcp)
register_document_form_tools(mcp)
register_user_template_tools(mcp)
register_folder_tools(mcp)
register_composition_tools(mcp)
register_ppt_morph_tools(mcp)
register_word_form_academic_tools(mcp)
register_excel_modeling_tools(mcp)
register_delivery_gate_tools(mcp)


# =============================================================================
# MCP Resources (컨텍스트 및 서식 가이드 리소스)
# =============================================================================

@mcp.resource("office://status")
def resource_office_status() -> str:
    """현재 데스크톱의 Word, Excel, PowerPoint, HWP 실행 상태 및 활성 문서 정보 JSON"""
    return json.dumps(com_manager.get_status(), ensure_ascii=False, indent=2)


@mcp.resource("office://guides/kisa1310-hwp")
def resource_kisa1310_hwp_guide() -> str:
    """KISA 범피스 1310 기반 한컴오피스(HWP) 행안부/공기업 표준 공문서 작성 가이드"""
    return KISA1310_OFFICIAL_REPORT_GUIDE


@mcp.resource("office://guides/aionui-domain-playbooks")
def resource_aionui_domain_playbooks() -> str:
    """iOfficeAI(AionUi · AionCore · OfficeCLI) 10대 오피스 도메인 플레이북 및 5-Gate 납품 감사 요약"""
    return """# AionUi · OfficeCLI 오피스 도메인 플레이북 & 5-Gate 기준

## 1. PowerPoint Morph & Pitch Deck
- **3-Prefix 규칙**: `!!scene-*`(배경 유지), `!!actor-*`(전경 진화, 변위 >= 5cm / 회전 >= 15° / 크기변화 >= 30%), `#sN-*`(슬라이드 N 전용 콘텐츠 -> 다음 슬라이드 복제 시 삭제 대신 우측 바깥 `1020.5pt` 또는 `SlideWidth+60pt`로 고스팅).
- **Pitch Deck 5단계 x 5대 버티컬**: `ppt_plan_pitch_deck(stage, vertical)`로 단계별 아크 및 C.1~C.5b 레이아웃 생성. 연속 슬라이드 동일 패턴 금지, 콘텐츠 슬라이드 발표자 노트 필수, 차트 Y축 `y_min = 0`.

## 2. Word 리얼 양식(SDT) & 학술 논문
- **리얼 양식**: `____` 밑줄 가짜 입력란 금지. `word_insert_sdt_control`로 실제 ContentControl(Title + Tag 필수) 삽입 후, 마지막에 `word_set_form_protection(enable=True)`(`wdAllowOnlyFormFields=2`) 적용. `word_audit_form_fields`로 감사.
- **학술 논문**: `word_apply_academic_venue_style(venue='APA7'|'Chicago17'|'IEEE'|'MLA9')` 및 `word_insert_academic_caption_or_ref`로 `SEQ Figure/Table/Equation`·`PAGEREF`·참고문헌 내어쓰기(36pt) 적용. Figure 캡션은 하단, Table 캡션은 상단 배치.

## 3. Excel 3-Zone 재무 모델 & 적응형 대시보드
- **3-Zone 탭**: Inputs(`Assumptions` 노랑 `#FFC000`) -> Calc(`P&L`/`Balance Sheet`/`Cash Flow`/`DCF` 파랑 `#4472C4`, 하드코딩 숫자 0개) -> Outputs(`Summary`/`Dashboard`/`Sensitivity` 초록 `#70AD47`, `Print_Area`).
- **CFO 5색 코딩**: `excel_apply_cfo_color_coding` — 파랑 `#0000FF`(입력), 검정 `#000000`(동일시트 수식), 초록 `#008000`(시트간 참조 `!`), 빨강 `#FF0000`(외부 링크 `[`), 노랑 채움 `#FFFF00`(핵심 가정).
- **대시보드 열너비 공식**: KPI `max(22, ceil((visible_chars + 2) * 1.3))`, 차트 폭 `ceil(len(title) * 0.18)`.

## 4. 전사 5-Gate 납품 감사 & Merge / Mark / Snapshot
- `office_run_delivery_gate`: Gate 1(구조) -> Gate 2(토큰 `{{}}`/마크다운 `**`/수식에러 `#REF!`/`###`) -> Gate 3(발표자 노트/TOC·PAGE/Print_Area) -> Gate 4(다크 저대비/`IMBALANCED`/Calc 하드코딩) -> Gate 5(Morph 고스트/Word 밑줄).
- `office_merge_template_tokens`: 서식 파괴 없는 `{{key}}` 일괄 병합.
- `office_mark_review_item` / `office_export_visual_snapshot`: 비파괴 리뷰 메모 + PNG/PDF 시각 스냅샷 검증.
"""


# =============================================================================
# MCP Prompts (대화형 읽기·작성·서식 적용 워크플로우 프롬프트)
# =============================================================================

@mcp.prompt()
def polish_active_selection(target_style: str = "비즈니스 네이비 보고서 서식") -> str:
    """현재 열려 있는 오피스 프로그램에서 사용자가 마우스로 선택한 영역을 읽고 문맥을 다듬은 뒤 원하는 서식으로 적용합니다."""
    return (
        f"1. 먼저 `office_get_status()`를 호출하여 현재 어떤 오피스 앱(Word, Excel, PowerPoint, HWP)이 열려 있는지 확인하세요.\n"
        f"2. 해당 앱의 선택 영역 읽기 도구(`word_read_document(scope='selection')`, `excel_read_data(use_selection=True)`, "
        f"`ppt_read_presentation(only_selection=True)`, `hwp_read_document(only_selection=True)`)를 호출하여 사용자가 선택한 원문을 읽으세요.\n"
        f"3. 원문의 핵심 의미를 유지하면서 전문적이고 명확한 비즈니스 문체로 다듬으세요.\n"
        f"4. 다듬은 내용과 함께 '{target_style}'에 맞는 작성/서식 도구를 호출하여 실시간으로 문서에 반영하고 결과를 설명해 주세요."
    )


@mcp.prompt()
def draft_official_hwp_report(topic: str, department: str = "디지털혁신담당관") -> str:
    """주제를 바탕으로 한컴오피스(HWP)에 행안부/공기업 표준 공문서 서식과 표가 포함된 보고서를 실시간 작성합니다."""
    return (
        f"주제 '{topic}'(담당부서: {department})에 대해 한컴오피스(HWP) 공문서 보고서를 작성해 주세요.\n"
        f"`hwp_render_kisa1310_report` 도구를 사용하여 `제목:`, `부서:`, `네모:`, `원:`, `바:`, `별:`, 그리고 `표:`(마지막 열은 '비고') 규격으로 "
        f"구조화된 보고서를 렌더링해 주세요."
    )


@mcp.prompt()
def create_dynamic_business_form(topic: str, form_type: str = "기획서", target_app: str = "word") -> str:
    """자연어로 요청된 업무 서식(회의록/기획서/점검표/견적서 등)을 참조 양식 베이스로 커스텀하여 Word 또는 Excel에 생성합니다."""
    return (
        f"주제 '{topic}'에 대해 {target_app.upper()}용 '{form_type}' 맞춤형 서식을 생성해 주세요.\n"
        f"1. 먼저 `office_list_forms(query='{form_type}')`로 적합한 표준 참조 양식을 검색하세요.\n"
        f"2. 단일 표/양식으로 해결 가능하면 1순위로 `office_insert_form(form_name=..., headers=[...], rows=[[...]], replacements={{...}})`를 호출하여 참조 양식 위에 바로 커스텀 데이터를 주입하세요.\n"
        f"3. 다중 블록 조립이 필요한 복합 서식이라면 `office_build_custom_form(..., base_form='검색된참조양식')`을 호출하되, 다음 블록들을 조합하여 `blocks` 리스트를 구성하세요:\n"
        f"0. `reference_form`: 표준 참조 양식 직접 삽입 및 커스텀 (form_name, args, headers, rows, replacements)\n"
        f"1. `title_header`: 대제목, 서브제목, 문서번호, 일자\n"
        f"2. `approval_line`: 결재란 (예: ['기안', '팀장', '본부장', '대표이사'])\n"
        f"3. `key_value_grid`: 2열 또는 4열 메타데이터 표 (일시, 장소, 참석자, 부서)\n"
        f"4. `callout_box`: 핵심 요약/주의/안내 상자 (box_type='info'|'tip'|'warning')\n"
        f"5. `table`: 겉바속촉 스마트 표 (headers, rows, 숫자 우측정렬)\n"
        f"6. `comparison`: AS-IS vs TO-BE 대칭 비교 카드\n"
        f"7. `checklist`: 점검 과제, 담당자, 마감기한, 완료여부 배지\n"
        f"8. `process_flow`: 단계별 카드 + ➜ 화살표\n"
        f"9. `kpi_cards`: 핵심 성과 지표 대시보드 카드\n"
        f"10. `signature`: 하단 선언문 + 서명란\n"
        f"4. 마무리로 `office_harmonize_document`를 호출해 문서 전체 톤앤매너를 통일해 주세요."
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="MyOffice Desktop COM & KISA1310 MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="MCP 통신 프로토콜 선택 (기본값: stdio)",
    )
    parser.add_argument("--host", default="127.0.0.1", help="SSE 모드 호스트 주소 (기본값: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8765, help="SSE 모드 포트 번호 (기본값: 8765)")
    args = parser.parse_args()

    if args.transport == "sse":
        if hasattr(mcp, "settings") and hasattr(mcp.settings, "host"):
            mcp.settings.host = args.host
            mcp.settings.port = args.port
            mcp.run(transport="sse")
        else:
            mcp.run(transport="sse", host=args.host, port=args.port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
