# -*- coding: utf-8 -*-
"""
mcp_server/tools/system_tools.py
오피스 4종(Word, Excel, PowerPoint, HWP) 실행 상태 감지, 앱 기동/파일 열기, 동적 COM 스크립트 실행 도구.
"""

from __future__ import annotations

import ast
import contextlib
import io
import traceback
from typing import Any

import win32com.client as win32
from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr, rgb_to_bgr

# office_run_com_script 1회 호출 한도 (1회 호출 = 요소 1종)
MAX_SCRIPT_LINES = 50
MAX_LOOP_BRANCHES = 2
BULK_COLLECTIONS = ("Paragraphs", "Shapes", "Slides")
SPLIT_HINT = (
    "요소 1종(예: 섹션 제목만)씩 나눠서 호출하고, 단계마다 화면을 보여주며 보고하세요. "
    "Word 문단은 word_format_selection(paragraph_pattern=...)을 쓰면 스크립트가 필요 없습니다."
)


def _check_script_scope(script_code: str) -> str | None:
    """한 번에 여러 요소를 처리하는 스크립트면 거절 사유를, 괜찮으면 None을 반환"""
    line_count = sum(1 for line in script_code.splitlines() if line.strip())
    if line_count > MAX_SCRIPT_LINES:
        return f"스크립트가 {line_count}줄로 한도({MAX_SCRIPT_LINES}줄)를 넘어 실행하지 않았습니다. {SPLIT_HINT}"

    try:
        tree = ast.parse(script_code)
    except SyntaxError:
        return None  # 문법 오류는 exec 쪽 오류 응답으로 처리

    # ponytail: for문의 iter 표현식만 검사, 컬렉션을 변수에 담아 돌리면 줄 수 한도만 적용됨
    for node in ast.walk(tree):
        if isinstance(node, ast.For) and any(name in ast.unparse(node.iter) for name in BULK_COLLECTIONS):
            branches = sum(isinstance(child, ast.If) for child in ast.walk(node))
            if branches > MAX_LOOP_BRANCHES:
                return (
                    f"한 루프에서 요소를 {branches}개 분기로 나눠 처리하려 해 실행하지 않았습니다"
                    f"(한도 {MAX_LOOP_BRANCHES}개). {SPLIT_HINT}"
                )
    return None


def register_system_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def office_get_status() -> dict[str, Any]:
        """
        현재 데스크톱의 4대 오피스 프로그램(MS Word, Excel, PowerPoint, 한컴 HWP)의
        설치 여부, 실시간 실행 여부, 현재 열려 있는 활성 문서/시트/슬라이드 이름,
        그리고 사용자가 마우스로 선택한 영역(Selection) 미리보기를 한눈에 조회합니다.
        작업을 시작하기 전에 어떤 문서가 열려 있는지 파악할 때 가장 먼저 호출하세요.
        """
        return com_manager.get_status()

    @mcp.tool()
    def office_launch_or_open(
        app_name: str,
        file_path: str | None = None,
        visible: bool = True,
    ) -> dict[str, Any]:
        """
        지정한 오피스 프로그램(word, excel, ppt, hwp)을 실행하거나 특정 문서 파일을 데스크톱 창에 엽니다.

        Args:
            app_name: 실행할 앱 종류 ("word", "excel", "ppt", "hwp" 중 하나)
            file_path: 열 문서의 파일 경로 (생략 시 빈 새 문서로 앱 활성화)
            visible: 화면에 창을 표시할지 여부 (기본값 True)
        """
        ensure_com()
        key = app_name.lower().strip()
        if file_path:
            return com_manager.open_file_in_desktop(key, file_path, visible=visible)

        if key in ("word", "docx"):
            app, doc = com_manager.get_word(create_if_missing=True, visible=visible)
            return {
                "success": app is not None,
                "app": "Word",
                "active_document": doc.Name if doc else None,
            }
        elif key in ("excel", "xlsx"):
            app, wb, sheet = com_manager.get_excel(create_if_missing=True, visible=visible)
            return {
                "success": app is not None,
                "app": "Excel",
                "active_workbook": wb.Name if wb else None,
                "active_sheet": sheet.Name if sheet else None,
            }
        elif key in ("ppt", "powerpoint", "pptx"):
            app, prs = com_manager.get_ppt(create_if_missing=True, visible=visible)
            return {
                "success": app is not None,
                "app": "PowerPoint",
                "active_presentation": prs.Name if prs else None,
            }
        elif key in ("hwp", "hwpx", "hancom"):
            hwp = com_manager.get_hwp(create_if_missing=True, visible=visible)
            return {
                "success": hwp is not None,
                "app": "HWP",
                "page_count": hwp.PageCount if hwp else 0,
            }
        return {"success": False, "error": f"지원하지 않는 앱 이름입니다: {app_name}"}

    @mcp.tool()
    def office_run_com_script(
        app_name: str,
        script_code: str,
    ) -> dict[str, Any]:
        """
        고수준 도구로 지원되지 않는 세부적인 COM/VBA 제어 로직을 파이썬 코드로 직접 실행합니다.
        1회 호출에는 요소 1종만 처리하세요. 다음 스크립트는 실행하지 않고 거절합니다(blocked=True):
        - 빈 줄을 뺀 50줄 초과
        - Paragraphs/Shapes/Slides를 도는 for문 안에 if/elif 3개 이상
        Word 표는 병합 셀이 있으면 Columns.Count만큼 Cell(r, c)가 없어 오류 5941이 납니다.
        표를 읽을 때는 `for c in tbl.Range.Cells:`로 순회하세요. 병합 행의 c.ColumnIndex는 화면 열 번호가 아니라 그 행의 셀 순번입니다.
        병합 여부는 tbl.Uniform이 아니라 tbl.Range.Cells.Count와 Rows.Count * Columns.Count를 비교해 판단하세요.
        앱마다 값과 동작이 다르니 주의하세요:
        - 가운데 정렬: Word 1, Excel -4108, PPT 2
        - 테두리: Word 왼쪽 -2, 아래 -3 / Excel 왼쪽 7, 아래 9
        - 선 굵기: Word LineWidth 4 = 0.5pt, Excel Weight 4 = 가장 굵음, PPT Weight = 실제 pt
        - 서식이 섞인 값: Word 9999999, PPT -2
        - PPT 문단 구분은 \\r (\\n은 문단을 나누지 않음), 도형 선 색은 Line.ForeColor.RGB
        - Word 고정 줄간격은 LineSpacingRule = 4를 먼저 설정 (없으면 18이 1.5줄로 해석됨)
        실행 컨텍스트에 다음 변수가 미리 주입되어 있습니다:
        - `word`, `doc` (app_name="word" 일 때)
        - `excel`, `wb`, `sheet` (app_name="excel" 일 때)
        - `ppt`, `prs` (app_name="ppt" 일 때)
        - `hwp` (app_name="hwp" 일 때)
        - `win32`, `com_manager`, `hex_to_bgr`, `rgb_to_bgr`
        코드 내에서 `result = ...` 변수에 값을 할당하거나 `print(...)`를 호출하면 결과로 반환됩니다.
        """
        scope_error = _check_script_scope(script_code)
        if scope_error:
            return {"success": False, "blocked": True, "error": scope_error}

        ensure_com()
        key = app_name.lower().strip()
        local_ns: dict[str, Any] = {
            "win32": win32,
            "com_manager": com_manager,
            "hex_to_bgr": hex_to_bgr,
            "rgb_to_bgr": rgb_to_bgr,
            "result": None,
        }

        if key in ("word", "docx"):
            word_app, doc = com_manager.get_word(create_if_missing=True, visible=True)
            local_ns["word"] = word_app
            local_ns["word_app"] = word_app
            local_ns["app"] = word_app
            local_ns["doc"] = doc
            local_ns["active_doc"] = doc
            try:
                local_ns["sel"] = word_app.Selection if word_app else None
                local_ns["selection"] = word_app.Selection if word_app else None
            except Exception:
                local_ns["sel"] = None
                local_ns["selection"] = None
        elif key in ("excel", "xlsx"):
            excel_app, wb, sheet = com_manager.get_excel(create_if_missing=True, visible=True)
            local_ns["excel"] = excel_app
            local_ns["excel_app"] = excel_app
            local_ns["app"] = excel_app
            local_ns["wb"] = wb
            local_ns["sheet"] = sheet
            local_ns["active_sheet"] = sheet
        elif key in ("ppt", "powerpoint", "pptx"):
            ppt_app, prs = com_manager.get_ppt(create_if_missing=True, visible=True)
            local_ns["ppt"] = ppt_app
            local_ns["ppt_app"] = ppt_app
            local_ns["app"] = ppt_app
            local_ns["prs"] = prs
            local_ns["presentation"] = prs
            try:
                local_ns["slide"] = ppt_app.ActiveWindow.View.Slide if ppt_app else None
            except Exception:
                local_ns["slide"] = None
        elif key in ("hwp", "hwpx", "hancom"):
            hwp = com_manager.get_hwp(create_if_missing=True, visible=True)
            local_ns["hwp"] = hwp
            local_ns["hwp_app"] = hwp
            local_ns["app"] = hwp

        stdout_buf = io.StringIO()
        try:
            # 함수 내부에서도 word, doc, sel 등에 접근할 수 있도록 globals에도 함께 주입
            exec_globals = {"__builtins__": __builtins__}
            exec_globals.update(local_ns)
            with contextlib.redirect_stdout(stdout_buf):
                exec(script_code, exec_globals, local_ns)
            return {
                "success": True,
                "result": local_ns.get("result"),
                "stdout": stdout_buf.getvalue().strip(),
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc(limit=4),
                "stdout": stdout_buf.getvalue().strip(),
            }
