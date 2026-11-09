# -*- coding: utf-8 -*-
"""
mcp_server/tools/component_tools.py
KISA 범피스(1310/960) 14만 줄 원천 알고리즘을 4대 오피스(Word, Excel, PowerPoint, HWP)용으로
완전 승화·이식한 7대 비즈니스 컴포넌트 빌더 및 스마트 서식 정제 도구.

[7대 비즈니스 컴포넌트]
1. office_insert_hierarchy: 계층 구조도/조직도 (투명 격자 표 + 병합 카드 + 칸 테두리 연결선)
2. office_insert_process_flow: 단계별 프로세스/업무 흐름도 (카드 + ➜ 화살표 2열 그리드)
3. office_insert_timeline: 추진일정/타임라인 로드맵 (간트 차트형 마일스톤 표 + 상태 배지)
4. office_insert_comparison: AS-IS vs TO-BE 대안/대비표 (대칭 2단 카드 + 차이점 강조)
5. office_insert_callout_box: 핵심 요약/콜아웃 박스 (좌측 3pt 강조선 + 은은한 배경음영)
6. office_insert_kpi_cards: 핵심 지표 KPI 대시보드 카드 (수치 + 라벨 + 증감율 배지)
7. office_insert_smart_table: 겉바속촉 프리미엄 비즈니스 표 (상하 굵은선/오픈 테이블/헤더 음영)

[2대 스마트 정제 도구]
8. word_clean_and_structure: Word 문서 궁극의 7단계 개조식 위계 및 자간/줄간격 교정
9. excel_auto_format_table: Excel 시트 겉바속촉 자동 서식화 (헤더/외곽선/천단위 콤마/AutoFit)
"""

from __future__ import annotations

import logging
import re
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr, word_undo_group
from .parsers import strip_markdown_syntax

logger = logging.getLogger("mcp_server.component_tools")

# 비즈니스 공통 컬러 팔레트 (BGR 정수 및 RGB)
COLOR_NAVY_BGR = hex_to_bgr("#1F3864")
COLOR_STEEL_BLUE_BGR = hex_to_bgr("#2E608A")
COLOR_EMERALD_BGR = hex_to_bgr("#16A34A")
COLOR_TEAL_BGR = hex_to_bgr("#0F766E")
COLOR_LIGHT_BG_BGR = hex_to_bgr("#F4F6F9")
COLOR_LIGHT_BLUE_BGR = hex_to_bgr("#EBF3FA")
COLOR_LIGHT_GREEN_BGR = hex_to_bgr("#EBF7EE")
COLOR_LIGHT_AMBER_BGR = hex_to_bgr("#FEF6E7")
COLOR_BORDER_GRAY_BGR = hex_to_bgr("#D9D9D9")
COLOR_TEXT_DARK_BGR = hex_to_bgr("#262626")
COLOR_WHITE_BGR = hex_to_bgr("#FFFFFF")

# Word 표 테두리 방향 (WdBorderType, Excel의 XlBordersIndex와 값이 다름)
WD_BORDER_TOP = -1
WD_BORDER_LEFT = -2
WD_BORDER_BOTTOM = -3
WD_BORDER_RIGHT = -4

# 계층도 격자 (범피스 조직도 방식: 투명 칸 위에 카드를 병합해 얹고 연결선은 칸 테두리로 긋는다)
HIERARCHY_CARD_COLS = 4  # 짝수라야 칸 경계가 카드 정중앙에 옴
HIERARCHY_GAP_COLS = 2  # 짝수라야 전체 칸 수가 짝수가 되어 대표 카드도 정중앙에 옴
HIERARCHY_ROOT_COLS = 6
HIERARCHY_MAX_BRANCHES = 7  # 8개면 46칸이라 A4 기본 여백에서 Word 최소 칸 너비(약 10.9pt)보다 좁아진다
# 깊이별 카드 서식: (배경, 테두리, 글자색, 크기, 굵게)
HIERARCHY_STYLES = [
    (COLOR_NAVY_BGR, COLOR_NAVY_BGR, COLOR_WHITE_BGR, 13, True),
    (COLOR_LIGHT_BLUE_BGR, COLOR_STEEL_BLUE_BGR, COLOR_NAVY_BGR, 11, True),
    (COLOR_LIGHT_BG_BGR, COLOR_STEEL_BLUE_BGR, COLOR_TEXT_DARK_BGR, 10, False),
    (COLOR_WHITE_BGR, COLOR_BORDER_GRAY_BGR, hex_to_bgr("#6B7280"), 9.5, False),
]


def _node_parts(node: str | dict[str, Any]) -> tuple[str, list[Any]]:
    """계층도 노드('이름' 또는 {'name': ..., 'children': [...]})를 (이름, 하위 노드)로 분해"""
    if isinstance(node, dict):
        return strip_markdown_syntax(str(node.get("name", ""))), node.get("children") or []
    return strip_markdown_syntax(str(node)), []


def _flatten_nodes(nodes: list[Any], depth: int = 2) -> list[tuple[str, int]]:
    """가지 아래 노드를 세로로 쌓을 순서대로 펼쳐 [(이름, 깊이)]로 반환"""
    flat: list[tuple[str, int]] = []
    for node in nodes:
        name, children = _node_parts(node)
        flat.append((name, depth))
        flat += _flatten_nodes(children, depth + 1)
    return flat


def _hierarchy_layout(branch_count: int) -> tuple[int, list[int], list[int], int]:
    """(전체 칸 수, 가지별 시작 칸, 가지별 중앙 칸, 전체 중앙 칸) 반환. 중앙 칸은 왼쪽 경계가 중앙선인 칸"""
    step = HIERARCHY_CARD_COLS + HIERARCHY_GAP_COLS
    cols = branch_count * step - HIERARCHY_GAP_COLS
    starts = [1 + i * step for i in range(branch_count)]
    centers = [start + HIERARCHY_CARD_COLS // 2 for start in starts]
    return cols, starts, centers, cols // 2 + 1


def _grid_line(cell: Any, side: int) -> None:
    border = cell.Borders(side)
    border.LineStyle = 1
    border.LineWidth = 8
    border.Color = COLOR_STEEL_BLUE_BGR


def _grid_card(tbl: Any, row: int, first_col: int, last_col: int, text: str, depth: int) -> None:
    """first_col~last_col을 병합해 카드를 만든다. 번호가 밀리지 않도록 같은 행은 오른쪽 카드부터 호출"""
    fill, border_color, font_color, size, bold = HIERARCHY_STYLES[min(depth, len(HIERARCHY_STYLES) - 1)]
    tbl.Cell(row, first_col).Merge(tbl.Cell(row, last_col))
    cell = tbl.Cell(row, first_col)
    cell.Range.Text = text
    cell.Range.Font.Name = "맑은 고딕"
    cell.Range.Font.Size = size
    cell.Range.Font.Bold = bold
    cell.Range.Font.Color = font_color
    cell.Range.ParagraphFormat.Alignment = 1  # 가운데
    cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
    cell.Shading.BackgroundPatternColor = fill
    for side in (WD_BORDER_TOP, WD_BORDER_LEFT, WD_BORDER_BOTTOM, WD_BORDER_RIGHT):
        border = cell.Borders(side)
        border.LineStyle = 1
        border.LineWidth = 8
        border.Color = border_color


# =============================================================================
# 1. Word COM 컴포넌트 렌더러
# =============================================================================
class WordComponentRenderer:
    """Word 문서 내 표 그리드 기반 비즈니스 컴포넌트 생성"""

    @staticmethod
    def insert_hierarchy(word_app: Any, doc: Any, title: str, levels: list[Any]) -> str:
        """조직도/구조도: 투명 격자 위에 카드를 병합해 얹고 연결선은 칸 테두리로 긋는다 (범피스 조직도 방식)"""
        branches = [(name, _flatten_nodes(children)) for name, children in map(_node_parts, levels)]
        n = len(branches)
        if not 1 <= n <= HIERARCHY_MAX_BRANCHES:
            return f"계층도의 하위 가지는 1~{HIERARCHY_MAX_BRANCHES}개여야 합니다. 현재 {n}개입니다."

        cols, starts, centers, mid = _hierarchy_layout(n)
        tail = max(len(nodes) for _, nodes in branches)
        rows = 4 + tail * 2  # 대표, 줄기, 분기선, 가지 + (연결선, 하위 카드) * 하위 단수

        sel = word_app.Selection
        page = sel.Sections(1).PageSetup  # 구역이 여러 개면 doc.PageSetup 값은 wdUndefined
        text_width = page.PageWidth - page.LeftMargin - page.RightMargin

        tbl = doc.Tables.Add(sel.Range, rows, cols)
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False  # 기본 전체 투명
        tbl.TopPadding = tbl.BottomPadding = tbl.LeftPadding = tbl.RightPadding = 0
        tbl.Range.Font.Size = 2  # 빈 칸 글자 크기가 연결선 행 높이를 밀어 올리지 않도록
        tbl.Range.ParagraphFormat.SpaceBefore = 0
        tbl.Range.ParagraphFormat.SpaceAfter = 0
        tbl.Range.ParagraphFormat.LineSpacingRule = 0  # wdLineSpaceSingle
        for c in range(1, cols + 1):
            tbl.Columns(c).Width = text_width / cols

        for r in range(1, rows + 1):
            is_card = r in (1, 4) or (r > 4 and r % 2 == 0)
            tbl.Rows(r).HeightRule = 1 if is_card else 2  # 카드는 wdRowHeightAtLeast(글자 잘림 방지), 연결선은 Exactly
            tbl.Rows(r).Height = {1: 34, 4: 30}.get(r, 26) if is_card else 11

        # 연결선은 병합 전에 격자 칸 테두리로 긋는다
        _grid_line(tbl.Cell(2, mid), WD_BORDER_LEFT)
        if n > 1:
            for c in range(centers[0], centers[-1]):
                _grid_line(tbl.Cell(3, c), WD_BORDER_TOP)
        for i, (_, nodes) in enumerate(branches):
            _grid_line(tbl.Cell(3, centers[i]), WD_BORDER_LEFT)
            for j in range(len(nodes)):
                _grid_line(tbl.Cell(5 + j * 2, centers[i]), WD_BORDER_LEFT)

        # 카드는 같은 행의 칸 번호가 밀리지 않도록 오른쪽 가지부터 병합
        for i, (name, nodes) in reversed(list(enumerate(branches))):
            last_col = starts[i] + HIERARCHY_CARD_COLS - 1
            for j, (node_name, depth) in enumerate(nodes):
                _grid_card(tbl, 6 + j * 2, starts[i], last_col, node_name, depth)
            _grid_card(tbl, 4, starts[i], last_col, name, 1)
        half = min(HIERARCHY_ROOT_COLS, cols) // 2
        _grid_card(tbl, 1, mid - half, mid + half - 1, strip_markdown_syntax(title), 0)

        # 표 탈출 후 단락 추가
        sel.Start = tbl.Range.End
        sel.InsertParagraphAfter()
        sel.MoveDown(Unit=5, Count=1)
        return f"Word 문서에 {n}개 분기 계층 구조도('{title}')가 생성되었습니다."

    @staticmethod
    def insert_process_flow(word_app: Any, doc: Any, steps: list[dict[str, str]]) -> str:
        """프로세스 흐름도: [카드] ➜ [카드] ➜ [카드]"""
        sel = word_app.Selection
        n = len(steps)
        cols = 2 * n - 1
        rows = 2  # 1행: 단계명, 2행: 세부내용

        tbl = doc.Tables.Add(sel.Range, rows, cols)
        tbl.Borders.Enable = False

        for i, step in enumerate(steps):
            c_idx = 2 * i + 1
            title = step.get("title", f"단계 {i+1}")
            desc = step.get("desc", "")

            # 1행 단계명 카드
            c1 = tbl.Cell(1, c_idx)
            c1.Range.Text = f"STEP {i+1}. {title}"
            c1.Range.Font.Name = "맑은 고딕"
            c1.Range.Font.Size = 11
            c1.Range.Font.Bold = True
            c1.Range.Font.Color = COLOR_WHITE_BGR
            c1.Range.ParagraphFormat.Alignment = 1
            c1.Shading.BackgroundPatternColor = COLOR_NAVY_BGR

            # 2행 설명 카드
            c2 = tbl.Cell(2, c_idx)
            c2.Range.Text = desc
            c2.Range.Font.Name = "맑은 고딕"
            c2.Range.Font.Size = 9.5
            c2.Range.Font.Color = COLOR_TEXT_DARK_BGR
            c2.Range.ParagraphFormat.Alignment = 0  # 왼쪽
            c2.Shading.BackgroundPatternColor = COLOR_LIGHT_BG_BGR
            for b in (-1, -2, -3, -4):
                c2.Borders(b).LineStyle = 1
                c2.Borders(b).LineWidth = 4
                c2.Borders(b).Color = COLOR_BORDER_GRAY_BGR

            # 카드 사이 화살표
            if i < n - 1:
                arrow_idx = 2 * i + 2
                a_cell = tbl.Cell(1, arrow_idx)
                a_cell.Range.Text = "➜"
                a_cell.Range.Font.Name = "Segoe UI Symbol"
                a_cell.Range.Font.Size = 16
                a_cell.Range.Font.Bold = True
                a_cell.Range.Font.Color = COLOR_STEEL_BLUE_BGR
                a_cell.Range.ParagraphFormat.Alignment = 1

        sel.Start = tbl.Range.End
        sel.InsertParagraphAfter()
        sel.MoveDown(Unit=5, Count=1)
        return f"Word 문서에 {n}단계 프로세스 흐름도가 생성되었습니다."

    @staticmethod
    def insert_timeline(word_app: Any, doc: Any, title: str, milestones: list[dict[str, str]]) -> str:
        """추진일정/타임라인 로드맵: 마일스톤 + 기간 + 과제 + 상태 배지 간트형 표"""
        sel = word_app.Selection
        rows = len(milestones) + 2  # 제목행 + 헤더행 + 데이터행들
        cols = 4

        tbl = doc.Tables.Add(sel.Range, rows, cols)
        tbl.Borders.Enable = False

        # 1행: 전체 제목 병합
        c_title = tbl.Cell(1, 1)
        if cols > 1:
            c_title.Merge(tbl.Cell(1, cols))
        c_title.Range.Text = f"■ {title}"
        c_title.Range.Font.Name = "맑은 고딕"
        c_title.Range.Font.Size = 12
        c_title.Range.Font.Bold = True
        c_title.Range.Font.Color = COLOR_NAVY_BGR
        c_title.Borders(WD_BORDER_BOTTOM).LineStyle = 1
        c_title.Borders(WD_BORDER_BOTTOM).LineWidth = 12
        c_title.Borders(WD_BORDER_BOTTOM).Color = COLOR_NAVY_BGR

        # 2행: 헤더
        headers = ["구분 / 마일스톤", "추진 일정", "주요 과제 및 핵심 산출물", "진행 현황"]
        col_widths = [110, 90, 200, 70]
        for c, h in enumerate(headers, 1):
            cell = tbl.Cell(2, c)
            cell.Range.Text = h
            cell.Range.Font.Name = "맑은 고딕"
            cell.Range.Font.Size = 10
            cell.Range.Font.Bold = True
            cell.Range.Font.Color = COLOR_WHITE_BGR
            cell.Shading.BackgroundPatternColor = COLOR_NAVY_BGR
            cell.Range.ParagraphFormat.Alignment = 1
            cell.Width = col_widths[c - 1]

        # 데이터 행
        for r, ms in enumerate(milestones, 1):
            row_idx = r + 2
            phase = ms.get("phase", f"{r}단계")
            period = ms.get("period", "-")
            task = ms.get("task", "-")
            status = ms.get("status", "진행중")

            c1 = tbl.Cell(row_idx, 1)
            c1.Range.Text = phase
            c1.Range.Font.Name = "맑은 고딕"
            c1.Range.Font.Size = 9.5
            c1.Range.Font.Bold = True
            c1.Range.ParagraphFormat.Alignment = 1
            c1.Width = col_widths[0]

            c2 = tbl.Cell(row_idx, 2)
            c2.Range.Text = period
            c2.Range.Font.Name = "맑은 고딕"
            c2.Range.Font.Size = 9.5
            c2.Range.ParagraphFormat.Alignment = 1
            c2.Width = col_widths[1]

            c3 = tbl.Cell(row_idx, 3)
            c3.Range.Text = task
            c3.Range.Font.Name = "맑은 고딕"
            c3.Range.Font.Size = 9.5
            c3.Range.ParagraphFormat.Alignment = 0
            c3.Width = col_widths[2]

            c4 = tbl.Cell(row_idx, 4)
            c4.Range.Text = f"[{status}]"
            c4.Range.Font.Name = "맑은 고딕"
            c4.Range.Font.Size = 9
            c4.Range.Font.Bold = True
            c4.Range.ParagraphFormat.Alignment = 1
            c4.Width = col_widths[3]

            # 상태별 색상
            if "완료" in status:
                c4.Range.Font.Color = hex_to_bgr("#059669")
            elif "진행" in status:
                c4.Range.Font.Color = hex_to_bgr("#2563EB")
            else:
                c4.Range.Font.Color = hex_to_bgr("#6B7280")

            # 줄무늬 배경 및 가로 구분선
            bg = COLOR_LIGHT_BG_BGR if r % 2 == 1 else COLOR_WHITE_BGR
            for c in range(1, cols + 1):
                cell = tbl.Cell(row_idx, c)
                cell.Shading.BackgroundPatternColor = bg
                cell.Borders(WD_BORDER_BOTTOM).LineStyle = 1
                cell.Borders(WD_BORDER_BOTTOM).LineWidth = 4
                cell.Borders(WD_BORDER_BOTTOM).Color = COLOR_BORDER_GRAY_BGR

        sel.Start = tbl.Range.End
        sel.InsertParagraphAfter()
        sel.MoveDown(Unit=5, Count=1)
        return f"Word 문서에 추진일정 타임라인 표('{title}')가 생성되었습니다."

    @staticmethod
    def insert_smart_table(word_app: Any, doc: Any, headers: list[str], rows: list[list[Any]]) -> str:
        """겉바속촉 프리미엄 비즈니스 표: 상하 굵은선, 헤더 이중선, 오픈 테이블(세로선 없음)"""
        sel = word_app.Selection
        num_rows = len(rows) + 1
        num_cols = len(headers)

        tbl = doc.Tables.Add(sel.Range, num_rows, num_cols)
        tbl.Borders.Enable = False

        # 1행 헤더
        for c, h in enumerate(headers, 1):
            cell = tbl.Cell(1, c)
            cell.Range.Text = str(h)
            cell.Range.Font.Name = "맑은 고딕"
            cell.Range.Font.Size = 10
            cell.Range.Font.Bold = True
            cell.Range.Font.Color = COLOR_WHITE_BGR
            cell.Shading.BackgroundPatternColor = COLOR_NAVY_BGR
            cell.Range.ParagraphFormat.Alignment = 1
            # 상단 굵은 실선
            cell.Borders(-1).LineStyle = 1
            cell.Borders(-1).LineWidth = 12
            cell.Borders(-1).Color = COLOR_NAVY_BGR
            # 헤더 하단 이중선
            cell.Borders(WD_BORDER_BOTTOM).LineStyle = 7  # wdLineStyleDouble = 7
            cell.Borders(WD_BORDER_BOTTOM).LineWidth = 6
            cell.Borders(WD_BORDER_BOTTOM).Color = COLOR_NAVY_BGR

        # 데이터 행
        for r_idx, r_data in enumerate(rows, 2):
            for c_idx, val in enumerate(r_data, 1):
                if c_idx > num_cols:
                    break
                cell = tbl.Cell(r_idx, c_idx)
                s_val = str(val)
                cell.Range.Text = s_val
                cell.Range.Font.Name = "맑은 고딕"
                cell.Range.Font.Size = 9.5
                # 숫자 판별 정렬
                clean_num = s_val.replace(",", "").replace("%", "").strip()
                if clean_num.replace(".", "", 1).isdigit() or (clean_num.startswith("-") and clean_num[1:].replace(".", "", 1).isdigit()):
                    cell.Range.ParagraphFormat.Alignment = 2  # 우측 정렬
                else:
                    cell.Range.ParagraphFormat.Alignment = 0  # 좌측 정렬

                # 내부 가로선
                cell.Borders(WD_BORDER_BOTTOM).LineStyle = 1
                cell.Borders(WD_BORDER_BOTTOM).LineWidth = 4
                cell.Borders(WD_BORDER_BOTTOM).Color = COLOR_BORDER_GRAY_BGR

        # 맨 하단 굵은 실선 마감
        last_row = tbl.Rows(num_rows)
        for c in range(1, num_cols + 1):
            last_row.Cells(c).Borders(WD_BORDER_BOTTOM).LineStyle = 1
            last_row.Cells(c).Borders(WD_BORDER_BOTTOM).LineWidth = 12
            last_row.Cells(c).Borders(WD_BORDER_BOTTOM).Color = COLOR_NAVY_BGR

        sel.Start = tbl.Range.End
        sel.InsertParagraphAfter()
        sel.MoveDown(Unit=5, Count=1)
        return f"Word 문서에 {num_rows}행 {num_cols}열 '겉바속촉' 비즈니스 표가 생성되었습니다."

    @staticmethod
    def insert_comparison(word_app: Any, doc: Any, title: str, as_is: list[str], to_be: list[str]) -> str:
        """AS-IS vs TO-BE 비교표"""
        sel = word_app.Selection
        rows = max(len(as_is), len(to_be)) + 2
        tbl = doc.Tables.Add(sel.Range, rows, 2)
        tbl.Borders.Enable = False

        # 1행: 전체 제목 병합
        tbl.Cell(1, 1).Merge(tbl.Cell(1, 2))
        c_title = tbl.Cell(1, 1)
        c_title.Range.Text = f"■ {title}"
        c_title.Range.Font.Name = "맑은 고딕"
        c_title.Range.Font.Size = 12
        c_title.Range.Font.Bold = True
        c_title.Range.Font.Color = COLOR_NAVY_BGR
        c_title.Range.ParagraphFormat.Alignment = 0
        c_title.Borders(WD_BORDER_BOTTOM).LineStyle = 1
        c_title.Borders(WD_BORDER_BOTTOM).LineWidth = 12
        c_title.Borders(WD_BORDER_BOTTOM).Color = COLOR_NAVY_BGR

        # 2행: 헤더 (AS-IS vs TO-BE)
        h1 = tbl.Cell(2, 1)
        h1.Range.Text = "AS-IS (현재 현안 / 한계)"
        h1.Range.Font.Name = "맑은 고딕"
        h1.Range.Font.Size = 10.5
        h1.Range.Font.Bold = True
        h1.Range.Font.Color = COLOR_TEXT_DARK_BGR
        h1.Shading.BackgroundPatternColor = hex_to_bgr("#E9ECEF")
        h1.Range.ParagraphFormat.Alignment = 1

        h2 = tbl.Cell(2, 2)
        h2.Range.Text = "TO-BE (개선 모델 / 기대 효과)"
        h2.Range.Font.Name = "맑은 고딕"
        h2.Range.Font.Size = 10.5
        h2.Range.Font.Bold = True
        h2.Range.Font.Color = COLOR_WHITE_BGR
        h2.Shading.BackgroundPatternColor = COLOR_TEAL_BGR
        h2.Range.ParagraphFormat.Alignment = 1

        # 데이터 행 채우기
        max_rows = max(len(as_is), len(to_be))
        for r in range(max_rows):
            row_idx = r + 3
            left_txt = strip_markdown_syntax(as_is[r]) if r < len(as_is) else "-"
            right_txt = strip_markdown_syntax(to_be[r]) if r < len(to_be) else "-"

            c_l = tbl.Cell(row_idx, 1)
            c_l.Range.Text = f"• {left_txt}"
            c_l.Range.Font.Name = "맑은 고딕"
            c_l.Range.Font.Size = 10
            c_l.Shading.BackgroundPatternColor = hex_to_bgr("#F8F9FA")
            c_l.Borders(WD_BORDER_BOTTOM).LineStyle = 1
            c_l.Borders(WD_BORDER_BOTTOM).LineWidth = 4
            c_l.Borders(WD_BORDER_BOTTOM).Color = COLOR_BORDER_GRAY_BGR

            c_r = tbl.Cell(row_idx, 2)
            c_r.Range.Text = f"✔ {right_txt}"
            c_r.Range.Font.Name = "맑은 고딕"
            c_r.Range.Font.Size = 10
            c_r.Range.Font.Bold = True
            c_r.Range.Font.Color = COLOR_NAVY_BGR
            c_r.Shading.BackgroundPatternColor = COLOR_LIGHT_BLUE_BGR
            c_r.Borders(WD_BORDER_BOTTOM).LineStyle = 1
            c_r.Borders(WD_BORDER_BOTTOM).LineWidth = 4
            c_r.Borders(WD_BORDER_BOTTOM).Color = COLOR_STEEL_BLUE_BGR

        sel.Start = tbl.Range.End
        sel.InsertParagraphAfter()
        sel.MoveDown(Unit=5, Count=1)
        return f"Word 문서에 AS-IS vs TO-BE 비교표('{title}')가 생성되었습니다."

    @staticmethod
    def insert_callout_box(word_app: Any, doc: Any, text: str, box_type: str = "info") -> str:
        """단일 셀 콜아웃 박스 (좌측 굵은 강조선 + 은은한 배경)"""
        sel = word_app.Selection
        tbl = doc.Tables.Add(sel.Range, 1, 1)
        cell = tbl.Cell(1, 1)

        type_styles = {
            "info": (COLOR_STEEL_BLUE_BGR, hex_to_bgr("#F4F7FA"), "[안내] "),
            "tip": (COLOR_TEAL_BGR, hex_to_bgr("#F2F8F6"), "※ [참고] "),
            "warning": (hex_to_bgr("#D97706"), hex_to_bgr("#FFFBEB"), "[유의사항] "),
            "quote": (hex_to_bgr("#475569"), hex_to_bgr("#F8FAFC"), "[핵심 요약] "),
        }
        border_color, bg_color, prefix = type_styles.get(box_type, type_styles["info"])

        clean_text = strip_markdown_syntax(text)
        cell.Range.Text = f"{prefix}{clean_text}"
        cell.Range.Font.Name = "맑은 고딕"
        cell.Range.Font.Size = 10.5
        cell.Range.Font.Color = COLOR_TEXT_DARK_BGR
        cell.Shading.BackgroundPatternColor = bg_color
        try:
            cell.Range.Font.NameFarEast = "맑은 고딕"
            cell.Range.Font.Color = hex_to_bgr("#1F2937")
            cell.Range.ParagraphFormat.FirstLineIndent = 0
            cell.Range.ParagraphFormat.LeftIndent = 0
            cell.Range.ParagraphFormat.SpaceBefore = 0
            cell.Range.ParagraphFormat.SpaceAfter = 0
            cell.Range.ParagraphFormat.LineSpacingRule = 5
            cell.Range.ParagraphFormat.LineSpacing = float(word_app.LinesToPoints(1.25))
            tbl.Rows(1).AllowBreakAcrossPages = False
            pfx_rng = doc.Range(cell.Range.Start, cell.Range.Start + len(prefix))
            pfx_rng.Font.Bold = -1
            pfx_rng.Font.Color = border_color
            tbl.TopPadding = 8
            tbl.BottomPadding = 8
            tbl.LeftPadding = 12
            tbl.RightPadding = 12
        except Exception:
            pass

        # 상하우 테두리 제거, 좌측만 24 (약 3pt) 실선
        for b in (-1, WD_BORDER_BOTTOM, -4):  # 상, 하, 우
            cell.Borders(b).LineStyle = 0
        cell.Borders(WD_BORDER_LEFT).LineStyle = 1  # 좌측
        cell.Borders(WD_BORDER_LEFT).LineWidth = 24  # 굵은 테두리
        cell.Borders(WD_BORDER_LEFT).Color = border_color

        # 셀 내부 좌측 여백 부여
        cell.LeftPadding = 12
        cell.RightPadding = 12
        cell.TopPadding = 8
        cell.BottomPadding = 8

        sel.Start = tbl.Range.End
        sel.InsertParagraphAfter()
        sel.MoveDown(Unit=5, Count=1)
        try:
            sel.Borders.Enable = False
            sel.Font.Bold = 0
            sel.Font.Color = hex_to_bgr("#1F2937")
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.LeftIndent = 0
        except Exception:
            pass
        return f"Word 문서에 {box_type.upper()} 요약 콜아웃 상자가 생성되었습니다."

    @staticmethod
    def insert_kpi_cards(word_app: Any, doc: Any, kpi_data: list[dict[str, str]]) -> str:
        """핵심 지표(KPI) 대시보드 카드"""
        sel = word_app.Selection
        n = len(kpi_data)
        tbl = doc.Tables.Add(sel.Range, 2, n)
        tbl.Borders.Enable = False

        for i, kpi in enumerate(kpi_data):
            label = kpi.get("label", f"지표 {i+1}")
            val = kpi.get("value", "0")
            change = kpi.get("change", "")
            is_up = not change.startswith("-")

            # 1행: 라벨 및 큰 수치
            c1 = tbl.Cell(1, i + 1)
            c1.Range.Text = f"{label}\n{val}"
            c1.Range.Font.Name = "맑은 고딕"
            # 첫 문단(라벨)
            p1 = c1.Range.Paragraphs(1)
            p1.Range.Font.Size = 9.5
            p1.Range.Font.Color = hex_to_bgr("#6B7280")
            p1.Alignment = 1
            # 둘째 문단(값)
            if c1.Range.Paragraphs.Count >= 2:
                p2 = c1.Range.Paragraphs(2)
                p2.Range.Font.Size = 16
                p2.Range.Font.Bold = True
                p2.Range.Font.Color = COLOR_NAVY_BGR
                p2.Alignment = 1

            c1.Shading.BackgroundPatternColor = COLOR_WHITE_BGR
            try:
                c1.VerticalAlignment = 1
                c1.TopPadding = 7
                c1.BottomPadding = 6
                c1.LeftPadding = 8
                c1.RightPadding = 8
                c1.Range.Font.NameFarEast = "맑은 고딕"
                c1.Range.ParagraphFormat.FirstLineIndent = 0
                tbl.Rows(1).AllowBreakAcrossPages = False
            except Exception:
                pass

            # 2행: 증감율 배지
            c2 = tbl.Cell(2, i + 1)
            badge_str = change.strip()
            is_num_gain = bool(re.match(r"^\+?\d", badge_str))
            is_neg = badge_str.startswith("-") or any(k in badge_str.upper() for k in ("위험", "초과", "FAIL"))
            badge_icon = ("▼ " if re.match(r"^-\d", badge_str) else "") if is_neg else ("▲ " if is_num_gain else "")
            c2.Range.Text = f"{badge_icon}{badge_str}" if badge_str else "-"
            c2.Range.Font.Name = "맑은 고딕"
            c2.Range.Font.Size = 9
            c2.Range.Font.Bold = True
            c2.Range.Font.Color = hex_to_bgr("#B91C1C") if is_neg else (hex_to_bgr("#047857") if (is_num_gain or any(k in badge_str.upper() for k in ("PASS", "적정", "완료", "OK"))) else hex_to_bgr("#334155"))
            c2.Range.ParagraphFormat.Alignment = 1
            c2.Shading.BackgroundPatternColor = COLOR_LIGHT_BG_BGR

            # 카드 외곽 테두리
            for b in (-1, -2, -3, -4):
                c1.Borders(b).LineStyle = 1
                c1.Borders(b).LineWidth = 4
                c1.Borders(b).Color = COLOR_BORDER_GRAY_BGR
                c2.Borders(b).LineStyle = 1
                c2.Borders(b).LineWidth = 4
                c2.Borders(b).Color = COLOR_BORDER_GRAY_BGR

        sel.Start = tbl.Range.End
        sel.InsertParagraphAfter()
        sel.MoveDown(Unit=5, Count=1)
        for k_col in range(1, n + 1):
            try:
                c2_k = tbl.Cell(2, k_col)
                c2_k.Shading.BackgroundPatternColor = hex_to_bgr("#F8FAFC")
                c2_k.VerticalAlignment = 1
                c2_k.TopPadding = 5
                c2_k.BottomPadding = 6
                c2_k.LeftPadding = 8
                c2_k.RightPadding = 8
                c2_k.Range.Font.NameFarEast = "맑은 고딕"
                c2_k.Range.ParagraphFormat.FirstLineIndent = 0
                tbl.Cell(1, k_col).Borders(WD_BORDER_TOP).LineWidth = 12
                tbl.Cell(1, k_col).Borders(WD_BORDER_TOP).Color = COLOR_NAVY_BGR
                c2_k.Borders(WD_BORDER_TOP).LineStyle = 0
            except Exception:
                pass
        return f"Word 문서에 {n}개 KPI 지표 카드가 생성되었습니다."

    @staticmethod
    def clean_and_structure_text(word_app: Any, doc: Any) -> str:
        """범피스 107개 궁극의정리 알고리즘: 7단계 개조식 위계 번호 감지 및 자동 들여쓰기 교정"""
        sel = word_app.Selection
        # 선택 영역이 지정되어 있으면 선택 영역 내 문단만, 커서 상태이면 전체 문서 문단 순회
        if sel.Start != sel.End:
            paragraphs = sel.Range.Paragraphs
        else:
            paragraphs = doc.Paragraphs

        rules = [
            # Level 1: 1., ■, ◆
            (re.compile(r"^(1\d*|[1-9])\.\s*(.*)"), 0, 0, True, 11.5, COLOR_NAVY_BGR),
            (re.compile(r"^([■◆])\s*(.*)"), 0, 0, True, 11.5, COLOR_NAVY_BGR),
            # Level 2: 가., ○, □
            (re.compile(r"^([가-하])\.\s*(.*)"), 15, 0, True, 10.5, COLOR_TEXT_DARK_BGR),
            (re.compile(r"^([○□])\s*(.*)"), 15, 0, True, 10.5, COLOR_TEXT_DARK_BGR),
            # Level 3: 1)
            (re.compile(r"^([1-9]\d*)\)\s*(.*)"), 30, 0, False, 10.0, COLOR_TEXT_DARK_BGR),
            # Level 4: 가)
            (re.compile(r"^([가-하])\)\s*(.*)"), 45, 0, False, 10.0, COLOR_TEXT_DARK_BGR),
            # Level 5: (1)
            (re.compile(r"^\(([1-9]\d*)\)\s*(.*)"), 60, 0, False, 10.0, COLOR_TEXT_DARK_BGR),
            # Level 6: (가)
            (re.compile(r"^\(([가-하])\)\s*(.*)"), 75, 0, False, 9.5, COLOR_TEXT_DARK_BGR),
            # Level 7: -, o, *, •
            (re.compile(r"^([o\-\*\u2022\u25cf\u25b8])\s*(.*)"), 90, 0, False, 9.5, hex_to_bgr("#4B5563")),
        ]

        count = 0
        for p in paragraphs:
            # 표 내부 셀 안의 문단은 건너뛰기
            try:
                if p.Range.Information(12):  # wdWithInTable = 12
                    continue
            except Exception:
                pass

            txt = p.Range.Text.strip()
            if not txt:
                continue
            for pattern, left_indent, hang_indent, bold, fsize, fcolor in rules:
                if pattern.match(txt):
                    try:
                        p.Format.CharacterUnitLeftIndent = 0
                        p.Format.CharacterUnitFirstLineIndent = 0
                    except Exception:
                        pass
                    p.Format.LeftIndent = left_indent
                    p.Format.FirstLineIndent = 0
                    p.Format.LineSpacingRule = 5
                    try:
                        p.Format.LineSpacing = float(word_app.LinesToPoints(1.22))
                    except Exception:
                        p.Format.LineSpacing = 14.6
                    p.Format.SpaceBefore = 8 if left_indent == 0 else (4 if left_indent == 15 else 1.5)
                    p.Format.KeepWithNext = -1 if bold else 0
                    p.Format.SpaceAfter = 3
                    p.Range.Font.Name = "맑은 고딕"
                    try:
                        p.Range.Font.NameFarEast = "맑은 고딕"
                    except Exception:
                        pass
                    p.Range.Font.Size = fsize
                    p.Range.Font.Bold = bold
                    p.Range.Font.Color = fcolor
                    count += 1
                    break

        return f"Word 문서 내 {count}개 문단의 개조식 위계 및 자간/들여쓰기를 자동 교정 완료했습니다."


# =============================================================================
# 2. Excel COM 컴포넌트 & 스마트 표 렌더러
# =============================================================================
class ExcelComponentRenderer:
    """Excel 시트 내 겉바속촉 자동 서식화 및 비즈니스 컴포넌트 렌더러"""

    @staticmethod
    def _parse_cell_pos(sheet: Any, start_cell: str | None = None) -> tuple[int, int]:
        """시작 셀 주소 파싱 (미지정 시 활성 셀 또는 사용 영역 다음 빈 행)"""
        if start_cell:
            c_str = start_cell.upper().strip()
            # 간단한 정규식으로 열 문자(A-Z)와 행 번호 분리
            m = re.match(r"^([A-Z]+)(\d+)$", c_str)
            if m:
                col_letters, row_str = m.groups()
                col_num = 0
                for ch in col_letters:
                    col_num = col_num * 26 + (ord(ch) - ord("A") + 1)
                return int(row_str), col_num
        
        # 기본 위치: 활성 셀이 A1이 아니면 활성 셀 위치, 비어있으면 A1
        try:
            act = sheet.Application.ActiveCell
            if act and (act.Row > 1 or act.Column > 1):
                return act.Row, act.Column
        except Exception:
            pass

        try:
            used = sheet.UsedRange
            if used and used.Rows.Count > 1:
                return used.Row + used.Rows.Count + 1, 1
        except Exception:
            pass

        return 1, 1

    @staticmethod
    def auto_format_table(excel_app: Any, sheet: Any, range_addr: str | None = None) -> str:
        """범피식 '겉바속촉' 실천 알고리즘 서식: 외곽 굵은선, 안쪽 가로선, 헤더 음영, 자동 서식화"""
        if range_addr:
            target_rng = sheet.Range(range_addr)
        else:
            target_rng = sheet.UsedRange

        total_rows = target_rng.Rows.Count
        total_cols = target_rng.Columns.Count
        if total_rows < 1 or total_cols < 1:
            return "유효한 표 영역을 찾을 수 없습니다."

        # 1. 1행 헤더 스타일링
        header = target_rng.Rows(1)
        header.Interior.Color = COLOR_NAVY_BGR
        header.Font.Name = "맑은 고딕"
        header.Font.Size = 10.5
        header.Font.Bold = True
        header.Font.Color = COLOR_WHITE_BGR
        header.HorizontalAlignment = -4108  # xlCenter
        header.VerticalAlignment = -4108    # xlCenter
        header.RowHeight = 24

        # 2. 데이터행 기본 정렬 및 패딩
        if total_rows > 1:
            data_rows = target_rng.Rows(f"2:{total_rows}")
            data_rows.Font.Name = "맑은 고딕"
            data_rows.Font.Size = 9.5
            data_rows.VerticalAlignment = -4108
            data_rows.RowHeight = 19

            for c in range(1, total_cols + 1):
                col_rng = target_rng.Columns(c)
                sample_val = target_rng.Cells(2, c).Value
                if isinstance(sample_val, (int, float)):
                    col_rng.NumberFormatLocal = "#,##0"
                    col_rng.HorizontalAlignment = -4152  # xlRight
                elif isinstance(sample_val, str) and sample_val.replace(",", "").replace(".", "").isdigit():
                    col_rng.NumberFormatLocal = "#,##0"
                    col_rng.HorizontalAlignment = -4152

        # 3. 겉바속촉 테두리 (Border Constants: 7=Left, 8=Top, 9=Bottom, 10=Right, 11=InsideVert, 12=InsideHoriz)
        try:
            for edge in (7, 8, 9, 10):
                target_rng.Borders(edge).LineStyle = 1
                target_rng.Borders(edge).Weight = -4138
                target_rng.Borders(edge).Color = COLOR_NAVY_BGR

            target_rng.Borders(12).LineStyle = 1
            target_rng.Borders(12).Weight = 2
            target_rng.Borders(12).Color = COLOR_BORDER_GRAY_BGR
            target_rng.Borders(11).LineStyle = -4142

            header.Borders(9).LineStyle = -4119  # xlDouble
            header.Borders(9).Color = COLOR_NAVY_BGR
        except Exception as e:
            logger.warning(f"Excel 테두리 설정 중 건너뜀: {e}")

        target_rng.Columns.AutoFit()
        return f"Excel '{sheet.Name}' 시트의 {target_rng.Address} 영역에 겉바속촉 프리미엄 표 서식이 적용되었습니다."

    @staticmethod
    def insert_smart_table(
        excel_app: Any,
        sheet: Any,
        headers: list[str],
        rows: list[list[Any]],
        start_cell: str | None = None,
    ) -> str:
        """KISA 1310 규격 스마트 비즈니스 표: 겉바속촉(좌우오픈, 헤더이중선) + 컬럼 지능형 정렬"""
        start_r, start_c = ExcelComponentRenderer._parse_cell_pos(sheet, start_cell)
        num_rows = len(rows) + 1
        num_cols = len(headers)

        target_rng = sheet.Range(
            sheet.Cells(start_r, start_c),
            sheet.Cells(start_r + num_rows - 1, start_c + num_cols - 1)
        )

        # 데이터 일괄 주입
        all_matrix = [headers] + [[str(val) if val is not None else "" for val in r] for r in rows]
        try:
            target_rng.Value = all_matrix
        except Exception:
            for r_i, r_data in enumerate(all_matrix):
                for c_i, v in enumerate(r_data):
                    target_rng.Cells(r_i + 1, c_i + 1).Value = v

        # 1행 헤더 스타일
        header = target_rng.Rows(1)
        header.Interior.Color = COLOR_NAVY_BGR
        header.Font.Name = "맑은 고딕"
        header.Font.Size = 10.5
        header.Font.Bold = True
        header.Font.Color = COLOR_WHITE_BGR
        header.HorizontalAlignment = -4108  # xlCenter
        header.VerticalAlignment = -4108
        header.RowHeight = 25

        # 겉바속촉 KISA 1310 테두리 설정
        try:
            target_rng.Borders(7).LineStyle = -4142  # 좌측 없음
            target_rng.Borders(10).LineStyle = -4142 # 우측 없음
            target_rng.Borders(8).LineStyle = 1      # 상단 굵은 실선
            target_rng.Borders(8).Weight = -4138     # xlMedium
            target_rng.Borders(8).Color = COLOR_NAVY_BGR
            target_rng.Borders(9).LineStyle = 1      # 하단 굵은 실선
            target_rng.Borders(9).Weight = -4138     # xlMedium
            target_rng.Borders(9).Color = COLOR_NAVY_BGR
            target_rng.Borders(12).LineStyle = 1     # 내부 가로 얇은 실선
            target_rng.Borders(12).Weight = 2        # xlThin
            target_rng.Borders(12).Color = COLOR_BORDER_GRAY_BGR
            target_rng.Borders(11).LineStyle = -4142 # 내부 세로선 없음

            header.Borders(9).LineStyle = -4119      # 헤더 하단 xlDouble 이중선
            header.Borders(9).Color = COLOR_NAVY_BGR
        except Exception as e:
            logger.warning(f"테두리 설정 오류 건너뜀: {e}")

        # 본문 행 스타일 및 HWP형 지능형 컬럼 정렬
        if len(rows) > 0:
            data_rows = target_rng.Rows(f"2:{num_rows}")
            data_rows.Font.Name = "맑은 고딕"
            data_rows.Font.Size = 9.5
            data_rows.VerticalAlignment = -4108
            data_rows.RowHeight = 20

            for c_idx in range(1, num_cols + 1):
                col_rng = target_rng.Columns(c_idx)
                col_data = col_rng.Offset(1, 0).Resize(num_rows - 1, 1)
                hdr_text = str(headers[c_idx - 1]).strip().upper()
                c_vals = [str(rows[r_i][c_idx - 1]).strip() for r_i in range(len(rows)) if c_idx - 1 < len(rows[r_i])]

                is_index = any(k in hdr_text for k in ["NO", "순번", "번호", "단계", "일자", "구분", "코드", "상태"]) or all(v.isdigit() and len(v) <= 3 for v in c_vals if v)
                has_num = any(sym in "".join(c_vals) for sym in [",", "%", "원", "$", "₩"]) or (sum(1 for v in c_vals if re.match(r'^-?\d+(?:\.\d+)?$', re.sub(r'[,%\s원\$₩]', '', v))) >= len(c_vals) * 0.7 if c_vals else False)

                if is_index:
                    col_data.HorizontalAlignment = -4108  # xlCenter
                elif has_num:
                    col_data.HorizontalAlignment = -4152  # xlRight
                    if any("." in v for v in c_vals):
                        col_data.NumberFormatLocal = "#,##0.00"
                    elif any("%" in v for v in c_vals):
                        col_data.NumberFormatLocal = "0.0%"
                    else:
                        col_data.NumberFormatLocal = "#,##0"
                elif max((len(v) for v in c_vals), default=0) <= 8:
                    col_data.HorizontalAlignment = -4108
                else:
                    col_data.HorizontalAlignment = -4131  # xlLeft

        try:
            target_rng.Columns.AutoFit()
        except Exception:
            pass

        return f"Excel '{sheet.Name}' 시트에 KISA 1310 겉바속촉 스마트 비즈니스 표({num_rows}행 {num_cols}열)를 성공적으로 생성했습니다. (위치: {target_rng.Address})"

    @staticmethod
    def insert_timeline(
        excel_app: Any,
        sheet: Any,
        title: str,
        milestones: list[dict[str, str]],
        start_cell: str | None = None,
    ) -> str:
        """간트차트형 마일스톤 타임라인 표 (좌측 마일스톤 메타 + 우측 M01~M12 월별 간트 바 음영)"""
        start_r, start_c = ExcelComponentRenderer._parse_cell_pos(sheet, start_cell)
        num_milestones = len(milestones)
        total_rows = num_milestones + 2  # 타이틀(1), 헤더(1), 데이터(N)
        total_cols = 4 + 12  # [단계, 과업명, 일정, 상태] + [M01 ~ M12]

        full_rng = sheet.Range(
            sheet.Cells(start_r, start_c),
            sheet.Cells(start_r + total_rows - 1, start_c + total_cols - 1)
        )

        # 1행: 타이틀 병합 행
        title_rng = sheet.Range(sheet.Cells(start_r, start_c), sheet.Cells(start_r, start_c + total_cols - 1))
        title_rng.Merge()
        title_rng.Value = f"■ 추진 일정 로드맵: {title}"
        title_rng.Font.Name = "맑은 고딕"
        title_rng.Font.Size = 11.5
        title_rng.Font.Bold = True
        title_rng.Font.Color = COLOR_NAVY_BGR
        title_rng.Interior.Color = COLOR_LIGHT_BG_BGR
        title_rng.VerticalAlignment = -4108
        title_rng.RowHeight = 26

        # 2행: 헤더 행
        headers = ["단계", "주요 과업 / 마일스톤", "추진 기간", "상태"] + [f"M{m:02d}" for m in range(1, 13)]
        hdr_rng = sheet.Range(sheet.Cells(start_r + 1, start_c), sheet.Cells(start_r + 1, start_c + total_cols - 1))
        hdr_rng.Value = headers
        hdr_rng.Interior.Color = COLOR_NAVY_BGR
        hdr_rng.Font.Name = "맑은 고딕"
        hdr_rng.Font.Size = 9.5
        hdr_rng.Font.Bold = True
        hdr_rng.Font.Color = COLOR_WHITE_BGR
        hdr_rng.HorizontalAlignment = -4108
        hdr_rng.VerticalAlignment = -4108
        hdr_rng.RowHeight = 22

        # 3행~: 마일스톤 데이터 및 간트 바 음영
        for idx, m_data in enumerate(milestones):
            curr_r = start_r + 2 + idx
            phase = m_data.get("phase", f"{idx+1}단계")
            task = m_data.get("task", "")
            period = m_data.get("period", "")
            status = m_data.get("status", "진행중")

            # 4개 기본 컬럼 주입
            sheet.Cells(curr_r, start_c).Value = phase
            sheet.Cells(curr_r, start_c + 1).Value = task
            sheet.Cells(curr_r, start_c + 2).Value = period
            sheet.Cells(curr_r, start_c + 3).Value = status

            row_rng = sheet.Range(sheet.Cells(curr_r, start_c), sheet.Cells(curr_r, start_c + 3))
            row_rng.Font.Name = "맑은 고딕"
            row_rng.Font.Size = 9.0
            row_rng.VerticalAlignment = -4108
            sheet.Cells(curr_r, start_c).HorizontalAlignment = -4108
            sheet.Cells(curr_r, start_c + 2).HorizontalAlignment = -4108
            sheet.Cells(curr_r, start_c + 3).HorizontalAlignment = -4108
            sheet.Rows(curr_r).RowHeight = 20

            # 기간 문자열에서 월 번호(1~12) 추출하여 간트 바 채우기
            # 예: "2026.03~2026.06", "3월~5월", "01-04", "M02~M05"
            months = [int(n) for n in re.findall(r'(?:M|월|)(\d{1,2})', period) if 1 <= int(n) <= 12]
            if len(months) >= 2:
                sm, em = min(months[0], months[1]), max(months[0], months[1])
            elif len(months) == 1:
                sm, em = months[0], months[0]
            else:
                sm, em = (idx % 12) + 1, min((idx % 12) + 2, 12)

            bar_start_c = start_c + 4 + (sm - 1)
            bar_end_c = start_c + 4 + (em - 1)
            bar_rng = sheet.Range(sheet.Cells(curr_r, bar_start_c), sheet.Cells(curr_r, bar_end_c))
            bar_rng.Interior.Color = COLOR_EMERALD_BGR
            try:
                for b_edge in (7, 8, 9, 10):
                    bar_rng.Borders(b_edge).LineStyle = 1
                    bar_rng.Borders(b_edge).Weight = 2
                    bar_rng.Borders(b_edge).Color = COLOR_TEAL_BGR
            except Exception:
                pass

        # 외곽 및 격자 서식
        try:
            full_rng.Borders(7).LineStyle = 1
            full_rng.Borders(10).LineStyle = 1
            full_rng.Borders(8).Weight = -4138
            full_rng.Borders(9).Weight = -4138
            full_rng.Borders(12).LineStyle = 1
            full_rng.Borders(12).Weight = 2
            full_rng.Borders(12).Color = COLOR_BORDER_GRAY_BGR
            # 월 컬럼 너비 컴팩트 조정
            for m_i in range(12):
                sheet.Columns(start_c + 4 + m_i).ColumnWidth = 4.2
        except Exception:
            pass

        return f"Excel '{sheet.Name}' 시트에 간트차트형 타임라인 로드맵 표({num_milestones}개 과제, M01~M12)를 생성했습니다. (위치: {full_rng.Address})"

    @staticmethod
    def insert_comparison(
        excel_app: Any,
        sheet: Any,
        title: str,
        as_is: list[str],
        to_be: list[str],
        start_cell: str | None = None,
    ) -> str:
        """AS-IS vs TO-BE 대칭 비교 카드 표 (좌측 연회색 현행 vs 우측 에메랄드 그린 목표)"""
        start_r, start_c = ExcelComponentRenderer._parse_cell_pos(sheet, start_cell)
        max_rows = max(len(as_is), len(to_be))
        total_rows = max_rows + 2

        comp_rng = sheet.Range(sheet.Cells(start_r, start_c), sheet.Cells(start_r + total_rows - 1, start_c + 3))

        # 1행: 비교 대주제 타이틀
        title_rng = sheet.Range(sheet.Cells(start_r, start_c), sheet.Cells(start_r, start_c + 3))
        title_rng.Merge()
        title_rng.Value = f"■ [비교 분석] {title}"
        title_rng.Font.Name = "맑은 고딕"
        title_rng.Font.Size = 11.5
        title_rng.Font.Bold = True
        title_rng.Font.Color = COLOR_NAVY_BGR
        title_rng.Interior.Color = COLOR_LIGHT_BG_BGR
        title_rng.VerticalAlignment = -4108
        title_rng.RowHeight = 26

        # 2행: 헤더 (AS-IS vs TO-BE)
        asis_hdr = sheet.Range(sheet.Cells(start_r + 1, start_c), sheet.Cells(start_r + 1, start_c + 1))
        asis_hdr.Merge()
        asis_hdr.Value = "AS-IS (현재 상태 및 한계)"
        asis_hdr.Font.Bold = True
        asis_hdr.Font.Color = COLOR_TEXT_DARK_BGR
        asis_hdr.Interior.Color = COLOR_LIGHT_BG_BGR
        asis_hdr.HorizontalAlignment = -4108
        asis_hdr.RowHeight = 24

        tobe_hdr = sheet.Range(sheet.Cells(start_r + 1, start_c + 2), sheet.Cells(start_r + 1, start_c + 3))
        tobe_hdr.Merge()
        tobe_hdr.Value = "TO-BE (개선 목표 및 기대효과)"
        tobe_hdr.Font.Bold = True
        tobe_hdr.Font.Color = COLOR_TEAL_BGR
        tobe_hdr.Interior.Color = COLOR_LIGHT_GREEN_BGR
        tobe_hdr.HorizontalAlignment = -4108

        # 3행~: 비교 데이터 대칭 입력
        for i in range(max_rows):
            curr_r = start_r + 2 + i
            as_text = f"• {as_is[i]}" if i < len(as_is) else ""
            to_text = f"✔ {to_be[i]}" if i < len(to_be) else ""

            left_rng = sheet.Range(sheet.Cells(curr_r, start_c), sheet.Cells(curr_r, start_c + 1))
            left_rng.Merge()
            left_rng.Value = as_text
            left_rng.Font.Name = "맑은 고딕"
            left_rng.Font.Size = 9.5
            left_rng.VerticalAlignment = -4108
            left_rng.WrapText = True

            right_rng = sheet.Range(sheet.Cells(curr_r, start_c + 2), sheet.Cells(curr_r, start_c + 3))
            right_rng.Merge()
            right_rng.Value = to_text
            right_rng.Font.Name = "맑은 고딕"
            right_rng.Font.Size = 9.5
            right_rng.Font.Bold = True
            right_rng.VerticalAlignment = -4108
            right_rng.WrapText = True

            sheet.Rows(curr_r).RowHeight = 24

        try:
            # 외곽 테두리 및 중앙 구분선
            comp_rng.Borders(7).LineStyle = 1
            comp_rng.Borders(10).LineStyle = 1
            comp_rng.Borders(8).Weight = -4138
            comp_rng.Borders(9).Weight = -4138
            comp_rng.Borders(12).LineStyle = 1
            comp_rng.Borders(12).Weight = 2
            comp_rng.Borders(12).Color = COLOR_BORDER_GRAY_BGR
            # 중앙 분할선 (Col 2와 3 사이)
            mid_col = sheet.Range(sheet.Cells(start_r + 1, start_c + 1), sheet.Cells(start_r + total_rows - 1, start_c + 1))
            mid_col.Borders(10).LineStyle = 1
            mid_col.Borders(10).Weight = -4138
            mid_col.Borders(10).Color = COLOR_NAVY_BGR
        except Exception:
            pass

        return f"Excel '{sheet.Name}' 시트에 AS-IS vs TO-BE 대칭 비교 카드 표를 생성했습니다. (위치: {comp_rng.Address})"

    @staticmethod
    def insert_callout_box(
        excel_app: Any,
        sheet: Any,
        text: str,
        box_type: str = "info",
        start_cell: str | None = None,
    ) -> str:
        """단일/다중 행 콜아웃 박스 (좌측 굵은 강조선 + 파스텔톤 배경음영)"""
        start_r, start_c = ExcelComponentRenderer._parse_cell_pos(sheet, start_cell)
        
        type_cfg = {
            "info": (COLOR_NAVY_BGR, COLOR_LIGHT_BLUE_BGR, "ℹ 안내"),
            "tip": (COLOR_EMERALD_BGR, COLOR_LIGHT_GREEN_BGR, "💡 팁"),
            "warning": (hex_to_bgr("#D97706"), COLOR_LIGHT_AMBER_BGR, "⚠ 주의"),
            "quote": (COLOR_STEEL_BLUE_BGR, COLOR_LIGHT_BG_BGR, "💬 핵심 요약"),
        }
        accent_color, bg_color, badge = type_cfg.get(box_type, type_cfg["info"])

        box_rng = sheet.Range(sheet.Cells(start_r, start_c), sheet.Cells(start_r + 1, start_c + 5))
        box_rng.Merge()
        box_rng.Value = f"{badge}: {text}"
        box_rng.Font.Name = "맑은 고딕"
        box_rng.Font.Size = 10.0
        box_rng.Font.Color = COLOR_TEXT_DARK_BGR
        box_rng.Interior.Color = bg_color
        box_rng.VerticalAlignment = -4108
        box_rng.WrapText = True
        sheet.Rows(start_r).RowHeight = 22
        sheet.Rows(start_r + 1).RowHeight = 22

        try:
            # 좌측만 굵은 선(-4138 xlMedium 또는 4 xlThick)
            box_rng.Borders(7).LineStyle = 1
            box_rng.Borders(7).Weight = -4138
            box_rng.Borders(7).Color = accent_color
            box_rng.Borders(8).LineStyle = -4142
            box_rng.Borders(9).LineStyle = -4142
            box_rng.Borders(10).LineStyle = -4142
        except Exception:
            pass

        return f"Excel '{sheet.Name}' 시트에 '{box_type}' 콜아웃 박스를 생성했습니다. (위치: {box_rng.Address})"

    @staticmethod
    def insert_approval_block(
        excel_app: Any,
        sheet: Any,
        approvers: list[str] | None = None,
        start_cell: str = "G1",
    ) -> str:
        """KISA 공문서 및 기업 표준 결재선 블록 (2행 N열 격자 및 서명 공란)"""
        start_r, start_c = ExcelComponentRenderer._parse_cell_pos(sheet, start_cell)
        if not approvers:
            approvers = ["담당", "검토", "심사", "승인"]

        num_appr = len(approvers)
        total_cols = num_appr + 1

        block_rng = sheet.Range(sheet.Cells(start_r, start_c), sheet.Cells(start_r + 1, start_c + total_cols - 1))

        # 1. '결재' 세로 병합 셀
        sign_label = sheet.Range(sheet.Cells(start_r, start_c), sheet.Cells(start_r + 1, start_c))
        sign_label.Merge()
        sign_label.Value = "결\n재"
        sign_label.Font.Name = "맑은 고딕"
        sign_label.Font.Size = 9.5
        sign_label.Font.Bold = True
        sign_label.Interior.Color = COLOR_LIGHT_BG_BGR
        sign_label.HorizontalAlignment = -4108
        sign_label.VerticalAlignment = -4108
        sheet.Columns(start_c).ColumnWidth = 3.5

        # 2. 직책 헤더 및 서명란
        for idx, role in enumerate(approvers):
            col_pos = start_c + 1 + idx
            # 직책명 헤더
            hdr_cell = sheet.Cells(start_r, col_pos)
            hdr_cell.Value = role
            hdr_cell.Font.Name = "맑은 고딕"
            hdr_cell.Font.Size = 9.0
            hdr_cell.Interior.Color = COLOR_LIGHT_BG_BGR
            hdr_cell.HorizontalAlignment = -4108
            hdr_cell.VerticalAlignment = -4108

            # 서명란 빈칸 (높이 확보)
            sign_cell = sheet.Cells(start_r + 1, col_pos)
            sign_cell.Value = ""
            sheet.Columns(col_pos).ColumnWidth = 8.5

        sheet.Rows(start_r).RowHeight = 18
        sheet.Rows(start_r + 1).RowHeight = 42

        # 3. 전체 격자 테두리
        try:
            for b_idx in (7, 8, 9, 10, 11, 12):
                block_rng.Borders(b_idx).LineStyle = 1
                block_rng.Borders(b_idx).Weight = 2
                block_rng.Borders(b_idx).Color = COLOR_NAVY_BGR
            # 외곽 굵은선
            for edge in (7, 8, 9, 10):
                block_rng.Borders(edge).Weight = -4138
        except Exception:
            pass

        return f"Excel '{sheet.Name}' 시트에 KISA 표준 결재선({num_appr}명)을 생성했습니다. (위치: {block_rng.Address})"

    @staticmethod
    def insert_kpi_cards(
        excel_app: Any,
        sheet: Any,
        kpi_data: list[dict[str, str]],
        start_cell: str | None = None,
    ) -> str:
        """KPI 요약 대시보드 카드 블록 (상단 라벨, 중앙 대형 지표, 하단 증감율 배지)"""
        start_r, start_c = ExcelComponentRenderer._parse_cell_pos(sheet, start_cell)
        num_cards = len(kpi_data)

        for c_i, card in enumerate(kpi_data):
            col_offset = start_c + c_i * 3  # 각 카드당 2열 차지 + 1열 간격
            card_rng = sheet.Range(sheet.Cells(start_r, col_offset), sheet.Cells(start_r + 2, col_offset + 1))

            label = card.get("label", "핵심 지표")
            val = card.get("value", "0")
            change = card.get("change", "")

            # 1행: 라벨
            lbl_rng = sheet.Range(sheet.Cells(start_r, col_offset), sheet.Cells(start_r, col_offset + 1))
            lbl_rng.Merge()
            lbl_rng.Value = label
            lbl_rng.Font.Name = "맑은 고딕"
            lbl_rng.Font.Size = 9.0
            lbl_rng.Font.Color = COLOR_TEXT_DARK_BGR
            lbl_rng.Interior.Color = COLOR_LIGHT_BG_BGR
            lbl_rng.HorizontalAlignment = -4108

            # 2행: 핵심 값
            val_rng = sheet.Range(sheet.Cells(start_r + 1, col_offset), sheet.Cells(start_r + 1, col_offset + 1))
            val_rng.Merge()
            val_rng.Value = val
            val_rng.Font.Name = "맑은 고딕"
            val_rng.Font.Size = 15.0
            val_rng.Font.Bold = True
            val_rng.Font.Color = COLOR_NAVY_BGR
            val_rng.HorizontalAlignment = -4108

            # 3행: 변동률
            chg_rng = sheet.Range(sheet.Cells(start_r + 2, col_offset), sheet.Cells(start_r + 2, col_offset + 1))
            chg_rng.Merge()
            chg_rng.Value = change
            chg_rng.Font.Name = "맑은 고딕"
            chg_rng.Font.Size = 9.5
            chg_rng.Font.Bold = True
            chg_rng.Font.Color = COLOR_TEAL_BGR if "-" not in change else hex_to_bgr("#DC2626")
            chg_rng.HorizontalAlignment = -4108

            sheet.Columns(col_offset).ColumnWidth = 8.0
            sheet.Columns(col_offset + 1).ColumnWidth = 8.0
            if c_i < num_cards - 1:
                sheet.Columns(col_offset + 2).ColumnWidth = 2.0  # 카드 사이 간격

            try:
                for b_e in (7, 8, 9, 10):
                    card_rng.Borders(b_e).LineStyle = 1
                    card_rng.Borders(b_e).Weight = 2
                    card_rng.Borders(b_e).Color = COLOR_BORDER_GRAY_BGR
            except Exception:
                pass

        sheet.Rows(start_r).RowHeight = 18
        sheet.Rows(start_r + 1).RowHeight = 28
        sheet.Rows(start_r + 2).RowHeight = 18

        return f"Excel '{sheet.Name}' 시트에 {num_cards}개 KPI 메트릭 카드를 생성했습니다."


# =============================================================================
# 3. PowerPoint COM 컴포넌트 렌더러
# =============================================================================
class PptComponentRenderer:
    """PowerPoint 슬라이드 내 반응형 도형/표 기반 비즈니스 컴포넌트 렌더러"""

    @staticmethod
    def insert_process_flow(ppt_app: Any, pres: Any, steps: list[dict[str, str]]) -> str:
        """PPT 슬라이드에 세련된 카드 + 쉐이프 기반 프로세스 생성"""
        slide = pres.Slides(ppt_app.ActiveWindow.View.Slide.SlideIndex)
        n = len(steps)
        slide_width = pres.PageSetup.SlideWidth
        slide_height = pres.PageSetup.SlideHeight

        margin_x = 40
        spacing = 15
        avail_width = slide_width - (margin_x * 2) - ((n - 1) * spacing)
        card_w = avail_width / n
        card_h = 220
        card_y = (slide_height - card_h) / 2

        for i, step in enumerate(steps):
            x = margin_x + i * (card_w + spacing)
            title = step.get("title", f"STEP {i+1}")
            desc = step.get("desc", "")

            # 둥근 사각형 카드 (msoShapeRoundedRectangle = 5)
            shp = slide.Shapes.AddShape(5, x, card_y, card_w, card_h)
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = COLOR_LIGHT_BG_BGR
            shp.Line.ForeColor.RGB = COLOR_STEEL_BLUE_BGR
            shp.Line.Weight = 1.5

            tf = shp.TextFrame
            tf.MarginLeft = 12
            tf.MarginRight = 12
            tf.MarginTop = 16
            tf.MarginBottom = 16

            tr = tf.TextRange
            tr.Text = f"STEP {i+1}\r{title}\r\r{desc}"  # PPT 문단 구분은 \r, \n은 문단을 나누지 않음
            tr.Font.Name = "맑은 고딕"

            # STEP 헤더
            p1 = tr.Paragraphs(1)
            p1.Font.Size = 11
            p1.Font.Bold = True
            p1.Font.Color.RGB = COLOR_TEAL_BGR

            # 제목
            p2 = tr.Paragraphs(2)
            p2.Font.Size = 14
            p2.Font.Bold = True
            p2.Font.Color.RGB = COLOR_NAVY_BGR

            # 설명 (3번째 문단은 빈 줄)
            if desc:
                p4 = tr.Paragraphs(4)
                p4.Font.Size = 10
                p4.Font.Bold = False
                p4.Font.Color.RGB = COLOR_TEXT_DARK_BGR

        return f"PPT 슬라이드에 {n}개 단계 프로세스 카드가 추가되었습니다."

    @staticmethod
    def insert_kpi_cards(ppt_app: Any, pres: Any, kpi_data: list[dict[str, str]]) -> str:
        """PPT 슬라이드 내 KPI 대시보드 카드 배치"""
        slide = pres.Slides(ppt_app.ActiveWindow.View.Slide.SlideIndex)
        n = len(kpi_data)
        slide_width = pres.PageSetup.SlideWidth
        margin_x = 50
        spacing = 20
        avail_width = slide_width - (margin_x * 2) - ((n - 1) * spacing)
        card_w = avail_width / n
        card_h = 130
        card_y = 180

        for i, kpi in enumerate(kpi_data):
            x = margin_x + i * (card_w + spacing)
            label = kpi.get("label", "지표")
            val = kpi.get("value", "0")
            change = kpi.get("change", "")
            is_up = not change.startswith("-")

            shp = slide.Shapes.AddShape(5, x, card_y, card_w, card_h)
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = COLOR_LIGHT_BLUE_BGR
            shp.Line.ForeColor.RGB = COLOR_STEEL_BLUE_BGR
            shp.Line.Weight = 1.5

            tr = shp.TextFrame.TextRange
            badge = f" (▲ {change})" if is_up and change else f" (▼ {change})" if change else ""
            tr.Text = f"{label}\r{val}\r{badge}"
            tr.Font.Name = "맑은 고딕"

            p1 = tr.Paragraphs(1)
            p1.Font.Size = 11
            p1.Font.Color.RGB = hex_to_bgr("#4B5563")

            p2 = tr.Paragraphs(2)
            p2.Font.Size = 20
            p2.Font.Bold = True
            p2.Font.Color.RGB = COLOR_NAVY_BGR

            if badge:
                p3 = tr.Paragraphs(3)
                p3.Font.Size = 10
                p3.Font.Bold = True
                p3.Font.Color.RGB = hex_to_bgr("#059669") if is_up else hex_to_bgr("#DC2626")

        return f"PPT 슬라이드에 {n}개 KPI 대시보드 카드가 추가되었습니다."


# =============================================================================
# 4. HWP COM 컴포넌트 렌더러
# =============================================================================
class HwpComponentRenderer:
    """한컴 HWP 내 Action 기반 비즈니스 컴포넌트 생성"""

    @staticmethod
    def insert_callout_box(hwp: Any, text: str, box_type: str = "info") -> str:
        """HWP 단일 셀 표 기반 콜아웃 박스"""
        ensure_com()
        prefix_map = {
            "info": "ℹ️ [안내] ",
            "tip": "💡 [TIP] ",
            "warning": "⚠️ [주의] ",
            "quote": "📌 [요약] ",
        }
        prefix = prefix_map.get(box_type, "ℹ️ [안내] ")
        content = f"{prefix}{text}"

        # HAction TableCreate
        act = hwp.CreateAction("TableCreate")
        set_ = act.CreateSet()
        act.GetDefault(set_)
        set_.SetItem("Rows", 1)
        set_.SetItem("Cols", 1)
        set_.SetItem("WidthType", 2)  # 단에 맞춤
        set_.SetItem("HeightType", 0)  # 자동
        act.Execute(set_)

        # 셀 내부 텍스트 삽입
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("ParagraphShape")
        hwp.InsertText(content)
        hwp.HAction.Run("CloseEx")
        return f"HWP 문서에 {box_type.upper()} 요약 박스가 생성되었습니다."


# =============================================================================
# MCPServer 도구 등록 함수
# =============================================================================
def register_component_tools(mcp: MCPServer) -> None:
    """7대 비즈니스 컴포넌트 및 정제 도구를 MCPServer 서버에 등록"""

    @mcp.tool()
    def office_insert_hierarchy(
        title: str,
        levels: list[str | dict[str, Any]],
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [비즈니스 컴포넌트 1] 조직도 및 계층 구조도 생성 도구.
        투명 격자 표에 카드와 연결선을 그려 호출 한 번으로 완성합니다. 3단 이상도 이 호출 하나로 만들고, 행을 덧붙이지 마세요.
        - title: 최상위 조직/개념명 (예: '대표이사')
        - levels: 하위 가지 목록(1~7개). 문자열 또는 {'name': ..., 'children': [...]} 형식이며 children도 같은 형식으로 중첩할 수 있습니다.
          예: [{'name': '경영기획본부', 'children': ['전략기획팀', '재무관리팀']}, '사업운영본부']
          가지 아래 노드는 세로로 쌓이고, 깊이에 따라 카드 색이 달라집니다.
        - target_app: 대상 오피스 프로그램 ('word', 'excel', 'ppt', 'hwp')
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, "Insert Hierarchy"):
                return WordComponentRenderer.insert_hierarchy(word_app, doc, title, levels)

        return f"현재 {target_app}용 조직도 생성을 준비 중입니다. Word에서 먼저 확인해 보세요."

    @mcp.tool()
    def office_insert_process_flow(
        steps: list[dict[str, str]],
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [비즈니스 컴포넌트 2] 단계별 프로세스 및 업무 흐름도 생성 도구.
        N단계 카드와 '➜' 화살표가 이어지는 모던 흐름도를 자동 생성합니다.
        - steps: 단계 목록 [{'title': '착수 및 기획', 'desc': '요구사항 수집 및 일정 수립'}, ...]
        - target_app: 대상 오피스 프로그램 ('word', 'ppt')
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, "Insert Process"):
                return WordComponentRenderer.insert_process_flow(word_app, doc, steps)
        elif target_app == "ppt":
            ppt_app, pres = com_manager.get_ppt()
            if not pres:
                return "열려 있는 PowerPoint 프레젠테이션을 찾을 수 없습니다."
            return PptComponentRenderer.insert_process_flow(ppt_app, pres, steps)

        return f"{target_app}용 프로세스 흐름도는 Word 또는 PPT를 권장합니다."

    @mcp.tool()
    def office_insert_timeline(
        title: str,
        milestones: list[dict[str, str]],
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [비즈니스 컴포넌트 3] 추진 일정 및 타임라인 로드맵 생성 도구.
        마일스톤, 기간, 핵심 과제, 진행 현황 배지가 포함된 간트형 표를 자동 생성합니다.
        - title: 일정표 제목 (예: '2026 하반기 신사업 추진 로드맵')
        - milestones: [{'phase': '1단계: 기획', 'period': '2026.01~03', 'task': '기획 수립', 'status': '완료'}, ...]
        - target_app: 대상 오피스 ('word')
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, "Insert Timeline"):
                return WordComponentRenderer.insert_timeline(word_app, doc, title, milestones)
        elif target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            return ExcelComponentRenderer.insert_timeline(excel_app, wb.ActiveSheet, title, milestones)
        return "현재 추진일정 로드맵은 Word 및 Excel을 지원합니다."

    @mcp.tool()
    def office_insert_smart_table(
        headers: list[str],
        rows: list[list[Any]],
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [비즈니스 컴포넌트 7] 겉바속촉 프리미엄 비즈니스 스마트 표 생성 도구.
        상하 굵은 실선, 헤더 이중선, 오픈 테이블(세로선 배제), 숫자 우측 정렬이 적용된 컨설팅 리포트형 표를 생성합니다.
        - headers: 헤더 열 제목 리스트 (예: ['항목', '단가', '수량', '합계'])
        - rows: 데이터 행 리스트 (예: [['서버 인프라', '1,000,000', '2', '2,000,000'], ...])
        - target_app: 대상 오피스 ('word')
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, "Insert Smart Table"):
                return WordComponentRenderer.insert_smart_table(word_app, doc, headers, rows)
        elif target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            return ExcelComponentRenderer.insert_smart_table(excel_app, wb.ActiveSheet, headers, rows)
        return "현재 스마트표 생성은 Word 및 Excel을 지원합니다."

    @mcp.tool()
    def office_insert_comparison(
        title: str,
        as_is: list[str],
        to_be: list[str],
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [비즈니스 컴포넌트 4] AS-IS vs TO-BE 비교 및 대안 대비표 생성 도구.
        현재 상태의 한계(좌측)와 개선 방안(우측 강조)을 대칭 카드로 시각화합니다.
        - title: 비교 주제 (예: '차세대 시스템 도입 전후 비교')
        - as_is: 현재 상태 항목 리스트
        - to_be: 개선 목표 항목 리스트
        - target_app: 대상 오피스 ('word')
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, "Insert Comparison"):
                return WordComponentRenderer.insert_comparison(word_app, doc, title, as_is, to_be)
        elif target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            return ExcelComponentRenderer.insert_comparison(excel_app, wb.ActiveSheet, title, as_is, to_be)
        return "현재 AS-IS vs TO-BE 비교표는 Word 및 Excel을 지원합니다."

    @mcp.tool()
    def office_insert_callout_box(
        text: str,
        box_type: Literal["info", "tip", "warning", "quote"] = "info",
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [비즈니스 컴포넌트 5] 핵심 요약 및 강조 콜아웃 상자 생성 도구.
        좌측 2.5pt 굵은 강조선과 파스텔 배경 음영이 적용된 프리미엄 박스를 생성합니다.
        - text: 상자 내에 들어갈 본문 텍스트
        - box_type: 'info'(블루 안내), 'tip'(녹색 팁), 'warning'(오렌지 주의), 'quote'(회색 요약)
        - target_app: 대상 오피스 ('word', 'hwp')
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, "Insert Callout"):
                return WordComponentRenderer.insert_callout_box(word_app, doc, text, box_type)
        elif target_app == "hwp":
            hwp = com_manager.get_hwp()
            if not hwp:
                return "열려 있는 한컴오피스 HWP 창을 찾을 수 없습니다."
            return HwpComponentRenderer.insert_callout_box(hwp, text, box_type)

        elif target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            return ExcelComponentRenderer.insert_callout_box(excel_app, wb.ActiveSheet, text, box_type)
        return f"{target_app}의 콜아웃 상자는 Word, Excel 또는 HWP를 지원합니다."

    @mcp.tool()
    def office_insert_kpi_cards(
        kpi_data: list[dict[str, str]],
        target_app: Literal["word", "excel", "ppt", "hwp"] = "word",
    ) -> str:
        """
        [비즈니스 컴포넌트 6] 핵심 성과 지표(KPI) 대시보드 카드 생성 도구.
        수치, 단위, 전년 대비 증감율 배지(▲ +15.4%)가 포함된 가로 카드 블록을 생성합니다.
        - kpi_data: [{'label': '연간 매출액', 'value': '1,240 억원', 'change': '+18.2%'}, ...]
        - target_app: 대상 오피스 ('word', 'ppt')
        """
        ensure_com()
        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."
            with word_undo_group(word_app, "Insert KPI"):
                return WordComponentRenderer.insert_kpi_cards(word_app, doc, kpi_data)
        elif target_app == "ppt":
            ppt_app, pres = com_manager.get_ppt()
            if not pres:
                return "열려 있는 PowerPoint 프레젠테이션을 찾을 수 없습니다."
            return PptComponentRenderer.insert_kpi_cards(ppt_app, pres, kpi_data)

        elif target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            return ExcelComponentRenderer.insert_kpi_cards(excel_app, wb.ActiveSheet, kpi_data)
        return f"{target_app}의 KPI 카드는 Word, Excel 또는 PPT를 지원합니다."


    @mcp.tool()
    def office_insert_approval_block(
        approvers: list[str] | None = None,
        target_app: Literal["excel", "word"] = "excel",
        start_cell: str = "G1",
    ) -> str:
        """
        [비즈니스 컴포넌트 8] KISA 공문서 및 기업 표준 결재선(결재란) 생성 도구.
        '결재' 란과 담당, 검토, 심사, 승인 직책 헤더 및 서명 영역을 생성합니다.
        - approvers: 결재자 직책 목록 (기본값: ['담당', '검토', '심사', '승인'])
        - target_app: 대상 오피스 ('excel')
        - start_cell: Excel 배치 시작 셀 (기본값: 'G1')
        """
        ensure_com()
        if target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            return ExcelComponentRenderer.insert_approval_block(excel_app, wb.ActiveSheet, approvers, start_cell=start_cell)
        return f"{target_app} 결재선은 추후 지원 예정입니다."

    @mcp.tool()
    def word_clean_and_structure() -> str:
        """
        [스마트 정제 도구 1] Word 문서 궁극의 7단계 개조식 위계 및 자간/줄간격 교정.
        범피스 107개 기관 서식의 정규식 엔진을 바탕으로 '1. -> 가. -> 1) -> 가) -> (1) -> (가) -> o'
        위계를 자동 인식하여 들여쓰기, 내어쓰기, 줄간격, 폰트 스타일을 즉시 정돈합니다.
        """
        ensure_com()
        word_app, doc = com_manager.get_word()
        if not doc:
            return "열려 있는 Word 문서를 찾을 수 없습니다."
        with word_undo_group(word_app, "Clean and Structure"):
            return WordComponentRenderer.clean_and_structure_text(word_app, doc)

    @mcp.tool()
    def excel_auto_format_table(range_address: str | None = None) -> str:
        """
        [스마트 정제 도구 2] Excel 시트 겉바속촉 자동 서식화 도구.
        범피스 '겉바속촉' 알고리즘: 외곽 굵은 테두리, 헤더 네이비 음영, 숫자 열 자동 천단위 콤마,
        내부 오픈 테이블(가로선만 유지), 열 너비 AutoFit을 원클릭으로 일괄 적용합니다.
        - range_address: 적용할 영역 주소 (예: 'A1:E15', 생략 시 전체 사용 영역 UsedRange)
        """
        ensure_com()
        excel_app, wb, sheet = com_manager.get_excel()
        if not sheet:
            return "열려 있는 Excel 시트를 찾을 수 없습니다."
        return ExcelComponentRenderer.auto_format_table(excel_app, sheet, range_address)

    @mcp.tool()
    def office_harmonize_document(
        font_family: str = "맑은 고딕",
        autofit_tables: bool = True,
        column_width_mode: Literal["content", "equal", "proportional"] = "content",
        auto_align_columns: bool = True,
        readability_style: Literal["open", "grid", "preserve"] = "open",
        zebra_striping: bool = True,
        repeat_header_row: bool = True,
        normalize_font_sizes: bool = True,
        theme_color: Literal["navy", "dark_grey", "blue", "teal"] | None = None,
        target_app: Literal["word", "excel", "ppt"] = "word",
    ) -> str:
        """
        [문서 일괄 조화 도구] 서로 다른 양식에서 가져온 컴포넌트들을 단일 톤앤매너로 일괄 통일하고 가독성을 극대화합니다.
        문서 전체 글꼴 통일, 표 역할(콜아웃/헤딩바/데이터표) 자동 감지, 헤더 줄바꿈 방지 가드, 오픈 테이블 스타일,
        데이터 성격별 지능형 정렬(숫자:우측, 순번/날짜:중앙, 설명:좌측), 교차 행 음영, 다중 페이지 헤더 반복을 한 번에 정규화합니다.
        - font_family: 통일할 대표 글꼴 (예: '맑은 고딕', 'Pretendard', 'KoPub돋움_Pro', 'Aptos')
        - autofit_tables: 문서 내 모든 표를 본문 폭(100%)에 맞춰 AutoFit하고 가운데 정렬 및 셀 여백 최적화
        - column_width_mode: 표 내부 칸(열) 너비 자동 맞춤 모드
          * 'content' (기본 추천): 헤더 줄바꿈 방지(헤더 텍스트 최소 폭 보장) + 내용 길이에 맞춘 지능형 폭 계산 후 본문 폭(100%) 비례 확장
          * 'equal': 모든 열의 너비를 수학적으로 균등 분할 (1/N 폭)
          * 'proportional': 기존 양식의 열 너비 비율을 보존하면서 전체 가로폭만 본문 100%에 맞춤
        - auto_align_columns: 표 데이터 성격에 따른 열별 자동 정렬 (헤더: 가운데, 숫자/금액/퍼센트: 오른쪽, 코드/순번/날짜/단어: 가운데, 설명/문장: 왼쪽)
        - readability_style: 표 시각 스타일 ('open': 세로선 제거 및 상하 1.5pt 강조의 현대적 오픈테이블, 'grid': 격자 유지, 'preserve': 기존 유지)
        - zebra_striping: 4행 이상 데이터 표의 짝수행에 극미세 음영(#F8FAFC) 자동 부여로 가로 시선 이탈 방지
        - repeat_header_row: 다중 페이지 표에서 매 페이지 상단에 헤더 행 자동 반복 및 행 중간 쪼개짐 방지
        - normalize_font_sizes: 대/중/소제목 및 본문 폰트 크기를 표준 스케일(22/16/13/10.5pt)로 정규화
        - theme_color: (선택) 표 헤더 및 강조 배경 색상 톤 통일 ('navy', 'dark_grey', 'blue', 'teal')
        - target_app: 대상 오피스 프로그램 ('word', 'excel', 'ppt')
        """
        ensure_com()
        pt_per_mm = 72 / 25.4
        theme_map = {
            "navy": 0x800000,       # RGB(0, 0, 128) -> BGR
            "dark_grey": 0x575450,   # RGB(80, 84, 87)
            "blue": 0xD2410E,       # RGB(14, 65, 210)
            "teal": 0x939A20,       # RGB(32, 154, 147)
        }

        def _est_text_width_pt(text: str, font_size_pt: float = 11.0) -> float:
            """한글/영문/숫자별 대략적인 텍스트 폭(pt) 계산 (헤더 어색한 줄바꿈 방지용 가드)"""
            w = 0.0
            for ch in text.strip():
                if ord(ch) >= 0xac00 or ord(ch) >= 0x4e00:
                    w += font_size_pt * 1.05
                elif ch.isspace():
                    w += font_size_pt * 0.35
                else:
                    w += font_size_pt * 0.58
            return w

        def _classify_column_alignment(values: list[str], header_text: str = "") -> int:
            """열 데이터 및 헤더 분석을 통한 최적 정렬 반환: 0 (왼쪽), 1 (가운데), 2 (오른쪽)"""
            clean_vals = [v.strip() for v in values if v.strip()]
            if not clean_vals:
                return 1  # 기본 가운데

            hdr_clean = header_text.strip().upper()
            is_index_header = any(k in hdr_clean for k in ["NO", "순번", "연번", "번호", "순위", "단계"])

            # 1. 순번/인덱스 열 감지 (헤더가 순번/NO이거나 값이 1~3자리의 단순 정수 나열) -> 가운데 정렬 (1)
            is_sequential_int = all(v.isdigit() and len(v) <= 3 for v in clean_vals)
            if is_index_header or is_sequential_int:
                return 1

            # 2. 금액, 수량, 퍼센트, 통화, 천단위 콤마 숫자 -> 우측 정렬 (2)
            numeric_count = 0
            has_currency_or_comma = False
            for val in clean_vals:
                if any(sym in val for sym in [",", "%", "원", "$", "₩", "억원", "백만원"]):
                    has_currency_or_comma = True
                cleaned = re.sub(r'[,%\s원\$₩억원백만원]', '', val)
                try:
                    float(cleaned)
                    numeric_count += 1
                except ValueError:
                    pass
            if (numeric_count / len(clean_vals) >= 0.7) and (has_currency_or_comma or not is_sequential_int):
                return 2

            # 3. 날짜, 상태, 코드, 단어 (평균 6자 이하의 짧은 단어) -> 가운데 정렬 (1)
            avg_len = sum(len(v) for v in clean_vals) / len(clean_vals)
            max_len = max(len(v) for v in clean_vals)
            if max_len <= 8 or avg_len <= 5.5:
                return 1

            # 4. 긴 서술형 텍스트, 설명문, 제목 -> 좌측 정렬 (0)
            return 0

        if target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."

            with word_undo_group(word_app, "Harmonize Document"):
                # 1. 폰트 패밀리 일괄 변경
                if font_family:
                    doc.Content.Font.Name = font_family
                    try:
                        doc.Content.Font.NameFarEast = font_family
                    except Exception:
                        pass
                    try:
                        doc.Content.Font.NameAscii = font_family
                    except Exception:
                        pass

                # 2. 표 자동 맞춤 및 역할별(박스/헤딩바/데이터표) 가독성 극대화
                table_count = doc.Tables.Count
                if autofit_tables and table_count > 0:
                    page = doc.Sections(1).PageSetup
                    body_width = page.PageWidth - page.LeftMargin - page.RightMargin
                    for tbl in list(doc.Tables):
                        try:
                            tbl.Rows.Alignment = 1  # 가운데 정렬
                            rows_cnt = tbl.Rows.Count
                            is_uniform = True
                            try:
                                col_count = tbl.Columns.Count
                            except Exception:
                                is_uniform = False
                                col_count = tbl.Rows(1).Cells.Count if rows_cnt > 0 else 1

                            # 표 역할 감지:
                            # 1) 콜아웃 박스 (1x1)
                            # 2) 인라인 헤딩 바 (1행 x 2~3열)
                            # 3) 실제 다단 데이터 표 (2행 이상 x 2열 이상)
                            is_callout_box = (rows_cnt == 1 and col_count == 1)
                            is_heading_bar = (rows_cnt == 1 and col_count in (2, 3))
                            is_data_table = (rows_cnt >= 2 and col_count >= 2)

                            if is_callout_box:
                                # 박스는 테두리 형태를 보존하고 상하 여백 숨통 확보
                                tbl.TopPadding = 1.2 * pt_per_mm
                                tbl.BottomPadding = 1.2 * pt_per_mm
                                tbl.LeftPadding = 1.5 * pt_per_mm
                                tbl.RightPadding = 1.5 * pt_per_mm
                                tbl.PreferredWidthType = 3
                                tbl.PreferredWidth = body_width

                            elif is_heading_bar:
                                # 소제목 배지 바는 배지 비율을 유지하며 본문 폭에 맞춤
                                tbl.TopPadding = 0.5 * pt_per_mm
                                tbl.BottomPadding = 0.5 * pt_per_mm
                                tbl.PreferredWidthType = 3
                                tbl.PreferredWidth = body_width

                            elif is_data_table:
                                # 정식 데이터 표 가독성 엔진 전면 가동
                                # 1) 수직 중앙 정렬 및 여백(숨통) 확보
                                tbl.TopPadding = 1.0 * pt_per_mm
                                tbl.BottomPadding = 1.0 * pt_per_mm
                                tbl.LeftPadding = 1.0 * pt_per_mm
                                tbl.RightPadding = 1.0 * pt_per_mm

                                for r_idx in range(1, rows_cnt + 1):
                                    for c in tbl.Rows(r_idx).Cells:
                                        c.VerticalAlignment = 1  # wdCellAlignVerticalCenter

                                # 2) 다중 페이지 분할 보호 (헤더 반복 및 행 중간 쪼개짐 방지)
                                if repeat_header_row:
                                    tbl.Rows(1).HeadingFormat = True
                                    tbl.Rows.AllowBreakAcrossPages = False

                                # 3) 오픈 테이블 스타일링 (세로선 제거, 상하 1.5pt 외곽선, 내부 0.5pt 가로선)
                                if readability_style == "open":
                                    tbl.Borders(-2).LineStyle = 0  # wdBorderLeft 없음
                                    tbl.Borders(-4).LineStyle = 0  # wdBorderRight 없음
                                    tbl.Borders(-6).LineStyle = 0  # wdBorderVertical 없음
                                    tbl.Borders(-1).LineStyle = 1  # wdBorderTop 실선
                                    tbl.Borders(-1).LineWidth = 12  # 1.5pt
                                    tbl.Borders(-3).LineStyle = 1  # wdBorderBottom 실선
                                    tbl.Borders(-3).LineWidth = 12  # 1.5pt
                                    tbl.Borders(-5).LineStyle = 1  # wdBorderHorizontal 내부 가로선
                                    tbl.Borders(-5).LineWidth = 4   # 0.5pt
                                    tbl.Borders(-5).Color = 0xD9D9D9  # 밝은 회색선

                                # 4) 칸(열) 너비 모드별 최적화 적용
                                if is_uniform and col_count > 0:
                                    if column_width_mode == "content":
                                        for col in tbl.Columns:
                                            col.AutoFit()
                                        raw_widths = [c.Width for c in tbl.Columns]

                                        # 헤더(1행) 텍스트 줄바꿈 방지 최소폭 가드
                                        min_header_widths = []
                                        for c_idx in range(1, col_count + 1):
                                            hdr_text = tbl.Cell(1, c_idx).Range.Text.replace("\r", "").replace("\x07", "")
                                            sz = tbl.Cell(1, c_idx).Range.Font.Size or 11.0
                                            hdr_pt = _est_text_width_pt(hdr_text, sz) + 12.0
                                            min_header_widths.append(hdr_pt)

                                        adjusted_raw = [max(r, m) for r, m in zip(raw_widths, min_header_widths)]
                                        total_raw = sum(adjusted_raw)
                                        if total_raw > 0:
                                            scale = body_width / total_raw
                                            for i, col in enumerate(tbl.Columns):
                                                col.Width = adjusted_raw[i] * scale

                                    elif column_width_mode == "equal":
                                        tbl.Columns.DistributeWidth()
                                        tbl.PreferredWidthType = 3
                                        tbl.PreferredWidth = body_width

                                    elif column_width_mode == "proportional":
                                        raw_widths = [c.Width for c in tbl.Columns]
                                        total_raw = sum(raw_widths)
                                        if total_raw > 0:
                                            scale = body_width / total_raw
                                            for i, col in enumerate(tbl.Columns):
                                                col.Width = raw_widths[i] * scale

                                    # 5) 열별 데이터 성격 자동 정렬 (auto_align_columns)
                                    if auto_align_columns:
                                        tbl.Rows(1).Range.ParagraphFormat.Alignment = 1  # 헤더 중앙
                                        for c_idx in range(1, col_count + 1):
                                            hdr_txt = tbl.Cell(1, c_idx).Range.Text.replace("\r", "").replace("\x07", "")
                                            col_texts = []
                                            for r_idx in range(2, rows_cnt + 1):
                                                try:
                                                    cell_txt = tbl.Cell(r_idx, c_idx).Range.Text.replace("\r", "").replace("\x07", "")
                                                    col_texts.append(cell_txt)
                                                except Exception:
                                                    pass
                                            best_align = _classify_column_alignment(col_texts, hdr_txt)
                                            for r_idx in range(2, rows_cnt + 1):
                                                try:
                                                    tbl.Cell(r_idx, c_idx).Range.ParagraphFormat.Alignment = best_align
                                                except Exception:
                                                    pass

                                # 6) 교차 행 얼룩말 음영 (zebra_striping, 4행 이상 표에 적용)
                                if zebra_striping and rows_cnt >= 4:
                                    for r_idx in range(2, rows_cnt + 1):
                                        if r_idx % 2 == 0:
                                            tbl.Rows(r_idx).Shading.BackgroundPatternColor = 0xF8FAFC  # 미세 블루그레이

                                tbl.PreferredWidthType = 3
                                tbl.PreferredWidth = body_width

                            else:
                                # 비정형 표
                                tbl.PreferredWidthType = 3
                                tbl.PreferredWidth = body_width

                        except Exception:
                            pass

                # 3. 폰트 크기 계층 정규화
                if normalize_font_sizes:
                    for p in list(doc.Paragraphs):
                        try:
                            sz = p.Range.Font.Size
                            if sz >= 20:
                                p.Range.Font.Size = 22  # 대제목
                            elif 15 <= sz < 20:
                                p.Range.Font.Size = 16  # 중제목
                            elif 12 <= sz < 15:
                                p.Range.Font.Size = 13.5  # 소제목
                            elif 7 <= sz < 12:
                                p.Range.Font.Size = 10.5  # 본문 / 표 내용
                        except Exception:
                            pass

                # 4. 테마 색상 정규화 (표 헤더 1행 음영)
                if theme_color and theme_color in theme_map and table_count > 0:
                    color_val = theme_map[theme_color]
                    for tbl in list(doc.Tables):
                        try:
                            if tbl.Rows.Count >= 2:
                                tbl.Rows(1).Shading.BackgroundPatternColor = color_val
                                tbl.Rows(1).Range.Font.Color = 0xFFFFFF  # 흰색 글씨
                        except Exception:
                            pass

            com_manager.activate_app("word")
            return f"Word 문서 '{doc.Name}' 조화 완료: 폰트 '{font_family}', 표 {table_count}개 자동 맞춤, 위계 크기 정규화 적용됨."

        elif target_app == "excel":
            excel_app, wb, sheet = com_manager.get_excel()
            if not sheet:
                return "열려 있는 Excel 시트를 찾을 수 없습니다."
            if font_family:
                sheet.Cells.Font.Name = font_family
            if autofit_tables:
                sheet.UsedRange.Columns.AutoFit()
            com_manager.activate_app("excel")
            return f"Excel 시트 '{sheet.Name}' 조화 완료: 폰트 '{font_family}', 열 너비 자동 맞춤 적용됨."

        elif target_app == "ppt":
            ppt_app, pres = com_manager.get_ppt()
            if not pres:
                return "열려 있는 PowerPoint 프레젠테이션을 찾을 수 없습니다."
            count = 0
            for sld in pres.Slides:
                for shp in sld.Shapes:
                    if shp.HasTextFrame and font_family:
                        try:
                            shp.TextFrame.TextRange.Font.Name = font_family
                            shp.TextFrame.TextRange.Font.NameFarEast = font_family
                            count += 1
                        except Exception:
                            pass
            com_manager.activate_app("ppt")
            return f"PowerPoint '{pres.Name}' 조화 완료: 폰트 '{font_family}' ({count}개 텍스트 상자 적용됨)."

        return f"{target_app}은 현재 지원하지 않는 프로그램입니다."
