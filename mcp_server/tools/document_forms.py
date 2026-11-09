# -*- coding: utf-8 -*-
"""
mcp_server/tools/document_forms.py
KISA 공공/비즈니스 표준 문서의 HWP 서식 레시피를 MS Word에서 그대로 실행하는 호환 계층.

공공/비즈니스 표준 문서 서식은 `표만들기`, `표오른쪽`, `문장` 같은 기본 동작을 나열한 함수다. 기본 동작만 Word로
다시 구현하고(WordFormBackend), 복원된 서식 함수에 이 객체를 self로 넘겨 실행한다.

디컴파일 코드는 분기가 깨지거나 본문이 사라진 곳이 있어 그대로 믿을 수 없다. 그래서 기본 동작은
재작성하고, 서식은 분기·HWP 직접 호출·소실된 본문이 없는 것만 실행한다(_blocker 참고).
"""

from __future__ import annotations

import ast
import json
import math
import re
import types
from collections import Counter, defaultdict
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, rgb_to_bgr, word_undo_group
from .component_tools import WD_BORDER_BOTTOM, WD_BORDER_LEFT, WD_BORDER_RIGHT, WD_BORDER_TOP
from .multi_doc_form_renderer import MultiDocFormRenderer

REFERENCE_DIR = Path(__file__).resolve().parents[2] / "reference" / "extracted_KISA1310" / "한컴오피스"
RECIPE_FILES = (  # 같은 이름은 뒤 파일이 우선
    ("한컴라이브러리_기본.py", "한컴기본"),
    ("한컴라이브러리_1.py", "한컴1"),
    ("한컴라이브러리_2.py", "한컴2"),
    ("한컴라이브러리_3.py", "한컴3"),
    ("한컴라이브러리.py", "기본한컴"),
)

PT_PER_MM = 72 / 25.4
WD_INSIDE_HORIZONTAL = -5
WD_INSIDE_VERTICAL = -6
WD_LINE_WIDTHS = (2, 4, 6, 8, 12, 18, 24, 36, 48)  # WdLineWidth, 1/8pt 단위
# HWP 테두리 종류 -> WdLineStyle (0 없음, 1 실선, 2·6 파선, 3·7 점선, 4 일점쇄, 5 이점쇄, 8 이중, 9~11 굵기 혼합)
HWP_LINE_STYLES = {0: 0, 1: 1, 2: 4, 3: 2, 4: 5, 5: 6, 6: 4, 7: 2, 8: 7, 9: 9, 10: 10, 11: 11}
# HWP 테두리 굵기 번호 -> mm
HWP_LINE_WIDTHS_MM = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0)
LINE_SPACING_BASE = 130  # HWP 줄간격 %는 글자 크기 기준, Word 배수는 글꼴 줄 높이 기준이라 보정
WORD_MIN_COLUMN_PT = 11  # Word는 패딩이 0이어도 약 10.9pt보다 좁은 칸을 거부한다
RESIZE_STEP_PT = 1 * PT_PER_MM
CELL_PADDING_MM = 0.5  # HWP 기본값은 1.8이지만 Word는 같은 글꼴도 넓게 잡혀 시간표 칸이 꺾인다. 측정으로 정한 값
FORM_MARGINS_MM = (20, 20, 15, 15, 10, 10)  # 공공/비즈니스 표준 문서 서식이 쓰는 문서여백 (좌, 우, 위, 아래, 머리, 꼬리)


def hwp_line_width(index: int) -> int:
    """HWP 테두리 굵기 번호를 가장 가까운 WdLineWidth로 변환"""
    eighths = HWP_LINE_WIDTHS_MM[index] * PT_PER_MM * 8
    return min(WD_LINE_WIDTHS, key=lambda width: abs(width - eighths))


def fit_widths(widths: list[float], available: float) -> list[float]:
    """칸 너비를 본문 폭에 맞춘다. 너무 좁은 칸은 최소 너비로 올리고 그만큼 넓은 칸에서 비례해 덜어낸다"""
    if len(widths) * WORD_MIN_COLUMN_PT > available:
        raise ValueError(f"칸 {len(widths)}개는 본문 폭에 들어가지 않습니다.")
    scale = min(1.0, available / sum(widths))
    fitted = [max(w * scale, WORD_MIN_COLUMN_PT) for w in widths]
    excess = sum(fitted) - available
    if excess <= 0:
        return fitted
    room = sum(w - WORD_MIN_COLUMN_PT for w in fitted)
    return [w - (w - WORD_MIN_COLUMN_PT) * excess / room for w in fitted]


class FormGrid:
    """병합을 반영한 표 커서 모델. HWP 셀 이동을 계산할 뿐 Word를 호출하지 않는다.

    행마다 칸을 [시작 열, 끝 열, 세로 병합 연속 칸 여부]로 둔다. 칸의 목록 번호는 Word의 Cell(r, c) 번호와 같다.
    """

    def __init__(self, rows: int, cols: int) -> None:
        self.rows = [[[c, c + 1, False] for c in range(cols)] for _ in range(rows)]

    def order(self) -> list[tuple[int, int]]:
        """HWP 셀 이동 순서: 왼쪽에서 오른쪽, 행 끝에서 다음 행으로.

        세로 병합 연속 칸도 한 칸으로 센다. 기본서식공모일정의 행 넘김 이동 횟수가 이 규칙이어야 맞는다.
        """
        return [(r, i) for r, row in enumerate(self.rows) for i in range(len(row))]

    def step(self, pos: tuple[int, int], count: int) -> tuple[int, int]:
        order = self.order()
        return order[max(0, min(len(order) - 1, order.index(pos) + count))]

    def owner(self, pos: tuple[int, int]) -> tuple[int, int]:
        """세로 병합 연속 칸이면 실제 셀인 시작 칸을 반환"""
        r, i = pos
        col = self.rows[r][i][0]
        while self.rows[r][i][2]:
            r -= 1
            i = next(n for n, slot in enumerate(self.rows[r]) if slot[0] <= col < slot[1])
        return r, i

    def vertical(self, pos: tuple[int, int], count: int) -> tuple[int, int]:
        """같은 열 위치를 유지한 채 위아래로 이동. 세로 병합 연속 칸이면 시작 칸으로 올라간다"""
        r, i = pos
        col = self.rows[r][i][0]
        target = max(0, min(len(self.rows) - 1, r + count))
        index = next(n for n, slot in enumerate(self.rows[target]) if slot[0] <= col < slot[1])
        return self.owner((target, index))

    def rect(self, a: tuple[int, int], b: tuple[int, int]) -> tuple[int, int, int, int]:
        """두 칸을 감싸는 사각형 (시작 행, 끝 행, 시작 열, 끝 열)"""
        slot_a, slot_b = self.rows[a[0]][a[1]], self.rows[b[0]][b[1]]
        return (
            min(a[0], b[0]),
            max(a[0], b[0]),
            min(slot_a[0], slot_b[0]),
            max(slot_a[1], slot_b[1]),
        )

    def cells_in(self, rect: tuple[int, int, int, int]) -> list[tuple[int, int]]:
        r0, r1, c0, c1 = rect
        return [
            (r, i)
            for r in range(r0, r1 + 1)
            for i, slot in enumerate(self.rows[r])
            if not slot[2] and c0 <= slot[0] and slot[1] <= c1
        ]

    def merge(self, rect: tuple[int, int, int, int]) -> tuple[tuple[int, int], tuple[int, int]]:
        """사각형을 한 칸으로 합치고, 병합 전 기준 (왼쪽 위 칸, 오른쪽 아래 칸) 번호를 반환"""
        r0, r1, c0, c1 = rect
        first = (r0, next(i for i, slot in enumerate(self.rows[r0]) if slot[0] == c0))
        last = (r1, next(i for i, slot in enumerate(self.rows[r1]) if slot[1] == c1))
        for r in range(r0, r1 + 1):
            row = self.rows[r]
            start = next(i for i, slot in enumerate(row) if slot[0] == c0)
            end = next(i for i, slot in enumerate(row) if slot[1] == c1)
            row[start : end + 1] = [[c0, c1, r != r0]]
        return first, last


class _HwpShim:
    """레시피의 self.대상.HAction.Run('...') 및 self.대상.MovePos(n) 직접 호출을 기본 동작으로 넘긴다"""

    def __init__(self, backend: "WordFormBackend") -> None:
        self.backend = backend
        self.HAction = self

    def Run(self, action: str) -> None:
        HWP_RUN_ACTIONS[action](self.backend)

    def MovePos(self, code: int) -> None:
        HWP_MOVE_ACTIONS[code](self.backend)


class WordFormBackend:
    """공공/비즈니스 표준 문서 기본 동작의 Word 구현. 메서드 이름은 공공/비즈니스 표준 문서와 같아서 서식 함수가 그대로 호출한다."""

    def __init__(self, word_app: Any, doc: Any) -> None:
        self.대상 = _HwpShim(self)
        self.after_table = False  # 표를 막 나온 직후 BreakPara는 이미 있는 문단을 쓴다
        self.block_extends = 0
        self.doc = doc
        self.sel = word_app.Selection
        self.pos = self.sel.Start  # 표 밖 커서
        self.doc_is_empty = len(doc.Content.Text) <= 1
        self.char: dict[str, Any] = {}  # 다음 `문장`에 적용할 글자 모양
        self.outer_char: dict[str, Any] = {}
        self.tbl: Any = None
        self.created_tables: list[Any] = []
        self.grid: FormGrid | None = None
        self.cur = (0, 0)
        self.anchor: tuple[int, int] | None = None
        self.block_all = False

    def __getattr__(self, name: str) -> Any:
        # 공공/비즈니스 표준 문서의 다른 서식 함수는 여기서 찾는다. 기본 동작이 없거나 HWP 전용이면 AttributeError
        reason = _blocker(name)
        if reason:
            raise AttributeError(f"Word에서 실행할 수 없습니다: {name} ({reason})")
        return types.MethodType(_compiled(name), self)

    # --- 공통 --------------------------------------------------------------
    def _cell(self, pos: tuple[int, int] | None = None) -> Any:
        r, i = self.grid.owner(pos or self.cur)
        return self.tbl.Cell(r + 1, i + 1)

    def _caret(self) -> Any:
        if self.tbl is None:
            return self.doc.Range(self.pos, self.pos)
        rng = self._cell().Range
        rng.MoveEnd(1, -1)  # 셀 끝 표시 앞까지
        rng.Collapse(0)  # wdCollapseEnd
        return rng

    def _block_rect(self) -> tuple[int, int, int, int] | None:
        if self.tbl is None:
            return None
        if self.block_all:
            return 0, len(self.grid.rows) - 1, 0, len(self.grid.rows[0]) and max(s[1] for s in self.grid.rows[0])
        return self.grid.rect(self.anchor if self.anchor else self.cur, self.cur)

    def _block_ranges(self) -> list[Any]:
        """블록이 있으면 대상 셀 범위 목록, 없으면 빈 목록"""
        if self.block_all:
            return [self.tbl.Range]
        if self.anchor is None:
            return []
        return [self._cell(pos).Range for pos in self.grid.cells_in(self._block_rect())]

    def _set_char(self, **changes: Any) -> None:
        ranges = self._block_ranges()
        for rng in ranges:
            self._apply_char(rng, changes)
        if not ranges:
            self.char.update(changes)

    def _set_para(self, apply: Callable[[Any], None]) -> None:
        for rng in self._block_ranges() or [self._caret()]:
            apply(rng.ParagraphFormat)

    @staticmethod
    def _apply_char(rng: Any, char: dict[str, Any]) -> None:
        if "font" in char:
            font_name = char["font"]
            rng.Font.Name = font_name
            try:
                rng.Font.NameFarEast = font_name
            except Exception:
                pass
            try:
                rng.Font.NameAscii = font_name
            except Exception:
                pass
        if "size" in char:
            rng.Font.Size = max(1.0, round(float(char["size"]) * 2) / 2)  # Word는 1pt 이상, 0.5pt 단위만 받는다
        if "bold" in char:
            rng.Font.Bold = -1 if char["bold"] else 0
        if "italic" in char:
            rng.Font.Italic = -1 if char["italic"] else 0
        if "underline" in char:
            rng.Font.Underline = 1 if char["underline"] else 0
        if "color" in char:
            rng.Font.Color = char["color"]
        if "spacing" in char:
            rng.Font.Spacing = float(char["spacing"])
        if "scaling" in char:
            rng.Font.Scaling = int(char["scaling"])
        if "position" in char:
            rng.Font.Position = int(char["position"])
        if "shade" in char:
            rng.Shading.BackgroundPatternColor = char["shade"]

    def _move(self, target: tuple[int, int]) -> None:
        self.cur = target
        self.block_all = False  # 레시피는 표전체 뒤에 캔슬 없이 표처음으로 빠져나온다
        if self.anchor is None:
            self.char = {}  # 셀을 옮기면 그 셀의 글자 모양을 따른다

    def _edge_cells(self, side: int) -> list[tuple[Any, int]]:
        """블록 바깥선을 이루는 (셀, 방향) 목록"""
        r0, r1, c0, c1 = self._block_rect()
        cells = []
        for pos in self.grid.cells_in((r0, r1, c0, c1)):
            slot = self.grid.rows[pos[0]][pos[1]]
            on_edge = {
                WD_BORDER_TOP: pos[0] == r0,
                WD_BORDER_BOTTOM: pos[0] == r1 or self._span_end(pos) == r1,
                WD_BORDER_LEFT: slot[0] == c0,
                WD_BORDER_RIGHT: slot[1] == c1,
            }[side]
            if not on_edge:
                continue
            above = self._same_span_above(pos) if side == WD_BORDER_TOP else None
            # 병합 행에서 Word는 셀 윗선을 옆 칸까지 번지게 하므로, 같은 경계를 위 칸의 아랫선으로 긋는다
            cells.append((self._cell(above), WD_BORDER_BOTTOM) if above else (self._cell(pos), side))
        return cells

    def _same_span_above(self, pos: tuple[int, int]) -> tuple[int, int] | None:
        """바로 위 행에 열 범위가 같은 실제 셀이 있으면 그 칸을 반환"""
        r, i = pos
        c0, c1, _ = self.grid.rows[r][i]
        if r == 0 or c1 - c0 < 2:  # 번짐은 병합 칸에서만 생긴다. 일반 칸은 윗선을 그대로 쓴다
            return None
        for index, (start, end, continued) in enumerate(self.grid.rows[r - 1]):
            if (start, end) == (c0, c1) and not continued:
                return r - 1, index
        return None

    def _span_end(self, pos: tuple[int, int]) -> int:
        """세로 병합 칸의 마지막 행 번호"""
        r, i = pos
        col = self.grid.rows[r][i][0]
        while r + 1 < len(self.grid.rows):
            below = next((s for s in self.grid.rows[r + 1] if s[0] <= col < s[1]), None)
            if below is None or not below[2]:
                break
            r += 1
        return r

    def _border(self, target: Any, side: int, style: int | None, width: int | None, color: int | None) -> None:
        border = target.Borders(side)
        if style is not None:
            border.LineStyle = HWP_LINE_STYLES[style]
        if not border.LineStyle:
            return
        if width is not None:
            try:
                border.LineWidth = hwp_line_width(width)
            except Exception:
                pass  # 이중선 등은 허용 굵기가 정해져 있어 Word 기본값을 유지
        if color is not None:
            border.Color = color

    def _edges(self, sides: dict[int, tuple[int | None, int | None, int | None]]) -> None:
        """블록 바깥 테두리 설정. sides는 {방향: (종류, 굵기, 색)}"""
        for side, (style, width, color) in sides.items():
            targets = [(self.tbl, side)] if self.block_all else self._edge_cells(side)
            for target, target_side in targets:
                self._border(target, target_side, style, width, color)

    # --- 문서 --------------------------------------------------------------
    def 새창(self) -> None:
        pass  # 임의로 새 문서를 만들지 않고 현재 커서에 삽입

    def 문서여백(self, 왼쪽: float, 오른쪽: float, 위: float, 아래: float, 머리: float, 꼬리: float) -> None:
        if not self.doc_is_empty:
            return  # 이미 내용이 있는 문서의 여백은 바꾸지 않는다
        page = self.doc.Range(self.pos, self.pos).Sections(1).PageSetup
        page.LeftMargin = 왼쪽 * PT_PER_MM
        page.RightMargin = 오른쪽 * PT_PER_MM
        page.TopMargin = 위 * PT_PER_MM
        page.BottomMargin = 아래 * PT_PER_MM
        page.HeaderDistance = 머리 * PT_PER_MM
        page.FooterDistance = 꼬리 * PT_PER_MM

    def 문단여백측정(self) -> float:
        page = self.doc.Range(self.pos, self.pos).Sections(1).PageSetup
        return round((page.LeftMargin + page.RightMargin) / PT_PER_MM, 1)

    def 문단여백측정세로(self) -> float:
        page = self.doc.Range(self.pos, self.pos).Sections(1).PageSetup
        return round((page.TopMargin + page.BottomMargin + page.HeaderDistance + page.FooterDistance) / PT_PER_MM, 1)

    def _text_width(self) -> float:
        page = self.doc.Range(self.pos, self.pos).Sections(1).PageSetup
        return page.PageWidth - page.LeftMargin - page.RightMargin

    # --- 글자 --------------------------------------------------------------
    def 문장(self, 내용: str) -> None:
        text = 내용.replace("\r\n", "\r").replace("\n", "\r")
        rng = self._caret()
        rng.InsertAfter(text)
        self._apply_char(rng, self.char)
        self.after_table = False
        if self.tbl is None:
            self.pos = rng.End

    def 엔터(self, 횟수: int) -> None:
        self.문장("\r" * 횟수)

    def _break_para(self) -> None:
        if self.after_table:
            self.after_table = False
        else:
            self.엔터(1)

    def _break_page(self) -> None:
        rng = self._caret()
        rng.InsertBreak(7)  # wdPageBreak
        self.pos = rng.End

    def 폰트(self, 폰트: str) -> None:
        self._set_char(font=폰트)

    def 글자크기(self, 크기: float) -> None:
        self._set_char(size=크기)

    def 글자색(self, 빨강: int, 초록: int, 파랑: int) -> None:
        self._set_char(color=rgb_to_bgr(빨강, 초록, 파랑))

    def 진하게(self) -> None:
        self._set_char(bold=True, italic=False, underline=False)

    def 기본글자(self) -> None:
        self._set_char(bold=False, italic=False, underline=False)

    def 밑줄기본(self) -> None:
        self._set_char(underline=True)

    def 글자간격(self, 간격: float) -> None:
        size = self.char.get("size", 10)
        self._set_char(spacing=size * 간격 / 100)

    def 글자장평(self, 장평: int) -> None:
        self._set_char(scaling=장평)

    def 글자위치(self, 값: float) -> None:
        self._set_char(position=self.char.get("size", 10) * 값 / 100)

    def 글자음영(self, 색상: int) -> None:
        self._set_char(shade=색상)

    # 공공/비즈니스 표준 문서 글꼴 묶음. 한양 계열은 일반 PC에 없어 가까운 글꼴로 대체
    def 중고딕(self) -> None:
        self.폰트("맑은 고딕")

    def 견고딕(self) -> None:
        self.폰트("맑은 고딕")
        self._set_char(bold=True)

    def 견명조(self) -> None:
        self.폰트("휴먼명조")
        self._set_char(bold=True)

    def 휴먼명조(self) -> None:
        self.폰트("휴먼명조")

    def 오늘날짜한글숫자만(self) -> None:
        now = datetime.today()
        self.문장(f"{now.year}년 {now.month}월 {now.day}일")

    # --- 문단 --------------------------------------------------------------
    def _align(self, value: int) -> None:
        self._set_para(lambda para: setattr(para, "Alignment", value))

    def 왼쪽정렬(self) -> None:
        self._align(0)

    def 가운데정렬(self) -> None:
        self._align(1)

    def 오른쪽정렬(self) -> None:
        self._align(2)

    def 기본정렬(self) -> None:
        self._align(3)

    def 줄간격(self, 간격: float) -> None:
        def apply(para: Any) -> None:
            para.LineSpacingRule = 5  # wdLineSpaceMultiple
            para.LineSpacing = 12 * 간격 / LINE_SPACING_BASE

        self._set_para(apply)

    # HWP 문단 값은 2배로 저장돼 공공/비즈니스 표준 문서가 pt에 2를 곱해 넘긴다. Word에는 pt 그대로 쓴다
    def 내어쓰기(self, 값: float) -> None:
        def apply(para: Any) -> None:
            para.LeftIndent = max(-값, 0)
            para.FirstLineIndent = 값

        self._set_para(apply)

    def 문단위(self, 간격: float) -> None:
        self._set_para(lambda para: setattr(para, "SpaceBefore", 간격))

    def 문단아래(self, 간격: float) -> None:
        self._set_para(lambda para: setattr(para, "SpaceAfter", 간격))

    def 문단여백(self, 좌측: float, 우측: float) -> None:
        def apply(para: Any) -> None:
            para.LeftIndent = 좌측
            para.RightIndent = 우측

        self._set_para(apply)

    def 자간헌터(self, 모드: int) -> None:
        pass  # HWP 줄나눔 옵션, Word에 대응 기능 없음

    # --- 표 ----------------------------------------------------------------
    def 표만들기(self, 가로크기: list[float], 세로크기: list[float]) -> None:
        if self.tbl is not None:
            raise NotImplementedError("표 안의 표는 지원하지 않습니다.")

        self._separate_from_previous_table()
        caret = self._caret()
        align = caret.ParagraphFormat.Alignment
        widths = fit_widths([w * PT_PER_MM for w in 가로크기], self._text_width())

        self.outer_char, self.char = self.char, {}
        tbl = self.doc.Tables.Add(caret, len(세로크기), len(가로크기))
        tbl.AllowAutoFit = False
        tbl.Rows.Alignment = align if align in (0, 1, 2) else 0
        tbl.Borders.Enable = True  # HWP 새 표는 전체 실선
        # 여백이 남아 있으면 Word가 좁은 칸 너비를 거부하므로 너비를 먼저 정한다
        tbl.LeftPadding = tbl.RightPadding = 0
        tbl.TopPadding = tbl.BottomPadding = 0.5 * PT_PER_MM
        tbl.Range.ParagraphFormat.Alignment = 3
        tbl.Range.ParagraphFormat.SpaceBefore = 0
        tbl.Range.ParagraphFormat.SpaceAfter = 0
        tbl.Range.ParagraphFormat.LineSpacingRule = 0
        tbl.Range.Cells.VerticalAlignment = 1  # HWP 셀 기본값은 세로 가운데
        for c, width in enumerate(widths, 1):
            tbl.Columns(c).Width = width
        tbl.LeftPadding = tbl.RightPadding = min(CELL_PADDING_MM * PT_PER_MM, min(widths) * 0.4)
        for r, height in enumerate(세로크기, 1):
            tbl.Rows(r).HeightRule = 1  # wdRowHeightAtLeast
            tbl.Rows(r).Height = height * PT_PER_MM

        self.tbl = tbl
        self.created_tables.append(tbl)
        self.grid = FormGrid(len(세로크기), len(가로크기))
        self.cur, self.anchor, self.block_all = (0, 0), None, False

    def _separate_from_previous_table(self) -> None:
        """Word는 맞붙은 표를 하나로 합치므로 바로 앞이 표면 1pt 빈 문단을 끼운다"""
        if self.pos == 0 or not self.doc.Range(self.pos - 1, self.pos).Information(12):  # wdWithInTable
            return
        spacer = self.doc.Range(self.pos, self.pos)
        spacer.InsertBefore("\r")
        spacer.Font.Size = 1
        spacer.ParagraphFormat.SpaceBefore = 0
        spacer.ParagraphFormat.SpaceAfter = 0
        spacer.ParagraphFormat.LineSpacingRule = 0
        self.pos = spacer.End

    def _leave_table(self) -> None:
        end = self.tbl.Range.End
        following = self.doc.Range(end, end)
        if following.Paragraphs(1).Range.Text.strip():
            following.InsertParagraphBefore()  # 뒤 내용과 붙지 않도록
        self.tbl, self.grid = None, None
        self.anchor, self.block_all = None, False
        self.char = self.outer_char
        self.pos = end
        self._caret().ParagraphFormat.Alignment = 3

    def 표나가기(self) -> None:
        self._leave_table()

    def 표나가기2(self) -> None:
        self._leave_table()

    def 표처음(self) -> None:
        self.anchor = None
        self._move((0, 0))

    def 표오른쪽(self, 횟수: int = 1) -> None:
        # 공공/비즈니스 표준 문서의 표오른쪽은 TableRightCellAppend라서 표 끝을 넘으면 행을 추가한다
        for _ in range(횟수):
            if self.anchor is None and self.cur == self.grid.order()[-1]:
                self.tbl.Rows.Add()
                self.grid.rows.append([[start, end, False] for start, end, _ in self.grid.rows[-1]])
            self._move(self.grid.step(self.cur, 1))

    def 표왼쪽(self, 횟수: int = 1) -> None:
        self._move(self.grid.step(self.cur, -횟수))

    def 표아래쪽(self, 횟수: int = 1) -> None:
        self._move(self.grid.vertical(self.cur, 횟수))

    def 표위쪽(self, 횟수: int = 1) -> None:
        self._move(self.grid.vertical(self.cur, -횟수))

    def 표오른쪽끝(self) -> None:
        self._move((self.cur[0], len(self.grid.rows[self.cur[0]]) - 1))

    def 표왼쪽끝(self) -> None:
        self._move((self.cur[0], 0))

    def 표위쪽끝(self) -> None:
        self._move(self.grid.vertical(self.cur, -self.cur[0]))

    def 표아래쪽끝(self) -> None:
        self._move(self.grid.vertical(self.cur, len(self.grid.rows)))

    def 셀선택(self) -> None:
        self.anchor, self.block_all = self.cur, False

    def 표전체(self) -> None:
        self.block_all = True

    def 캔슬(self) -> None:
        self.anchor, self.block_all = None, False

    def _block_start(self) -> None:
        self.셀선택()
        self.block_extends = 0

    def _block_extend(self) -> None:
        # HWP F5 누적: Block + Extend는 한 칸, Extend를 한 번 더 하면 표 전체
        self.block_extends += 1
        if self.block_extends >= 2:
            self.표전체()

    def _haction_move_right(self) -> None:
        """HWP MoveRight: 표 안에서 마지막 셀이 아니면 오른쪽 셀로, 마지막 셀이면 표 밖으로 나간다"""
        if self.tbl is not None:
            if self.grid is not None and self.cur != self.grid.order()[-1]:
                self.표오른쪽(1)
            else:
                self._exit_table()

    def _exit_table(self) -> None:
        if self.tbl is not None:
            self._leave_table()
            self.after_table = True

    def 표너비줄이기(self, 횟수: int) -> None:
        # ponytail: HWP TableResizeExLeft 1회 이동량 미확인, 원본 비교가 가능해지면 RESIZE_STEP_PT 보정
        r, i = self.grid.owner(self.cur)
        cell = self.tbl.Cell(r + 1, i + 1)
        shrink = min(횟수 * RESIZE_STEP_PT, cell.Width - WORD_MIN_COLUMN_PT)
        if shrink <= 0:
            return
        cell.Width -= shrink
        if i + 1 < len(self.grid.rows[r]):
            neighbor = self.tbl.Cell(r + 1, i + 2)
            neighbor.Width += shrink  # 표 전체 폭 유지

    def 셀세로정렬(self, 숫자: int) -> None:
        for pos in self.grid.cells_in(self._block_rect()):
            self._cell(pos).VerticalAlignment = {0: 0, 1: 1, 2: 3}[숫자]

    def 셀여백지정(self, 상: float, 하: float, 좌: float, 우: float) -> None:
        for pos in self.grid.cells_in(self._block_rect()):
            cell = self._cell(pos)
            cell.TopPadding, cell.BottomPadding = 상 * PT_PER_MM, 하 * PT_PER_MM
            cell.LeftPadding, cell.RightPadding = 좌 * PT_PER_MM, 우 * PT_PER_MM

    def 셀한줄(self, 옵션: int) -> None:
        for pos in self.grid.cells_in(self._block_rect()):
            self._cell(pos).FitText = bool(옵션)  # HWP 자간 줄임에 가장 가까운 Word 기능

    def 표내부선색(self, 빨강: int, 초록: int, 파랑: int) -> None:
        color = rgb_to_bgr(빨강, 초록, 파랑)
        if self.block_all:
            self._border(self.tbl, WD_INSIDE_HORIZONTAL, None, None, color)
            self._border(self.tbl, WD_INSIDE_VERTICAL, None, None, color)

    def 표테두리단일선색(self, 종류: str, 빨강: int, 초록: int, 파랑: int) -> None:
        sides = {"상": WD_BORDER_TOP, "하": WD_BORDER_BOTTOM, "좌": WD_BORDER_LEFT, "우": WD_BORDER_RIGHT}
        self._edges({sides[종류]: (None, None, rgb_to_bgr(빨강, 초록, 파랑))})

    def 셀병합(self) -> None:
        first, last = self.grid.merge(self._block_rect())
        if first != last:  # 한 칸만 선택하면 HWP도 아무 일도 하지 않는다
            self.tbl.Cell(first[0] + 1, first[1] + 1).Merge(self.tbl.Cell(last[0] + 1, last[1] + 1))
        self.cur, self.anchor = first, None

    def 셀여백제로(self) -> None:
        self.tbl.LeftPadding = self.tbl.RightPadding = self.tbl.TopPadding = self.tbl.BottomPadding = 0

    def 표배경색(self, 빨강: int, 초록: int, 파랑: int) -> None:
        color = rgb_to_bgr(빨강, 초록, 파랑)
        if self.block_all:
            self.tbl.Shading.BackgroundPatternColor = color
            return
        rect = self._block_rect()
        for pos in self.grid.cells_in(rect):
            self._cell(pos).Shading.BackgroundPatternColor = color

    def 표테두리타입(self, 상: int, 하: int, 좌: int, 우: int) -> None:
        self._edges({
            WD_BORDER_TOP: (상, None, None),
            WD_BORDER_BOTTOM: (하, None, None),
            WD_BORDER_LEFT: (좌, None, None),
            WD_BORDER_RIGHT: (우, None, None),
        })

    def 표테두리굵기(self, 상: int, 하: int, 좌: int, 우: int) -> None:
        self._edges({
            WD_BORDER_TOP: (None, 상, None),
            WD_BORDER_BOTTOM: (None, 하, None),
            WD_BORDER_LEFT: (None, 좌, None),
            WD_BORDER_RIGHT: (None, 우, None),
        })

    def 표테두리색(self, 빨강: int, 초록: int, 파랑: int) -> None:
        color = rgb_to_bgr(빨강, 초록, 파랑)
        self._edges({side: (None, None, color) for side in (WD_BORDER_TOP, WD_BORDER_BOTTOM, WD_BORDER_LEFT, WD_BORDER_RIGHT)})

    def 표단일선(self, 위치: str, 굵기: int, 종류: int) -> None:
        sides = {"상": WD_BORDER_TOP, "하": WD_BORDER_BOTTOM, "좌": WD_BORDER_LEFT, "우": WD_BORDER_RIGHT}
        self._edges({sides[위치]: (종류, 굵기, None)})

    def 표테두리단일선(self, 종류: str, 굵기: int, 타입: int) -> None:
        self.표단일선(종류, 굵기, 타입)

    def 표내부선타입(self, 가로: int, 세로: int) -> None:
        if self.block_all:
            self._border(self.tbl, WD_INSIDE_HORIZONTAL, 가로, None, None)
            self._border(self.tbl, WD_INSIDE_VERTICAL, 세로, None, None)
            return
        r0, r1, c0, c1 = self._block_rect()
        for pos in self.grid.cells_in((r0, r1, c0, c1)):
            slot = self.grid.rows[pos[0]][pos[1]]
            cell = self._cell(pos)
            if self._span_end(pos) < r1:
                self._border(cell, WD_BORDER_BOTTOM, 가로, None, None)
            if slot[1] < c1:
                self._border(cell, WD_BORDER_RIGHT, 세로, None, None)


HWP_RUN_ACTIONS: dict[str, Callable[[WordFormBackend], None]] = {
    "BreakPara": WordFormBackend._break_para,
    "BreakPage": WordFormBackend._break_page,
    "Cancel": WordFormBackend.캔슬,
    "CharShapeBold": lambda b: b._set_char(bold=True),
    "CharShapeItalic": lambda b: b._set_char(italic=True),
    "CharShapeNormal": WordFormBackend.기본글자,
    "CharShapeUnderline": lambda b: b._set_char(underline=True),
    "CloseEx": WordFormBackend._exit_table,
    "InsertFixedWidthSpace": lambda b: b.문장(" "),
    "InsertTab": lambda b: b.문장("\t"),
    "MoveLineEnd": lambda b: None,  # 표 밖 커서는 항상 입력 끝에 있다
    "MoveRight": WordFormBackend._haction_move_right,  # 셀이 남았으면 오른쪽 셀로, 끝이면 표 밖으로
    "ParagraphShapeAlignCenter": WordFormBackend.가운데정렬,
    "ParagraphShapeAlignDistribute": lambda b: b._align(4),
    "ParagraphShapeAlignJustify": WordFormBackend.기본정렬,
    "ParagraphShapeAlignRight": WordFormBackend.오른쪽정렬,
    "TableCellBlock": WordFormBackend._block_start,
    "TableCellBlockExtend": WordFormBackend._block_extend,
    "TableColBegin": WordFormBackend.표왼쪽끝,
    "TableLeftCell": lambda b: b.표왼쪽(1),
    "TableLowerCell": lambda b: b.표아래쪽(1),
    "TableMergeCell": WordFormBackend.셀병합,
    "TableResizeExLeft": lambda b: b.표너비줄이기(1),
    "TableRightCell": lambda b: b.표오른쪽(1),
    "TableRightCellAppend": lambda b: b.표오른쪽(1),
    "TableUpperCell": lambda b: b.표위쪽(1),
}
HWP_MOVE_ACTIONS: dict[int, Callable[[WordFormBackend], None]] = {
    100: lambda b: b.표왼쪽(1),
    101: lambda b: b.표오른쪽(1),
    102: lambda b: b.표위쪽(1),
    103: lambda b: b.표아래쪽(1),
    104: WordFormBackend.표왼쪽끝,
    105: WordFormBackend.표오른쪽끝,
    106: WordFormBackend.표위쪽끝,
    107: WordFormBackend.표아래쪽끝,
}


# =============================================================================
# 공공/비즈니스 표준 문서 서식 불러오기와 실행 가능 판별
# =============================================================================
@lru_cache(maxsize=1)
def _recipes() -> dict[str, ast.FunctionDef]:
    """{서식 이름: AST}. 복원 파일에 문법 오류가 있어 import하지 않고 구문만 읽는다"""
    if not REFERENCE_DIR.exists():
        return {}

    found: dict[str, ast.FunctionDef] = {}
    for file_name, class_name in RECIPE_FILES:
        tree = ast.parse((REFERENCE_DIR / file_name).read_text(encoding="utf-8", errors="replace"))
        for node in tree.body:
            if isinstance(node, ast.ClassDef) and node.name == class_name:
                found.update({item.name: item for item in node.body if isinstance(item, ast.FunctionDef)})
    return found


@lru_cache(maxsize=None)
def _compiled(name: str) -> Callable[..., Any]:
    """_blocker를 통과한 함수만 하나씩 컴파일한다"""
    node = _recipes()[name]
    namespace: dict[str, Any] = {"datetime": datetime, "re": re}
    exec(compile(ast.Module([node], []), f"<공공/비즈니스 표준 문서:{name}>", "exec"), namespace)
    return namespace[name]


def _self_attrs(node: ast.FunctionDef) -> set[str]:
    return {
        n.attr
        for n in ast.walk(node)
        if isinstance(n, ast.Attribute) and isinstance(n.value, ast.Name) and n.value.id == "self"
    }


def _unsupported_hwp_call(node: ast.FunctionDef) -> str | None:
    """대응표에 없는 self.대상 직접 호출을 찾아 반환"""
    supported = 0
    for n in ast.walk(node):
        if not isinstance(n, ast.Call):
            continue
        target = ast.unparse(n.func)
        if target not in ("self.대상.HAction.Run", "self.대상.MovePos"):
            continue
        table = HWP_RUN_ACTIONS if target.endswith("Run") else HWP_MOVE_ACTIONS
        value = n.args[0].value if n.args and isinstance(n.args[0], ast.Constant) else None
        if value not in table:
            return f"{target}({value})"
        supported += 1
    total = sum(isinstance(n, ast.Attribute) and n.attr == "대상" for n in ast.walk(node))
    return None if total == supported else "HWP 직접 호출"


def _missing_arguments(node: ast.FunctionDef) -> str | None:
    """복원 때 인자가 사라진 서식 호출을 찾는다"""
    recipes = _recipes()
    for n in ast.walk(node):
        if not (isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)):
            continue
        callee = recipes.get(n.func.attr) if ast.unparse(n.func.value) == "self" else None
        if callee is None or n.func.attr in WordFormBackend.__dict__:
            continue
        required = len(callee.args.args) - 1 - len(callee.args.defaults)
        if len(n.args) < required:
            return f"{n.func.attr} 인자 누락"
    return None


def _is_private_char(char: str) -> bool:
    return 0xE000 <= ord(char) <= 0xF8FF or ord(char) >= 0xF0000


def _damage(node: ast.FunctionDef) -> str | None:
    """복원이 깨졌을 수 있는 함수면 사유를 반환"""
    body = [s for s in node.body if not (isinstance(s, ast.Expr) and isinstance(s.value, ast.Constant))]
    if all(isinstance(s, ast.Pass) for s in body):
        return "본문 소실"
    for n in ast.walk(node):
        if isinstance(n, (ast.If, ast.For, ast.While, ast.Try)):
            return "분기 포함"
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Constant):
            return "복원 손상"
        if isinstance(n, ast.Constant) and isinstance(n.value, str) and any(_is_private_char(c) for c in n.value):
            return "HWP 전용 문자"  # Word에서는 네모로 깨진다
    if any(isinstance(n, ast.Return) for s in body[:-1] for n in ast.walk(s)):
        return "복원 손상"
    return None


@lru_cache(maxsize=None)
def _blocker(name: str) -> str | None:
    """서식 또는 기본 동작을 Word에서 실행할 수 없는 사유. 실행 가능하면 None"""
    if name in WordFormBackend.__dict__:
        return None
    recipes = _recipes()
    if name not in recipes:
        return "함수 없음"
    node = recipes[name]
    damage = _damage(node) or _unsupported_hwp_call(node) or _missing_arguments(node)
    if damage:
        return damage
    _blocker_guard.add(name)
    try:
        for attr in sorted(_self_attrs(node) - {"대상"}):
            if attr in _blocker_guard:
                continue
            inner = _blocker(attr)
            if inner:
                return f"{attr}: {inner}"
    finally:
        _blocker_guard.discard(name)
    return None


_blocker_guard: set[str] = set()

CATALOG_FILE = Path(__file__).resolve().parent / "form_catalog.json"
_CATALOG_CACHE: dict[str, Any] | None = None


def _get_catalog() -> dict[str, Any]:
    global _CATALOG_CACHE
    if _CATALOG_CACHE is None:
        if CATALOG_FILE.exists():
            import json
            _CATALOG_CACHE = json.loads(CATALOG_FILE.read_text(encoding="utf-8"))
        else:
            _CATALOG_CACHE = {"by_clean": {}, "by_old": {}}
    return _CATALOG_CACHE


def _resolve_form_name(form_name: str) -> tuple[str, str | None]:
    """탈지자체 표준 컴포넌트 명칭을 내부 레시피 함수명으로 변환. (내부 레시피명, 오류사유)"""
    cat = _get_catalog()
    # 1. 탈지자체 표준 컴포넌트 명칭으로 들어온 경우
    if form_name in cat.get("by_clean", {}):
        return cat["by_clean"][form_name], None
    # 2. 만약 레시피 원본명과 직접 일치하는 경우도 차단 여부 검사
    if form_name in _recipes():
        # 지자체 명칭 직접 호출 방지
        clean = cat.get("by_old", {}).get(form_name)
        if clean:
            return form_name, None
    return "", f"존재하지 않는 서식 컴포넌트입니다: '{form_name}'"


def _uses_table(name: str, seen: set[str] | None = None) -> bool:
    seen = seen if seen is not None else set()
    if name == "표만들기":
        return True
    if name in seen or name not in _recipes():
        return False
    seen.add(name)
    return any(_uses_table(attr, seen) for attr in _self_attrs(_recipes()[name]))


def list_forms() -> dict[str, list[str]]:
    """실행 가능한 표 서식 {표준 컴포넌트 명칭: 인자 이름} (지자체/기관명 완전 배제)"""
    cat = _get_catalog()
    if not _recipes():
        return {
            name: list(item.get("required_args", []))
            for name, item in _get_bm25_searcher().knowledge.items()
        }
    clean_map = cat.get("by_clean", {})
    if not clean_map:
        return {
            name: [arg.arg for arg in node.args.args[1:]]
            for name, node in _recipes().items()
            if name not in WordFormBackend.__dict__ and _uses_table(name) and _blocker(name) is None
        }

    recipes = _recipes()
    result = {}
    for clean_name, orig_name in clean_map.items():
        node = recipes.get(orig_name)
        if node and orig_name not in WordFormBackend.__dict__ and _uses_table(orig_name) and _blocker(orig_name) is None:
            result[clean_name] = [arg.arg for arg in node.args.args[1:]]
    return result


def _set_cell_text_preserving_style(cell: Any, text: str) -> None:
    """셀의 기존 폰트·크기·굵기·색상·정렬·음영·테두리를 그대로 보존하면서 텍스트만 교체"""
    rng = cell.Range
    rng.MoveEnd(1, -1)  # 셀 끝 마커(\r\x07) 제외
    try:
        font_name = rng.Font.Name
        font_size = rng.Font.Size
        font_bold = rng.Font.Bold
        font_color = rng.Font.Color
        para_align = rng.ParagraphFormat.Alignment
    except Exception:
        font_name = font_size = font_bold = font_color = para_align = None

    rng.Text = str(text).replace("\r\n", "\r").replace("\n", "\r")

    try:
        if font_name:
            rng.Font.Name = font_name
        if font_size and font_size < 999999:
            rng.Font.Size = font_size
        if font_bold is not None and font_bold != 9999999:
            rng.Font.Bold = font_bold
        if font_color is not None and font_color != 9999999:
            rng.Font.Color = font_color
        if para_align is not None and para_align in (0, 1, 2, 3):
            rng.ParagraphFormat.Alignment = para_align
    except Exception:
        pass


def _customize_rendered_table(
    tbl: Any,
    headers: list[str] | None = None,
    rows: list[list[Any]] | None = None,
    replacements: dict[str, str] | None = None,
) -> None:
    """렌더링된 참조 양식 표의 디자인·병합·테두리를 보존한 채 문구 치환, 행 확장/축소, 셀 데이터를 주입한다."""
    # 1. 고정 문구 치환 (병합 표, 결재선, 동의서 등 모든 양식 지원)
    if replacements:
        for cell in tbl.Range.Cells:
            try:
                raw_text = cell.Range.Text[:-2]
            except Exception:
                continue
            new_text = raw_text
            compact_raw = re.sub(r"\s+", "", raw_text)
            for old_k, new_v in replacements.items():
                if old_k in new_text:
                    new_text = new_text.replace(old_k, str(new_v))
                elif compact_raw and re.sub(r"\s+", "", old_k) == compact_raw:
                    new_text = str(new_v)
            if new_text != raw_text:
                _set_cell_text_preserving_style(cell, new_text)

    # 2. 헤더(1행) 텍스트 주입
    if headers:
        for c_idx, h_val in enumerate(headers, 1):
            try:
                cell = tbl.Cell(1, c_idx)
                _set_cell_text_preserving_style(cell, str(h_val))
            except Exception:
                break

    # 3. 데이터 행 동적 조정 및 값 주입
    if rows is not None:
        start_row = 2 if (headers is not None or tbl.Rows.Count >= 2) else 1
        target_total_rows = (start_row - 1) + len(rows)

        # 기존 마지막 행 하단 마감선 및 내부 가로선 스타일 백업
        closing_style = closing_width = closing_color = None
        inner_style = inner_width = inner_color = None
        try:
            b_bot = tbl.Rows(tbl.Rows.Count).Borders(WD_BORDER_BOTTOM)
            closing_style, closing_width, closing_color = b_bot.LineStyle, b_bot.LineWidth, b_bot.Color
        except Exception:
            pass
        if tbl.Rows.Count >= start_row + 1:
            try:
                b_in = tbl.Rows(start_row).Borders(WD_BORDER_BOTTOM)
                inner_style, inner_width, inner_color = b_in.LineStyle, b_in.LineWidth, b_in.Color
            except Exception:
                pass

        # 행이 부족하면 참조 양식 마지막 행 서식을 복제하여 추가
        try:
            while tbl.Rows.Count < target_total_rows:
                tbl.Rows.Add()
            # 빈 행이 남으면 하단부터 정리
            while tbl.Rows.Count > target_total_rows and target_total_rows >= start_row:
                tbl.Rows(tbl.Rows.Count).Delete()
        except Exception:
            pass  # 세로 병합이 포함된 특수 표에서 Rows.Add/Delete가 제한될 경우 가능한 범위만 채움

        # 내부 가로선 및 하단 마감선 복원
        try:
            if inner_style is not None:
                for r_i in range(start_row, tbl.Rows.Count):
                    b = tbl.Rows(r_i).Borders(WD_BORDER_BOTTOM)
                    b.LineStyle = inner_style
                    if inner_style:
                        b.LineWidth = inner_width
                        b.Color = inner_color
            if closing_style is not None and tbl.Rows.Count >= start_row:
                b_last = tbl.Rows(tbl.Rows.Count).Borders(WD_BORDER_BOTTOM)
                b_last.LineStyle = closing_style
                if closing_style:
                    b_last.LineWidth = closing_width
                    b_last.Color = closing_color
        except Exception:
            pass

        # 데이터 셀 텍스트 주입
        for r_offset, row_vals in enumerate(rows):
            r_idx = start_row + r_offset
            if r_idx > tbl.Rows.Count:
                break
            for c_idx, val in enumerate(row_vals, 1):
                try:
                    cell = tbl.Cell(r_idx, c_idx)
                    _set_cell_text_preserving_style(cell, str(val))
                except Exception:
                    continue


def _extract_table_info(tbl: Any) -> dict[str, Any]:
    """생성/커스텀된 Word 표의 행·열 구조와 셀 미리보기를 추출"""
    rows_cnt = tbl.Rows.Count
    cols_cnt = tbl.Columns.Count
    preview: list[list[str]] = []
    for r in range(1, min(rows_cnt + 1, 10)):
        row_cells: list[str] = []
        for c in range(1, cols_cnt + 1):
            try:
                txt = tbl.Cell(r, c).Range.Text[:-2].replace("\r", " ").strip()
                row_cells.append(txt)
            except Exception:
                pass
        if row_cells:
            preview.append(row_cells)
    return {
        "rows": rows_cnt,
        "cols": cols_cnt,
        "cells_preview": preview,
    }


def render_form(
    word_app: Any,
    doc: Any,
    form_name: str,
    args: list[Any] | None = None,
    headers: list[str] | None = None,
    rows: list[list[Any]] | None = None,
    replacements: dict[str, str] | None = None,
) -> dict[str, Any] | None:
    """서식을 현재 커서에 삽입하고 선택적 커스텀(문구 치환, 헤더/행 주입)을 수행. 실패하면 호출 전 상태로 되돌린다"""
    backend = WordFormBackend(word_app, doc)
    table_info: dict[str, Any] | None = None
    try:
        with word_undo_group(word_app, f"공공/비즈니스 표준 문서 서식: {form_name}"):
            backend.문서여백(*FORM_MARGINS_MM)  # 레시피 표 폭은 이 여백 기준. 빈 문서에서만 적용된다
            getattr(backend, form_name)(*(args or []))
            if backend.tbl is not None:
                backend._leave_table()
            if backend.created_tables:
                target_tbl = backend.created_tables[-1]
                if headers is not None or rows is not None or replacements is not None:
                    _customize_rendered_table(target_tbl, headers=headers, rows=rows, replacements=replacements)
                    backend.pos = target_tbl.Range.End
                table_info = _extract_table_info(target_tbl)
    except Exception:
        doc.Undo(1)
        raise
    safe_pos = max(0, min(backend.pos, doc.Content.End - 1))
    doc.Range(safe_pos, safe_pos).Select()
    return table_info


KNOWLEDGE_FILE = Path(__file__).resolve().parent / "form_knowledge.json"
_BM25_SEARCHER: "BM25Searcher" | None = None

STOPWORDS = {
    "서식", "문서", "만들어줘", "넣어줘", "생성해줘", "하나", "좀", "있는", "할", "해줘",
    "추가해줘", "부탁해", "작성해줘", "관련", "위한", "으로", "에서", "을", "를", "이", "가", "의", "에",
    "적을", "찍는", "하는", "볼", "수", "있는", "필요해"
}

QUERY_EXPANSIONS = {
    "결재란": ["결재선", "결재", "서명", "도장", "승인"],
    "서명란": ["결재선", "결재", "서명", "직인"],
    "도장": ["결재선", "결재", "서명", "승인"],
    "식순": ["시간계획", "일정표", "타임테이블", "일정"],
    "타임테이블": ["시간계획", "일정표", "시간표"],
    "타임라인": ["시간계획", "추진일정", "연간계획"],
    "주의사항": ["점선박스", "참고", "글상자", "요약"],
    "콜아웃": ["점선박스", "글상자", "참고"],
    "부록": ["붙임", "별첨", "첨부"],
    "별첨": ["붙임", "첨부", "부록"],
    "테이블": ["기본표", "3단표", "5단표", "명단"],
    "남색": ["네이비"],
    "하늘색": ["스카이블루"],
    "파란색": ["스카이블루"],
    "빨간색": ["레드"],
}


def _tokenize(text: str) -> list[str]:
    tokens = re.findall(r"[가-힣a-zA-Z0-9]+", text.lower())
    result = []
    for t in tokens:
        if t in STOPWORDS:
            continue
        result.append(t)
        if len(t) >= 4 and re.match(r"^[가-힣]+$", t):
            result.append(t[:2])
            result.append(t[2:])
    return result


class BM25Searcher:
    """순수 파이썬 구현 Okapi BM25 지능형 서식 검색기"""

    def __init__(self, knowledge_path: Path, k1: float = 1.5, b: float = 0.75):
        self.k1 = k1
        self.b = b
        self.knowledge: dict[str, dict[str, Any]] = {}
        if knowledge_path.exists():
            self.knowledge = json.loads(knowledge_path.read_text(encoding="utf-8"))
        self.doc_count = len(self.knowledge)

        self.doc_tokens: dict[str, Counter[str]] = {}
        self.doc_lens: dict[str, float] = {}
        self.df: dict[str, int] = defaultdict(int)

        total_len = 0.0
        for name, item in self.knowledge.items():
            counter: Counter[str] = Counter()
            for t in _tokenize(name):
                counter[t] += 6.0
            for t in _tokenize(item.get("category", "")):
                counter[t] += 3.0
            for syn in item.get("synonyms", []):
                for t in _tokenize(syn):
                    counter[t] += 2.5
            for lit in item.get("literals", []):
                for t in _tokenize(lit):
                    counter[t] += 1.5
            for t in _tokenize(item.get("description", "")):
                counter[t] += 1.0

            doc_len = sum(counter.values())
            self.doc_tokens[name] = counter
            self.doc_lens[name] = doc_len
            total_len += doc_len

            for token in counter.keys():
                self.df[token] += 1

        self.avgdl = total_len / max(1, self.doc_count)

        self.idf: dict[str, float] = {}
        for token, freq in self.df.items():
            self.idf[token] = math.log((self.doc_count - freq + 0.5) / (freq + 0.5) + 1.0)

    def search(self, query: str, top_k: int = 5) -> list[dict[str, Any]]:
        raw_tokens = _tokenize(query)
        if not raw_tokens:
            return []

        expanded_tokens = list(raw_tokens)
        for t in raw_tokens:
            if t in QUERY_EXPANSIONS:
                expanded_tokens.extend(QUERY_EXPANSIONS[t])

        query_counts = Counter(expanded_tokens)
        scores: dict[str, float] = defaultdict(float)

        for token, q_tf in query_counts.items():
            if token in self.idf:
                idf = self.idf[token]
                for name, doc_counter in self.doc_tokens.items():
                    tf = doc_counter.get(token, 0.0)
                    if tf > 0:
                        doc_len = self.doc_lens[name]
                        numerator = tf * (self.k1 + 1.0)
                        denominator = tf + self.k1 * (1.0 - self.b + self.b * (doc_len / self.avgdl))
                        scores[name] += idf * (numerator / denominator) * q_tf
            else:
                for name, doc_counter in self.doc_tokens.items():
                    for doc_t, tf in doc_counter.items():
                        if token in doc_t or doc_t in token:
                            scores[name] += 1.2 * (tf / (tf + 1.0))

        for name in scores.keys():
            for t in raw_tokens:
                if t in name:
                    scores[name] += 10.0

        if not scores:
            return []

        ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)[:top_k]

        results = []
        for name, score in ranked:
            item = self.knowledge.get(name, {})
            req_args = item.get("required_args", [])
            has_merged = item.get("has_merged_cells", False)
            literals = item.get("literals", [])
            if req_args:
                hint = f"office_insert_form(form_name='{name}', args={req_args})"
            elif has_merged:
                sample_lit = {lit: "변경할값" for lit in literals[:3]} if literals else {"샘플": "값"}
                hint = f"office_insert_form(form_name='{name}', replacements={sample_lit})"
            else:
                hint = f"office_insert_form(form_name='{name}', headers=[...], rows=[[...]])"
            results.append({
                "form_name": name,
                "bm25_score": round(score, 2),
                "category": item.get("category", "기타"),
                "structure": item.get("structure", "문서 서식"),
                "cols": item.get("cols", 0),
                "rows": item.get("rows", 0),
                "has_merged_cells": has_merged,
                "bg_colors": item.get("bg_colors", []),
                "literals": literals[:6],
                "description": item.get("description", ""),
                "required_args": req_args,
                "action_hint": hint,
            })
        return results


def _get_bm25_searcher() -> BM25Searcher:
    global _BM25_SEARCHER
    if _BM25_SEARCHER is None:
        _BM25_SEARCHER = BM25Searcher(KNOWLEDGE_FILE)
    return _BM25_SEARCHER


def register_document_form_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def office_list_forms(
        query: str = "",
        top_k: int = 5,
        doc_type: str = "all",
    ) -> dict[str, Any]:
        """
        [4대 오피스 공용 RAG 지능형 검색] Okapi BM25 알고리즘과 지식 베이스를 바탕으로
        Word, HWP, Excel, PPT 4대 오피스에 적합한 표준 참조 서식을 추천합니다.
        - query: 자연어 검색어 (예: '부서장 승인 결재란', '행사 식순 타임테이블', '주의사항 점선박스', '실적 집계표', '3단 KPI 카드'). 비우면 전체 카테고리 안내
        - top_k: 반환할 추천 서식 개수 (기본 5개)
        - doc_type: 대상 문서 종류 ('all', 'word', 'hwp', 'excel', 'ppt')
        """
        try:
            forms = list_forms()
            searcher = _get_bm25_searcher()
            q_clean = query.strip()
            dt = doc_type.lower().strip()

            # PPT 전용 검색 분기 (slice_roles 및 recipes)
            ppt_recommendations = []
            if dt in ("ppt", "powerpoint", "all"):
                ppt_cat_path = Path(__file__).resolve().parent / "ppt_slice_catalog.json"
                if ppt_cat_path.exists():
                    try:
                        with open(ppt_cat_path, "r", encoding="utf-8") as pf:
                            ppt_cat = json.load(pf)
                        # 1. slice_roles 인덱싱 (kpi_cards, comparison, roadmap 등)
                        sroles = ppt_cat.get("slice_roles", {})
                        for s_id, s_meta in sroles.items():
                            s_name = str(s_meta.get("name", "") or "")
                            s_use = str(s_meta.get("intended_use", "") or "")
                            target_text = f"{s_id} {s_name} {s_use}".lower()
                            if not q_clean or q_clean.lower() in target_text:
                                ppt_recommendations.append({
                                    "form_name": f"ppt_role_{s_id}",
                                    "category": "PPT_역할_아키타입",
                                    "description": s_name or s_use,
                                    "doc_type": "ppt",
                                    "structure": s_id,
                                })

                        # 2. recipes 인덱싱
                        recipes = ppt_cat.get("recipes", {})
                        for r_name, r_meta in recipes.items():
                            desc = str(r_meta.get("description", "") or "")
                            label = str(r_meta.get("name", "") or "")
                            role = str(r_meta.get("role", "") or "")
                            target_text = f"{r_name} {label} {role} {desc}".lower()
                            if not q_clean or q_clean.lower() in target_text:
                                ppt_recommendations.append({
                                    "form_name": f"ppt_{r_name}",
                                    "category": "PPT_슬라이스",
                                    "description": label or desc or r_name,
                                    "doc_type": "ppt",
                                    "structure": role or "slide_layout",
                                })
                    except Exception:
                        pass

            if not q_clean:
                cat_summary: dict[str, list[str]] = defaultdict(list)
                for name, item in searcher.knowledge.items():
                    cat_summary[item.get("category", "기타")].append(name)
                return {
                    "success": True,
                    "mode": "category_overview",
                    "total": len(forms),
                    "total_forms": searcher.doc_count,
                    "matched": len(forms),
                    "forms": forms,
                    "categories": {cat: names[:5] for cat, names in cat_summary.items()},
                    "ppt_slices": [p["form_name"] for p in ppt_recommendations[:5]],
                    "guide": "원하시는 서식의 용도나 특징을 자연어로 검색하세요 (예: query='행사 식순 타임테이블').",
                }

            results = searcher.search(q_clean, top_k=top_k)
            # doc_type 정보 주입
            for r in results:
                r["doc_type"] = "common (word/hwp/excel)"

            matched_forms = {r["form_name"]: r["required_args"] for r in results}
            for name, req_args in forms.items():
                if q_clean in name and name not in matched_forms:
                    matched_forms[name] = req_args

            all_recs = results
            if dt in ("ppt", "all"):
                all_recs = (ppt_recommendations[:top_k] + results)[:top_k*2]

            return {
                "success": True,
                "mode": "bm25_recommendation",
                "query": query,
                "doc_type": doc_type,
                "total": len(forms) + len(ppt_recommendations),
                "matched": len(matched_forms) + len(ppt_recommendations),
                "total_matched": len(all_recs),
                "recommendations": all_recs,
                "forms": matched_forms,
            }
        except Exception as e:
            return {"success": False, "error": f"문서 컴포넌트 서식을 검색하지 못했습니다: {e}"}

    @mcp.tool()
    def office_insert_form(
        form_name: str,
        args: list[Any] | None = None,
        headers: list[str] | None = None,
        rows: list[list[Any]] | None = None,
        replacements: dict[str, str] | None = None,
        doc_type: str = "auto",
    ) -> dict[str, Any]:
        """
        [4대 오피스 공용: 참조 양식 베이스 커스텀 삽입]
        표준 참조 서식을 현재 열려 있는 오피스 창(Word, HWP, Excel, PPT)의 커서/시트/슬라이드 위치에 삽입하고,
        디자인과 레이아웃을 100% 유지한 채 사용자 데이터를 즉시 주입하여 커스텀 양식을 완성합니다.

        - form_name: office_list_forms로 검색한 표준 컴포넌트 이름 (예: '소제목_무테두리_그레이', '시간계획', '결재선_무테두리', '3단표_스카이블루')
        - args: 서식 함수 자체 인자 리스트 (예: 소제목 ['1', '추진 배경'], 점선박스 ['핵심 안내'])
        - headers: 표 헤더 덮어쓸 커스텀 제목 리스트 (예: ['시간', '세션명', '발표자'])
        - rows: 표 데이터 행 2차원 리스트
        - replacements: 고정 문구 1:1 치환 딕셔너리
        - doc_type: 'auto'(현재 활성 창 자동 감지), 'word', 'hwp', 'excel', 'ppt'
        """
        ensure_com()
        target_app = MultiDocFormRenderer.detect_active_target(doc_type)

        # 지식베이스에서 메타데이터 획득
        meta = {}
        knowledge_path = Path(__file__).resolve().parent / "form_knowledge.json"
        if knowledge_path.exists():
            try:
                with open(knowledge_path, "r", encoding="utf-8") as kf:
                    k_data = json.load(kf)
                    clean_name = form_name.replace("ppt_", "")
                    meta = k_data.get(clean_name, {})
            except Exception:
                pass

        # 1. HWP로 라우팅
        if target_app == "hwp":
            return MultiDocFormRenderer.render_to_hwp(
                form_name=form_name,
                meta=meta,
                args=args,
                headers=headers,
                rows=rows,
                replacements=replacements,
            )

        # 2. Excel로 라우팅
        if target_app == "excel":
            return MultiDocFormRenderer.render_to_excel(
                form_name=form_name,
                meta=meta,
                args=args,
                headers=headers,
                rows=rows,
                replacements=replacements,
            )

        # 3. PPT로 라우팅
        if target_app == "ppt":
            return MultiDocFormRenderer.render_to_ppt(
                form_name=form_name,
                meta=meta,
                args=args,
                headers=headers,
                rows=rows,
                replacements=replacements,
            )

        # 4. Word(기본) 렌더링
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "열려 있는 Word 문서를 찾을 수 없습니다 (HWP, Excel, PPT 창도 확인되지 않음)."}
        try:
            if not _recipes():
                if not meta:
                    return {"success": False, "error": f"존재하지 않는 서식 컴포넌트입니다: '{form_name}'"}
                return MultiDocFormRenderer.render_to_word(
                    form_name=form_name,
                    meta=meta,
                    args=args,
                    headers=headers,
                    rows=rows,
                    replacements=replacements,
                )
            recipe_name, err = _resolve_form_name(form_name)
            if err:
                return {"success": False, "error": err}
            reason = _blocker(recipe_name)
            if reason:
                return {"success": False, "error": f"실행할 수 없는 서식입니다: {form_name} ({reason})"}
            table_info = render_form(
                word_app,
                doc,
                recipe_name,
                args=args,
                headers=headers,
                rows=rows,
                replacements=replacements,
            )
        except Exception as e:
            return {"success": False, "error": f"Word 서식 삽입 실패: {e}"}
        com_manager.activate_app("word")
        res: dict[str, Any] = {"success": True, "target_app": "word", "document_name": doc.Name, "form_name": form_name}
        if table_info is not None:
            res["table_info"] = table_info
        return res
