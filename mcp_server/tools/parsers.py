# -*- coding: utf-8 -*-
"""
mcp_server/tools/parsers.py
MCP 서버 내부 자체 마크다운/G서식 계층 문서 파서 및 유틸리티.
레거시 word_helper 의존성을 완전히 제거하고 독립적으로 동작합니다.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any


def _consume_code_block(lines: list[str], i: int) -> tuple[dict[str, str] | None, int]:
    code_lines = []
    first_line = lines[i].strip()
    first_line_content = re.sub(r"^```[a-zA-Z0-9_\-\.]*\s*", "", first_line).strip()
    if first_line_content:
        code_lines.append(first_line_content)

    i += 1
    while i < len(lines) and not lines[i].strip().startswith("```"):
        code_lines.append(lines[i].rstrip())
        i += 1

    if i < len(lines):
        closing_line = lines[i].strip()
        remainder = re.sub(r"^```\s*", "", closing_line).strip()
        i += 1
        if remainder:
            lines.insert(i, remainder)

    elem = {"type": "CODE_BLOCK", "text": "\n".join(code_lines)} if code_lines else None
    return elem, i


def _consume_blockquote(lines: list[str], i: int) -> tuple[dict[str, str] | None, int]:
    quote_lines = []
    while i < len(lines) and lines[i].strip().startswith(">"):
        quote_lines.append(re.sub(r"^>\s?", "", lines[i].strip()))
        i += 1
    elem = {"type": "CODE_BLOCK", "text": "\n".join(quote_lines)} if quote_lines else None
    return elem, i


def _parse_table_row_cells(line: str) -> list[str] | None:
    if re.match(r"^\|[\s\:\-]*(\|[\s\:\-]*)+$", line):
        return None
    cells = []
    for c in line.split("|")[1:-1]:
        clean_cell = re.sub(r"<br\s*/?>", " ", c, flags=re.IGNORECASE)
        clean_cell = re.sub(r"<[^>]+>", "", clean_cell).strip()
        cells.append(clean_cell)
    return cells


def _is_numbered_headline(lines: list[str], current_i: int) -> bool:
    """
    '1. ', '2. ' 형태의 줄이 헤드라인(TITLE_BOX)인지 단순 목록(NORMAL)인지 판별.
    - 직전 또는 직후에 바로 다음 번호가 연속해서 이어지는 경우: 목록 (False)
    - 문장 종결 부호(~다., ~함., 마침표)가 있고 40자 이상인 경우: 목록 (False)
    - 번호 뒤에 하위 불릿(-, *, ○)이나 본문 설명이 뒤따르는 독립 섹션인 경우: 헤드라인 (True)
    """
    import re
    line = lines[current_i].strip()
    m = re.match(r"^(\d+)\.\s+(.*)$", line)
    if not m:
        return False

    num = int(m.group(1))
    content = m.group(2).strip()

    # 긴 문장형 종결이면 목록으로 판정
    ends_with_terminator = bool(re.search(r"(\.|\?|!|[다함음임됨]\.?)$", content))
    if len(content) > 40 and ends_with_terminator:
        return False

    # 이전 유효(non-empty) 줄 탐색
    prev_num = None
    for k in range(current_i - 1, -1, -1):
        prev_l = lines[k].strip()
        if not prev_l:
            continue
        m_prev = re.match(r"^(\d+)\.\s+", prev_l)
        if m_prev:
            prev_num = int(m_prev.group(1))
        break

    # 다음 유효(non-empty) 줄 탐색
    next_num = None
    for k in range(current_i + 1, len(lines)):
        next_l = lines[k].strip()
        if not next_l:
            continue
        m_next = re.match(r"^(\d+)\.\s+", next_l)
        if m_next:
            next_num = int(m_next.group(1))
        break

    # 연속성 체크: 바로 이전 번호가 num-1 이거나 바로 다음 번호가 num+1 이면 연속 목록형
    if prev_num is not None and prev_num == num - 1:
        return False
    if next_num is not None and next_num == num + 1:
        return False

    return True


def parse_document_elements(raw_text: str) -> list[dict[str, Any]]:
    """
    G서식 + 마크다운 통합 4단계 계층 파서:
    - TABLE_ROW     : | 파이프 기호 표 데이터줄
    - TABLE_SEP     : |---|---| 표 구분선
    - TITLE_BOX     : 1단계 주 제목 (# , 1. , 제목:, [제목])
    - SQUARE_HEADER : 2단계 중제목 (## , 1-1., 가., 중제목:, [네모])
    - CIRCLE_ITEM   : 3단계 소제목 (### , ○, ●, ①, 1), 가), 소제목:, [동그라미])
    - DASH_ITEM     : 4단계 세부항목 (#### , - [ ]/- [x], -, *, ※, 세부항목:, 주석:, [바], [주석])
    - CODE_BLOCK    : ``` 코드 블록, > 인용구, 인용:/참고:/안내: 등 콜아웃
    - NORMAL        : 일반 본문
    """
    lines = raw_text.splitlines()
    elements: list[dict[str, Any]] = []
    i = 0

    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        # 0-A. 코드 블록 (```)
        if line.startswith("```"):
            elem, i = _consume_code_block(lines, i)
            if elem:
                elements.append(elem)
            continue

        # 0-A2. 마크다운 인용구 (>)
        if line.startswith(">"):
            elem, i = _consume_blockquote(lines, i)
            if elem:
                elements.append(elem)
            continue

        # 0-A3. 참고/안내/인용 명령어 접두어
        m_callout = re.match(
            r"^(?:\[(?:참고|안내|인용|Callout|callout|주의|경고)\]|(?:참고|안내|인용|Callout)[:：])\s*(.*)$",
            line,
            re.IGNORECASE,
        )
        if m_callout:
            elements.append({"type": "CODE_BLOCK", "text": m_callout.group(1)})
            i += 1
            continue

        # 0-B. 표 데이터줄 (| Header | Content |)
        if line.startswith("|") and line.endswith("|"):
            cells = _parse_table_row_cells(line)
            if cells is None:
                elements.append({"type": "TABLE_SEP"})
                i += 1
                continue
            if cells:
                elements.append({"type": "TABLE_ROW", "cells": cells})
                i += 1
                continue

        # 0-C. 마크다운 수평 구분선 (Horizontal Rule: ---, ***, ___)
        if re.match(r"^(?:-{3,}|\*{3,}|_{3,})$", line):
            elements.append({"type": "HORIZONTAL_RULE"})
            i += 1
            continue

        clean_text = line
        if line.startswith("#"):
            clean_text = re.sub(r"^#+\s*", "", line).strip()

        # 1. 주 제목 (1단계)
        m_title_md = re.match(r"^#\s+(.*)$", line)
        m_title_num = re.match(r"^(\d+\.)\s+(.*)$", line)
        m_title_kw = re.match(r"^(?:\[제목\]|제목|제목상자|【제목】)[:：\s]\s*(.*)$", line, re.IGNORECASE)
        if m_title_kw:
            elements.append({"type": "TITLE_BOX", "text": m_title_kw.group(1)})
            i += 1
            continue
        if m_title_md:
            elements.append({"type": "TITLE_BOX", "text": clean_text})
            i += 1
            continue
        if m_title_num and not re.match(r"^\d+\-\d+", line):
            if _is_numbered_headline(lines, i):
                elements.append({"type": "TITLE_BOX", "text": clean_text})
                i += 1
                continue
            # 연속 목록형 번호(1. 2. 3.)는 TITLE_BOX가 아닌 본문(NORMAL)으로 흐름

        # 2. 소제목/중제목 (2단계)
        m_sub_md = re.match(r"^##\s+(.*)$", line)
        m_sub_num = re.match(r"^(?:\d+\-\d+\.?|[가-하]\.)\s+(.*)$", line)
        m_square_kw = re.match(r"^(?:\[네모\]|네모|네모개요|중제목)[:：\s]\s*(.*)$", line, re.IGNORECASE)
        m_square_sym = re.match(r"^(?:■|□)[:：\s]\s*(.*)$", line)
        if m_square_kw:
            elements.append({"type": "SQUARE_HEADER", "text": m_square_kw.group(1)})
            i += 1
            continue
        if m_sub_md or m_sub_num or m_square_sym:
            elements.append({"type": "SQUARE_HEADER", "text": clean_text})
            i += 1
            continue

        # 3. 세부항목 (4단계)
        m_dash_md = re.match(r"^####\s+(.*)$", line)
        m_chk_un = re.match(r"^[\-\*]\s*\[\s*\]\s*(.*)$", line)
        m_chk_ch = re.match(r"^[\-\*]\s*\[[xX]\]\s*(.*)$", line)
        if m_dash_md:
            elements.append({"type": "DASH_ITEM", "text": clean_text})
            i += 1
            continue
        if m_chk_un:
            elements.append({"type": "DASH_ITEM", "text": f"☐ {m_chk_un.group(1)}"})
            i += 1
            continue
        if m_chk_ch:
            elements.append({"type": "DASH_ITEM", "text": f"☑ {m_chk_ch.group(1)}"})
            i += 1
            continue

        m_dash_kw = re.match(r"^(?:\[바\]|바|세부항목|\[주석\]|주석|참고)[:：\s]\s*(.*)$", line, re.IGNORECASE)
        if m_dash_kw:
            elements.append({"type": "DASH_ITEM", "text": f"- {m_dash_kw.group(1)}"})
            i += 1
            continue

        # 볼드(**)로 감싸진 줄은 불릿 항목으로 오인되지 않도록 가드
        m_dash_sym = re.match(r"^(?:–|-|\*|※)\s*[:：\s]?\s*(.*)$", line)
        if m_dash_sym and not line.startswith("**") and not re.match(r"^-\s*\d+", line):
            elements.append({"type": "DASH_ITEM", "text": clean_text})
            i += 1
            continue

        # 4. 소제목/항목 (3단계)
        m_circle_md = re.match(r"^###\s+(.*)$", line)
        m_circle_num = re.match(r"^(?:[①-⑮]|\d+\)|[가-하]\))\s*(.*)$", line)
        m_circle_kw = re.match(r"^(?:\[동그라미\]|동그라미|소제목)[:：\s]\s*(.*)$", line, re.IGNORECASE)
        m_circle_sym = re.match(r"^(?:○|●)\s*[:：\s]?\s*(.*)$", line)

        if m_circle_kw:
            elements.append({"type": "CIRCLE_ITEM", "text": f"○ {m_circle_kw.group(1)}"})
            i += 1
            continue

        if m_circle_md or m_circle_num or m_circle_sym:
            elements.append({"type": "CIRCLE_ITEM", "text": clean_text})
            i += 1
            continue

        # 5. 일반 본문 문장 (단순 콜론 포함 문장 등은 여기로 안전하게 분기)
        elements.append({"type": "NORMAL", "text": clean_text})
        i += 1

    return elements


def parse_inline_spans(text: str) -> list[dict[str, Any]]:
    """
    텍스트 내 인라인 마크다운(**볼드**, *이탤릭*, `코드`)을 파싱하여
    각 구간(span)별 서식 정보(bold, italic, code)를 분리해 리스트로 반환합니다.
    모든 마크다운 원시 기호는 제거된 상태로 text에 담깁니다.
    """
    if not text:
        return []

    # 정규식 패턴: **bold**, *italic*, `code`
    pattern = re.compile(r"(\*\*.*?\*\*|\*.*?\*|`.*?`)")
    tokens = pattern.split(str(text))
    spans: list[dict[str, Any]] = []

    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) >= 4:
            clean_tok = token[2:-2]
            spans.append({"text": clean_tok, "bold": True, "italic": False, "code": False})
        elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
            clean_tok = token[1:-1]
            spans.append({"text": clean_tok, "bold": False, "italic": True, "code": False})
        elif token.startswith("`") and token.endswith("`") and len(token) >= 2:
            clean_tok = token[1:-1]
            spans.append({"text": clean_tok, "bold": False, "italic": False, "code": True})
        else:
            # 고립된 잔여 별표/언더스코어 제거
            clean_tok = token.replace("**", "").replace("__", "")
            spans.append({"text": clean_tok, "bold": False, "italic": False, "code": False})

    return spans


def calculate_korean_age(rrn: str | Any, base_date: datetime | None = None) -> str:
    """
    주민등록번호 13자리 텍스트 데이터로부터 만 나이를 자동 산출 (KISA 범피스 1310 호환)
    """
    if not rrn:
        return ""
    clean_id = str(rrn).replace("-", "").strip()
    if len(clean_id) < 7 or not clean_id[:6].isdigit():
        return ""

    try:
        birth_y = int(clean_id[:2])
        birth_m = int(clean_id[2:4])
        birth_d = int(clean_id[4:6])
        gender_code = clean_id[6]

        # 2000년대 이후 출생자 구분 (3, 4, 7, 8)
        if gender_code in ("3", "4", "7", "8"):
            full_year = 2000 + birth_y
        else:
            full_year = 1900 + birth_y

        birth_date = datetime(full_year, birth_m, birth_d)
    except Exception:
        return ""

    today = base_date or datetime.today()
    age = today.year - birth_date.year
    if (today.month, today.day) < (birth_date.month, birth_date.day):
        age -= 1
    return str(age)


def strip_markdown_syntax(text: str) -> str:
    """
    텍스트 내의 모든 마크다운 서식 기호(**, *, __, ~~, `, #, 불릿 등)를 완전히 제거하여
    오피스(Word, Excel, HWP) 문서에 '**' 같은 문법 기호가 노출되지 않도록 정제합니다.
    """
    if not text:
        return ""
    # [텍스트](url) -> 텍스트
    s = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', str(text))
    # **굵게** 또는 __굵게__ -> 굵게
    s = re.sub(r'\*\*(.*?)\*\*', r'\1', s)
    s = re.sub(r'(?<!\w)__([^_\n]+?)__(?!\w)', r'\1', s)
    # *기울임* 또는 _기울임_ -> 기울임
    s = re.sub(r'\*(.*?)\*', r'\1', s)
    s = re.sub(r'(?<!\w)_([^_\n]+?)_(?!\w)', r'\1', s)
    # ~~취소선~~ -> 취소선
    s = re.sub(r'~~(.*?)~~', r'\1', s)
    # `인라인코드` -> 인라인코드
    s = re.sub(r'`(.*?)`', r'\1', s)
    # 줄 시작의 #, ##, ### 등 마크다운 헤더 기호 제거
    s = re.sub(r'(?m)^#{1,6}\s*', '', s)
    # - [ ] 또는 - [x] 체크박스 마크다운 정제
    s = re.sub(r'-\s*\[[ xX]\]\s*', '• ', s)
    s = re.sub(r'\[[ xX]\]\s*', '', s)
    # 남아있는 고립된 ** 또는 __ 별표 제거
    s = s.replace("**", "").replace("__", "")
    return s
