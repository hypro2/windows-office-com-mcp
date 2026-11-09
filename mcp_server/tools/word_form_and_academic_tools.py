# -*- coding: utf-8 -*-
"""
mcp_server/tools/word_form_and_academic_tools.py
iOfficeAI(OfficeCLI) officecli-word-form & officecli-academic-paper 규격 기반 Word 실시간 COM 도구:
1. 리얼 인터랙티브 폼 (SDT ContentControls + 양식 보호 wdAllowOnlyFormFields + 밑줄 안티패턴 감사)
2. 학술 논문 조판 (APA7 / Chicago17 / IEEE / MLA9 프리셋 + 2단 분단 + SEQ/PAGEREF + 참고문헌 내어쓰기)
"""

from __future__ import annotations

import re
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, word_undo_group
from .parsers import strip_markdown_syntax

WD_NO_PROTECTION = -1
WD_ALLOW_ONLY_FORM_FIELDS = 2

WD_CONTENT_CONTROL_TYPES: dict[str, int] = {
    "richtext": 0,  # wdContentControlRichText
    "text": 1,      # wdContentControlText
    "picture": 2,   # wdContentControlPicture
    "combobox": 3,  # wdContentControlComboBox
    "dropdown": 4,  # wdContentControlDropdownList
    "date": 6,      # wdContentControlDate
    "group": 7,     # wdContentControlGroup
    "checkbox": 8,  # wdContentControlCheckBox
}

_UNDERSCORE_FAKE_FIELD_RE = re.compile(r"_{4,}")
_REF_HEADER_RE = re.compile(
    r"^\s*(?:\d+\.|[ivxlcdm]+\.)?\s*(?:references|bibliography|works\s+cited|참고문헌)(?:\s*[\(:].*)?$",
    re.IGNORECASE,
)

ACADEMIC_VENUE_SPECS: dict[str, dict[str, Any]] = {
    "APA7": {
        "name": "APA 7th Edition",
        "font_name": "Times New Roman",
        "body_size_pt": 12.0,
        "line_spacing_mult": 2.0,
        "first_line_indent_pt": 36.0,  # 0.5 inch (720 twips)
        "margin_pt": 72.0,
        "h1_align": "center",
        "h1_numbered": False,
        "two_column_default": False,
        "reference_section_title": "References",
        "in_text_citation": "(Author, Year)",
    },
    "Chicago17": {
        "name": "Chicago 17th Edition (Notes-Bibliography)",
        "font_name": "Times New Roman",
        "body_size_pt": 12.0,
        "line_spacing_mult": 1.5,
        "first_line_indent_pt": 36.0,
        "margin_pt": 72.0,
        "h1_align": "left",
        "h1_numbered": True,
        "two_column_default": False,
        "reference_section_title": "Bibliography",
        "in_text_citation": "Superscript footnote",
    },
    "IEEE": {
        "name": "IEEE Transactions / Conference",
        "font_name": "Times New Roman",
        "body_size_pt": 10.0,
        "line_spacing_mult": 1.15,
        "first_line_indent_pt": 14.4,  # 0.2 inch (288 twips)
        "margin_pt": 72.0,
        "h1_align": "center",
        "h1_numbered": True,
        "h1_roman_caps": True,
        "two_column_default": True,
        "reference_section_title": "References",
        "in_text_citation": "[1]",
    },
    "MLA9": {
        "name": "MLA 9th Edition",
        "font_name": "Times New Roman",
        "body_size_pt": 12.0,
        "line_spacing_mult": 2.0,
        "first_line_indent_pt": 36.0,
        "margin_pt": 72.0,
        "h1_align": "left",
        "h1_numbered": False,
        "two_column_default": False,
        "reference_section_title": "Works Cited",
        "in_text_citation": "(Author Page)",
    },
    "KCI": {
        "name": "KCI Korean Journal Standard",
        "font_name": "맑은 고딕",
        "body_size_pt": 10.5,
        "line_spacing_mult": 1.5,
        "first_line_indent_pt": 10.0,
        "margin_pt": 56.7,
        "h1_align": "left",
        "h1_numbered": True,
        "two_column_default": False,
        "reference_section_title": "참고문헌",
        "in_text_citation": "(저자, 연도)",
    },
}


def get_academic_venue_spec(venue: str) -> dict[str, Any]:
    """학술 스타일 이름 정규화 후 규격 딕셔너리 반환."""
    key = (venue or "APA7").strip().lower().replace("-", "").replace(" ", "")
    if key in ("apa", "apa7", "apa7th"):
        return ACADEMIC_VENUE_SPECS["APA7"]
    if key in ("chicago", "chicago17", "chicago17th"):
        return ACADEMIC_VENUE_SPECS["Chicago17"]
    if key in ("ieee", "ieeetrans"):
        return ACADEMIC_VENUE_SPECS["IEEE"]
    if key in ("mla", "mla9", "mla9th"):
        return ACADEMIC_VENUE_SPECS["MLA9"]
    if key in ("kci", "kcijournal", "korean"):
        return ACADEMIC_VENUE_SPECS["KCI"]
    return ACADEMIC_VENUE_SPECS["APA7"]


def roman_numeral(n: int) -> str:
    """IEEE 섹션/표 번호용 로마자 대문자 변환 (1 -> 'I')."""
    if n <= 0:
        return str(n)
    numerals = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    out: list[str] = []
    rem = int(n)
    for val, sym in numerals:
        while rem >= val:
            out.append(sym)
            rem -= val
    return "".join(out)


def scan_form_antipatterns(
    paragraphs: list[str],
    controls: list[dict[str, Any]],
    protection_type: int | None,
    legacy_formfields: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Word 양식 필드 무결성 및 밑줄(`____`) 가짜 입력란 안티패턴 순수 로직 검사."""
    underscore_leaks: list[dict[str, Any]] = []
    for idx, p_text in enumerate(paragraphs, start=1):
        clean = str(p_text or "").strip()
        if _UNDERSCORE_FAKE_FIELD_RE.search(clean):
            underscore_leaks.append(
                {
                    "paragraph_index": idx,
                    "snippet": clean[:80],
                    "message": "밑줄(`____`) 가짜 입력칸이 감지되었습니다. `word_insert_sdt_control`로 실제 SDT 컨트롤을 삽입하세요.",
                }
            )

    control_issues: list[dict[str, Any]] = []
    for c_idx, ctrl in enumerate(controls, start=1):
        title = str(ctrl.get("title", "")).strip()
        tag = str(ctrl.get("tag", "")).strip()
        c_type = ctrl.get("type")
        entries_count = int(ctrl.get("entries_count", 0))
        date_fmt = str(ctrl.get("date_format", "")).strip()

        if not title:
            control_issues.append({"index": c_idx, "code": "missing_alias", "message": f"SDT #{c_idx}에 Title(alias)이 누락되었습니다."})
        if not tag:
            control_issues.append({"index": c_idx, "code": "missing_tag", "message": f"SDT #{c_idx}에 Tag가 누락되었습니다."})
        if c_type in (3, 4, "combobox", "dropdown") and entries_count == 0:
            control_issues.append({"index": c_idx, "code": "empty_dropdown", "message": f"SDT #{c_idx}('{title or tag}') 드롭다운/콤보박스 항목이 비어 있습니다."})
        if c_type in (6, "date") and not date_fmt:
            control_issues.append({"index": c_idx, "code": "missing_date_format", "message": f"날짜 SDT #{c_idx}에 DateDisplayFormat이 없습니다."})

    for f_idx, ff in enumerate(legacy_formfields or [], start=1):
        fname = str(ff.get("name", ""))
        if len(fname) > 20:
            control_issues.append({"index": f_idx, "code": "formfield_name_too_long", "message": f"FormField '{fname}' 이름이 20자를 초과했습니다 ({len(fname)}자)."})

    total_fields = len(controls) + len(legacy_formfields or [])
    protection_enforced = protection_type == WD_ALLOW_ONLY_FORM_FIELDS
    if total_fields > 0 and not protection_enforced:
        control_issues.append(
            {
                "code": "missing_form_protection",
                "message": "양식 컨트롤은 존재하나 양식 보호(wdAllowOnlyFormFields=2)가 활성화되지 않았습니다.",
            }
        )

    passed = len(underscore_leaks) == 0 and len(control_issues) == 0
    suggestions: list[str] = []
    if underscore_leaks:
        suggestions.append("밑줄(`____`) 구간을 지우고 `word_insert_sdt_control`로 실제 SDT 필드를 삽입하세요.")
    if not protection_enforced and total_fields > 0:
        suggestions.append("모든 정적 레이아웃 완성 후 `word_set_form_protection(enable=True)`를 호출해 양식 외 본문을 잠그세요.")

    return {
        "passed": passed,
        "total_controls": len(controls),
        "legacy_formfields_count": len(legacy_formfields or []),
        "protection_type": protection_type,
        "protection_enforced": protection_enforced,
        "underscore_leaks_count": len(underscore_leaks),
        "underscore_leaks": underscore_leaks,
        "control_issues": control_issues,
        "suggestions": suggestions,
    }


def register_word_form_academic_tools(mcp: MCPServer) -> None:
    """Word SDT 양식 컨트롤 및 학술 논문 조판 MCP 도구 등록."""

    @mcp.tool()
    def word_insert_sdt_control(
        control_type: Literal["text", "richtext", "dropdown", "combobox", "date", "checkbox", "picture", "group"] = "text",
        alias: str = "",
        tag: str = "",
        label: str = "",
        default_text: str = "",
        items: list[str] | None = None,
        date_format: str = "yyyy-MM-dd",
        lock_control: bool = True,
    ) -> dict[str, Any]:
        """
        현재 Word 커서 위치에 밑줄(`___`) 대신 실제 채움형 콘텐츠 컨트롤(SDT ContentControl)을 삽입합니다.

        Args:
            control_type: "text", "richtext", "dropdown", "combobox", "date", "checkbox", "picture", "group"
            alias: 사용자에게 표시되는 필드 제목 (`Title`)
            tag: 데이터 추출/바인딩용 식별 태그 (`Tag`)
            label: 컨트롤 앞에 함께 입력할 레이블 텍스트 (예: "성명: ")
            default_text: 초기 안내 텍스트
            items: dropdown/combobox 선택 항목 리스트
            date_format: 날짜 컨트롤 표시 형식 (기본 "yyyy-MM-dd")
            lock_control: True이면 사용자가 실수로 컨트롤을 삭제하지 못하도록 잠금 (`LockContentControl`)
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=True, visible=True)
        if word_app is None or doc is None:
            return {"success": False, "error": "Word 연결 실패"}

        if getattr(doc, "ProtectionType", WD_NO_PROTECTION) != WD_NO_PROTECTION:
            return {
                "success": False,
                "error": "문서에 보호가 적용되어 있어 새 SDT 컨트롤을 삽입할 수 없습니다.",
                "suggestion": "먼저 `word_set_form_protection(enable=False)`로 보호를 해제한 뒤 컨트롤을 추가하고 마지막에 다시 잠그세요.",
            }

        cc_type = WD_CONTENT_CONTROL_TYPES.get(control_type.lower().strip(), 1)
        clean_tag = tag.strip() or re.sub(r"[^a-zA-Z0-9_가-힣]+", "_", alias.strip()).strip("_") or f"field_{cc_type}"
        clean_alias = alias.strip() or clean_tag

        try:
            with word_undo_group(word_app, f"MCP SDT 삽입 ({clean_alias})"):
                sel = word_app.Selection
                if sel.Start != sel.End:
                    sel.Collapse(0)
                in_table = False
                try:
                    in_table = bool(sel.Information(12))
                except Exception:
                    pass
                try:
                    sel.ParagraphFormat.FirstLineIndent = 0.0
                except Exception:
                    pass
                # 직전 제목/강조 색상·크기가 SDT에 전이되지 않도록 표준 가독성 타이포로 정규화
                try:
                    norm_font = "맑은 고딕"
                    norm_size = 10.5
                    try:
                        nst = doc.Styles(-1)
                        cand_fn = str(getattr(nst.Font, "NameFarEast", "") or getattr(nst.Font, "Name", "") or "").strip()
                        if cand_fn:
                            norm_font = cand_fn
                        cand_sz = float(getattr(nst.Font, "Size", 10.5) or 10.5)
                        if 9.5 <= cand_sz <= 12.0:
                            norm_size = cand_sz
                    except Exception:
                        pass
                    sel.Font.Name = norm_font
                    try:
                        sel.Font.NameFarEast = norm_font
                    except Exception:
                        pass
                    sel.Font.Size = norm_size
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                except Exception:
                    pass
                if label:
                    try:
                        sel.Font.Bold = -1
                        sel.Font.Color = 0x3B291E  # #1E293B (BGR)
                    except Exception:
                        pass
                    sel.TypeText(strip_markdown_syntax(label))
                    if not label.endswith(" "):
                        sel.TypeText(" ")
                try:
                    sel.Font.Bold = 0
                    sel.Font.Color = 0x37291F  # #1F2937 (BGR)
                except Exception:
                    pass

                cc = doc.ContentControls.Add(cc_type, sel.Range)
                cc.Title = clean_alias
                cc.Tag = clean_tag
                cc.LockContentControl = bool(lock_control)

                if cc_type in (3, 4) and items:
                    for opt in items:
                        opt_str = strip_markdown_syntax(str(opt))
                        cc.DropdownListEntries.Add(opt_str, opt_str)

                if cc_type == 6 and date_format:
                    cc.DateDisplayFormat = date_format

                if default_text and cc_type not in (2, 8):
                    cc.Range.Text = strip_markdown_syntax(default_text)

                # 커서를 컨트롤 바깥 뒤쪽으로 이동해 연속 입력 시 중첩 방지 (표 셀 경계 안전 가드)
                try:
                    after_rng = cc.Range.Duplicate
                    after_rng.Collapse(0)
                    if not in_table:
                        after_rng.Move(1, 1)  # wdCharacter = 1
                    after_rng.Select()
                except Exception:
                    pass

            com_manager.activate_app("word")
            return {
                "success": True,
                "control_type": control_type,
                "wd_type": cc_type,
                "alias": clean_alias,
                "tag": clean_tag,
                "items_added": len(items) if items and cc_type in (3, 4) else 0,
                "locked": bool(lock_control),
            }
        except Exception as e:
            return {"success": False, "error": f"SDT 컨트롤 삽입 실패: {e}"}

    @mcp.tool()
    def word_set_form_protection(
        enable: bool = True,
        password: str = "",
    ) -> dict[str, Any]:
        """
        Word 양식 문서의 정적 레이아웃 보호(`wdAllowOnlyFormFields=2`)를 켜거나 끕니다.
        보호를 켜면 일반 본문은 읽기 전용으로 잠기고 SDT/FormField 영역만 입력 가능해집니다.
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "열려 있는 Word 문서가 없습니다."}

        try:
            cur_prot = getattr(doc, "ProtectionType", WD_NO_PROTECTION)
            if enable:
                if cur_prot != WD_NO_PROTECTION:
                    doc.Unprotect(password)
                doc.Protect(WD_ALLOW_ONLY_FORM_FIELDS, True, password)
                state = "protected"
            else:
                if cur_prot != WD_NO_PROTECTION:
                    doc.Unprotect(password)
                state = "unprotected"

            com_manager.activate_app("word")
            return {
                "success": True,
                "document_name": doc.Name,
                "state": state,
                "protection_type": getattr(doc, "ProtectionType", WD_NO_PROTECTION),
            }
        except Exception as e:
            return {"success": False, "error": f"양식 보호 설정 실패: {e}"}

    @mcp.tool()
    def word_audit_form_fields() -> dict[str, Any]:
        """
        현재 열린 Word 문서 내 모든 SDT 콘텐츠 컨트롤/FormField 목록, alias/tag 누락,
        빈 드롭다운, 본문 및 표 내 밑줄(`____`) 가짜 입력칸 안티패턴, 문서 보호 여부를 감사(Audit)합니다.
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if word_app is None or doc is None:
            return {"success": False, "error": "열려 있는 Word 문서가 없습니다."}

        try:
            controls: list[dict[str, Any]] = []
            for i in range(1, doc.ContentControls.Count + 1):
                cc = doc.ContentControls(i)
                cc_type = int(getattr(cc, "Type", 1))
                entries_cnt = 0
                if cc_type in (3, 4):
                    try:
                        entries_cnt = int(cc.DropdownListEntries.Count)
                    except Exception:
                        pass
                date_fmt = ""
                if cc_type == 6:
                    try:
                        date_fmt = str(cc.DateDisplayFormat or "")
                    except Exception:
                        pass
                text_prev = ""
                try:
                    text_prev = str(cc.Range.Text or "")[:60]
                except Exception:
                    pass
                controls.append(
                    {
                        "index": i,
                        "title": str(getattr(cc, "Title", "") or ""),
                        "tag": str(getattr(cc, "Tag", "") or ""),
                        "type": cc_type,
                        "entries_count": entries_cnt,
                        "date_format": date_fmt,
                        "locked": bool(getattr(cc, "LockContentControl", False)),
                        "text_preview": text_prev,
                    }
                )

            legacy_ff: list[dict[str, Any]] = []
            try:
                for i in range(1, doc.FormFields.Count + 1):
                    ff = doc.FormFields(i)
                    legacy_ff.append({"index": i, "name": str(getattr(ff, "Name", "") or ""), "type": getattr(ff, "Type", None)})
            except Exception:
                pass

            paragraphs: list[str] = []
            for i in range(1, min(doc.Paragraphs.Count + 1, 800)):
                try:
                    paragraphs.append(
                        str(doc.Paragraphs(i).Range.Text or "").replace("\r", "").replace("\x07", "").strip()
                    )
                except Exception:
                    continue
            try:
                for sh_i in range(1, getattr(doc.Shapes, "Count", 0) + 1):
                    sh = doc.Shapes(sh_i)
                    if getattr(sh, "TextFrame", None) and getattr(sh.TextFrame, "HasText", 0):
                        for ln in str(sh.TextFrame.TextRange.Text or "").replace("\r", "\n").splitlines():
                            if ln.strip():
                                paragraphs.append(ln.strip())
            except Exception:
                pass

            report = scan_form_antipatterns(
                paragraphs=paragraphs,
                controls=controls,
                protection_type=getattr(doc, "ProtectionType", WD_NO_PROTECTION),
                legacy_formfields=legacy_ff,
            )
            report["success"] = True
            report["document_name"] = doc.Name
            report["controls"] = controls
            return report
        except Exception as e:
            return {"success": False, "error": f"양식 필드 감사 실패: {e}"}

    @mcp.tool()
    def word_apply_academic_venue_style(
        venue: Literal["APA7", "Chicago17", "IEEE", "MLA9", "KCI"] = "APA7",
        two_column_body: bool | None = None,
        bibliography_hanging_indent: bool = True,
        fareast_font: str = "맑은 고딕",
        apply_first_line_indent: bool = True,
        insert_toc: bool = False,
        ensure_page_numbers: bool = False,
    ) -> dict[str, Any]:
        """
        활성 Word 문서에 학술 논문 투고 규격(APA 7, Chicago 17, IEEE, MLA 9, KCI 국문학술지)을 적용합니다:
        - 학회별 표준 여백·서체·줄간격(2.0x / 1.5x / 1.15x) 및 양쪽 정렬(Justify)
        - Heading 1~3 위계 정비, 표/캡션 첫 줄 들여쓰기 침범 방지
        - IEEE 2단 조판(`two_column_body=True`) 및 참고문헌/Bibliography 문단 내어쓰기(0.5인치=36pt)
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=True, visible=True)
        if word_app is None or doc is None:
            return {"success": False, "error": "Word 연결 실패"}

        spec = get_academic_venue_spec(venue)
        use_two_col = spec["two_column_default"] if two_column_body is None else bool(two_column_body)
        font_name = spec["font_name"]
        body_pt = float(spec["body_size_pt"])
        line_mult = float(spec["line_spacing_mult"])

        try:
            with word_undo_group(word_app, f"MCP 학술 스타일 ({spec['name']})"):
                # 1. 페이지 여백 (1인치 = 72pt)
                for s_i in range(1, doc.Sections.Count + 1):
                    ps = doc.Sections(s_i).PageSetup
                    ps.TopMargin = spec["margin_pt"]
                    ps.BottomMargin = spec["margin_pt"]
                    ps.LeftMargin = spec["margin_pt"]
                    ps.RightMargin = spec["margin_pt"]

                # 2. Normal / Heading 1~3 스타일 정비 (다국어 로캘 호환: WdBuiltinStyle -1,-2,-3,-4 우선)
                def _get_style(*candidates: Any) -> Any:
                    for cand in candidates:
                        try:
                            return doc.Styles(cand)
                        except Exception:
                            continue
                    return None

                normal_style = _get_style(-1, "Normal", "표준")
                if normal_style is not None:
                    normal_style.Font.Name = font_name
                    normal_style.Font.Size = body_pt
                    try:
                        normal_style.Font.NameAscii = font_name
                        normal_style.Font.NameOther = font_name
                        if fareast_font.strip():
                            normal_style.Font.NameFarEast = fareast_font.strip()
                    except Exception:
                        pass
                    normal_pf = normal_style.ParagraphFormat
                    normal_pf.Alignment = 3  # wdAlignParagraphJustify (양쪽 정렬)
                    normal_pf.LineSpacingRule = 5  # wdLineSpaceMultiple
                    try:
                        normal_pf.LineSpacing = float(word_app.LinesToPoints(line_mult))
                    except Exception:
                        normal_pf.LineSpacing = body_pt * line_mult
                    try:
                        normal_pf.FarEastLineBreakControl = -1
                        normal_pf.WordWrap = -1
                        normal_pf.SpaceAfter = 4.0 if use_two_col else 5.0
                    except Exception:
                        pass
                    if apply_first_line_indent:
                        try:
                            normal_pf.FirstLineIndent = float(spec.get("first_line_indent_pt", 0.0))
                        except Exception:
                            pass
                        # 기존 Word 표/콜아웃 셀 및 Figure/Table 캡션은 첫 줄 들여쓰기 제외
                        try:
                            for t_i in range(1, doc.Tables.Count + 1):
                                doc.Tables(t_i).Range.ParagraphFormat.FirstLineIndent = 0.0
                        except Exception:
                            pass
                        try:
                            for p_i in range(1, doc.Paragraphs.Count + 1):
                                dp = doc.Paragraphs(p_i)
                                dtxt = str(dp.Range.Text or "").lstrip()
                                if dtxt.startswith(("Figure ", "Table ")):
                                    dp.Range.ParagraphFormat.FirstLineIndent = 0.0
                                    dp.Range.ParagraphFormat.Alignment = 0
                                    dp.Range.Font.Size = max(8.5, body_pt - 1.0)
                                    dp.Range.Font.Bold = -1
                                    dp.Range.ParagraphFormat.SpaceBefore = 4.0
                                    dp.Range.ParagraphFormat.SpaceAfter = 4.0
                        except Exception:
                            pass

                h1_align = 1 if spec["h1_align"] == "center" else 0
                if use_two_col or body_pt <= 10.5:
                    h_sizes = (12.0, 10.5, 10.0)
                else:
                    h_sizes = (18.0, 14.0, 12.0)
                heading_specs = [
                    ((-2, "Heading 1", "제목 1"), h_sizes[0], True, False, h1_align),
                    ((-3, "Heading 2", "제목 2"), h_sizes[1], True, False, 0),
                    ((-4, "Heading 3", "제목 3"), h_sizes[2], True, True, 0),
                ]
                for cands, st_pt, is_b, is_i, st_align in heading_specs:
                    st = _get_style(*cands)
                    if st is None:
                        continue
                    try:
                        st.Font.Name = font_name
                        try:
                            st.Font.NameAscii = font_name
                            st.Font.NameOther = font_name
                            if fareast_font.strip():
                                st.Font.NameFarEast = fareast_font.strip()
                        except Exception:
                            pass
                        st.Font.Size = st_pt
                        st.Font.Bold = -1 if is_b else 0
                        st.Font.Italic = -1 if is_i else 0
                        st.ParagraphFormat.Alignment = st_align
                        st.ParagraphFormat.FirstLineIndent = 0.0
                        st.ParagraphFormat.SpaceBefore = 8.0 if use_two_col else 12.0
                        st.ParagraphFormat.SpaceAfter = 4.0 if use_two_col else 6.0
                    except Exception:
                        continue

                # 3. 2단 조판 적용 + 기존 수식(Equation) 탭 정지점 2단 폭 재보정 + 캡션/초록 들여쓰기 보호
                if use_two_col:
                    target_sec = doc.Sections(doc.Sections.Count)
                    try:
                        target_sec.PageSetup.TextColumns.SetCount(2)
                        target_sec.PageSetup.TextColumns.Spacing = 18.0
                    except Exception:
                        pass
                    try:
                        ps2 = target_sec.PageSetup
                        col_w = (
                            float(ps2.PageWidth)
                            - float(ps2.LeftMargin)
                            - float(ps2.RightMargin)
                            - float(ps2.TextColumns.Spacing)
                        ) / 2.0
                        if col_w > 80.0:
                            for p_i in range(1, target_sec.Range.Paragraphs.Count + 1):
                                tp = target_sec.Range.Paragraphs(p_i)
                                raw_ptxt = str(tp.Range.Text or "")
                                if raw_ptxt.startswith("\t") and raw_ptxt.count("\t") >= 2:
                                    tpf = tp.Range.ParagraphFormat
                                    tpf.FirstLineIndent = 0.0
                                    tpf.TabStops.ClearAll()
                                    tpf.TabStops.Add(col_w / 2.0, 1)
                                    tpf.TabStops.Add(col_w, 2)
                                elif raw_ptxt.startswith(("Figure ", "Table ")):
                                    tp.Range.ParagraphFormat.FirstLineIndent = 0.0
                    except Exception:
                        pass
                    if doc.Sections.Count >= 2:
                        try:
                            for p_i in range(1, doc.Sections(1).Range.Paragraphs.Count + 1):
                                p1 = doc.Sections(1).Range.Paragraphs(p_i)
                                p1_txt = str(p1.Range.Text or "").strip()
                                if p1_txt.lower().startswith("abstract") or p1.Range.ParagraphFormat.Alignment == 1:
                                    p1.Range.ParagraphFormat.FirstLineIndent = 0.0
                        except Exception:
                            pass

                # 4. 참고문헌 섹션 이후 내어쓰기 일괄 적용 (2단은 18pt, 1단은 36pt)
                hang_pt = 18.0 if use_two_col else 36.0
                hanging_applied = 0
                if bibliography_hanging_indent:
                    in_refs = False
                    for p_i in range(1, doc.Paragraphs.Count + 1):
                        p = doc.Paragraphs(p_i)
                        txt = (p.Range.Text or "").replace("\r", "").replace("\x07", "").strip()
                        if _REF_HEADER_RE.match(txt):
                            p.Range.ParagraphFormat.FirstLineIndent = 0.0
                            in_refs = True
                            continue
                        if in_refs and txt:
                            p.Range.ParagraphFormat.LeftIndent = hang_pt
                            p.Range.ParagraphFormat.FirstLineIndent = -hang_pt
                            hanging_applied += 1

                # 5. 바닥글 페이지 번호(PAGE) 및 목차(TOC) 선택적 삽입/갱신
                if ensure_page_numbers:
                    try:
                        last_sec = doc.Sections(doc.Sections.Count)
                        last_sec.Footers(1).PageNumbers.Add(2, True)  # wdAlignPageNumberRight = 2
                    except Exception:
                        pass

                if insert_toc:
                    try:
                        if doc.TablesOfContents.Count > 0:
                            doc.TablesOfContents(1).Update()
                        else:
                            if doc.Sections.Count >= 2:
                                sec1_end = max(0, int(doc.Sections(1).Range.End) - 1)
                                anchor_rng = doc.Range(sec1_end, sec1_end)
                            else:
                                anchor_rng = doc.Range(0, 0)
                            anchor_rng.InsertBefore("Contents\r")
                            toc_title_rng = anchor_rng.Duplicate
                            toc_pos = anchor_rng.Duplicate
                            toc_pos.Collapse(0)

                            toc_title_rng.Font.Name = font_name
                            if fareast_font.strip():
                                try:
                                    toc_title_rng.Font.NameFarEast = fareast_font.strip()
                                except Exception:
                                    pass
                            toc_title_rng.Font.Size = max(10.5, body_pt)
                            toc_title_rng.Font.Bold = -1
                            toc_title_rng.ParagraphFormat.Alignment = 0
                            toc_title_rng.ParagraphFormat.FirstLineIndent = 0.0
                            toc_title_rng.ParagraphFormat.SpaceBefore = 6.0
                            toc_title_rng.ParagraphFormat.SpaceAfter = 2.0

                            toc_obj = doc.TablesOfContents.Add(
                                Range=toc_pos,
                                UseHeadingStyles=True,
                                UpperHeadingLevel=1,
                                LowerHeadingLevel=3,
                            )
                            try:
                                toc_obj.Range.ParagraphFormat.FirstLineIndent = 0.0
                                toc_obj.Range.Font.Size = min(body_pt, 9.5) if use_two_col else body_pt
                            except Exception:
                                pass
                    except Exception:
                        pass

            com_manager.activate_app("word")
            return {
                "success": True,
                "document_name": doc.Name,
                "venue": spec["name"],
                "font_name": font_name,
                "body_size_pt": body_pt,
                "line_spacing_mult": line_mult,
                "two_column_body": use_two_col,
                "hanging_indent_paragraphs": hanging_applied,
            }
        except Exception as e:
            return {"success": False, "error": f"학술 스타일 적용 실패: {e}"}

    @mcp.tool()
    def word_insert_academic_caption_or_ref(
        element_type: Literal["Figure", "Table", "Equation", "Reference"] = "Figure",
        caption_text: str = "",
        bookmark_name: str = "",
        insert_pageref_at_selection: bool = False,
    ) -> dict[str, Any]:
        """
        학술 논문용 동적 필드(`SEQ Figure`, `SEQ Table`, `SEQ Equation`, `PAGEREF`) 또는
        내어쓰기가 적용된 참고문헌(`Reference`) 문단을 삽입합니다.
        (원칙: Figure 캡션은 그림 하단, Table 캡션은 표 상단 배치)
        """
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=True, visible=True)
        if word_app is None or doc is None:
            return {"success": False, "error": "Word 연결 실패"}

        try:
            with word_undo_group(word_app, f"MCP 학술 요소 ({element_type})"):
                sel = word_app.Selection
                if sel.Start != sel.End:
                    sel.Collapse(0)

                if insert_pageref_at_selection:
                    if not bookmark_name:
                        return {"success": False, "error": "PAGEREF 삽입 시 bookmark_name이 필요합니다."}
                    fld = doc.Fields.Add(sel.Range, -1, f"PAGEREF {bookmark_name} \\h", True)
                    try:
                        fld.Update()
                    except Exception:
                        pass
                    com_manager.activate_app("word")
                    return {"success": True, "action": "pageref_inserted", "bookmark": bookmark_name, "result": str(getattr(fld.Result, "Text", ""))}

                if element_type == "Reference":
                    pf = sel.ParagraphFormat
                    pf.LeftIndent = 36.0
                    pf.FirstLineIndent = -36.0
                    sel.TypeText(strip_markdown_syntax(caption_text))
                    sel.TypeParagraph()
                    try:
                        sel.ParagraphFormat.LeftIndent = 0.0
                        sel.ParagraphFormat.FirstLineIndent = 0.0
                    except Exception:
                        pass
                    com_manager.activate_app("word")
                    return {"success": True, "action": "reference_inserted", "hanging_indent_pt": 36.0}

                if element_type == "Equation":
                    # 단일 줄 2단 탭(중앙 수식 + 우측 끝 번호) 규격
                    try:
                        sec_idx = max(1, doc.Sections.Count)
                        ps = doc.Sections(sec_idx).PageSetup
                        text_w = float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)
                        if int(ps.TextColumns.Count) >= 2:
                            text_w = (text_w - float(ps.TextColumns.Spacing)) / float(ps.TextColumns.Count)
                    except Exception:
                        text_w = 220.0
                    try:
                        sel.ParagraphFormat.FirstLineIndent = 0.0
                    except Exception:
                        pass
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(text_w / 2.0, 1)  # wdAlignTabCenter = 1
                    sel.ParagraphFormat.TabStops.Add(text_w, 2)        # wdAlignTabRight = 2
                    sel.TypeText(f"\t{strip_markdown_syntax(caption_text)}\t(")
                    fld = doc.Fields.Add(sel.Range, -1, "SEQ Equation", True)
                    try:
                        fld.Update()
                    except Exception:
                        pass
                    sel.Collapse(0)
                    sel.TypeText(")")
                    if bookmark_name:
                        doc.Bookmarks.Add(bookmark_name, sel.Paragraphs(1).Range)
                    sel.TypeParagraph()
                    try:
                        sel.ParagraphFormat.TabStops.ClearAll()
                    except Exception:
                        pass
                    com_manager.activate_app("word")
                    return {"success": True, "action": "equation_inserted", "bookmark": bookmark_name, "seq_result": str(getattr(fld.Result, "Text", ""))}

                # Figure / Table SEQ 캡션
                try:
                    sel.ParagraphFormat.FirstLineIndent = 0.0
                except Exception:
                    pass
                start_pos = sel.Start
                sel.TypeText(f"{element_type} ")
                fld = doc.Fields.Add(sel.Range, -1, f"SEQ {element_type}", True)
                try:
                    fld.Update()
                except Exception:
                    pass
                sel.Collapse(0)
                if caption_text:
                    sel.TypeText(f": {strip_markdown_syntax(caption_text)}")
                end_pos = sel.End

                if bookmark_name:
                    bm_rng = doc.Range(start_pos, end_pos)
                    doc.Bookmarks.Add(bookmark_name, bm_rng)
                sel.TypeParagraph()

            com_manager.activate_app("word")
            placement = "below_figure" if element_type == "Figure" else "above_table"
            return {
                "success": True,
                "action": "caption_inserted",
                "element_type": element_type,
                "placement_rule": placement,
                "bookmark": bookmark_name,
                "seq_result": str(getattr(fld.Result, "Text", "")),
            }
        except Exception as e:
            return {"success": False, "error": f"학술 캡션/참조 삽입 실패: {e}"}
