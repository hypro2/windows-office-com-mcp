# -*- coding: utf-8 -*-
"""
mcp_server/tools/delivery_gate_and_merge_tools.py
iOfficeAI(AionUi · OfficeCLI) 납품 품질 5-Gate 감사기, 전사 {{key}} 템플릿 Merge,
Mark-and-Review 비파괴 마킹 및 시각 스냅샷(Render -> See -> Fix) 브릿지 도구.
"""

from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr, word_undo_group
from .excel_modeling_tools import FinancialModelClassifier, _col_letter
from .excel_tools import _normalize_com_value, excel_freeze_display
from .ppt_morph_engine import MorphAuditor, _read_presentation_morph_snapshot, classify_morph_shape_prefix
from .ppt_tools import ppt_freeze_display
from .spatial_auditor import BoundingBox, SpatialAuditor
from .word_authenticity_gate import (
    compare_real_vs_generated,
    evaluate_docx_authenticity,
    extract_docx_style_profile,
    is_neon_low_contrast_fill,
    profile_from_word_snapshot,
)
from .word_form_and_academic_tools import WD_ALLOW_ONLY_FORM_FIELDS, WD_NO_PROTECTION, scan_form_antipatterns

# Excel COM CVErr 정수 -> 에러 문자열 매핑
EXCEL_CVERR_MAP: dict[int, str] = {
    -2146826288: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
}
EXCEL_ERROR_STRINGS = tuple(EXCEL_CVERR_MAP.values())

_MUSTACHE_RE = re.compile(r"\{\{([^{}\n]{1,60})\}\}")
_SINGLE_BRACE_VAR_RE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]{1,40})\}")
_TODO_SENTINEL_RE = re.compile(
    r"(?:<TODO>|TODO:|FIXME:|\[작성\s*필요\]|\[확인\s*필요\]|lorem\s+ipsum|\bxxxx+\b)",
    re.IGNORECASE,
)
_SHELL_DOLLAR_RE = re.compile(r"\\\$")
_WIN_PATH_RE = re.compile(r"[A-Za-z]:\\[^\s]+")
_LITERAL_ESC_RE = re.compile(r"\\(?:n|t)(?![A-Za-z0-9_]*\\)")
_MD_BOLD_RE = re.compile(r"\*\*[^*\n]{1,80}\*\*")
_MD_HEAD_RE = re.compile(r"(?:^|\n)#{1,6}\s+\S+")
_MD_CHECK_RE = re.compile(r"(?:^|\n)[-*]\s*\[[ xX]\]")
_UNDERSCORE_FIELD_RE = re.compile(r"_{4,}")


def bgr_to_rgb(bgr: int) -> tuple[int, int, int]:
    """MS Office COM COLORREF (BGR 정수) -> (R, G, B) 변환."""
    val = int(bgr) & 0xFFFFFF
    return (val & 0xFF, (val >> 8) & 0xFF, (val >> 16) & 0xFF)


def compute_luminance(rgb: tuple[int, int, int]) -> float:
    """RGB (0..255) 상대 밝기(0.0~1.0) 계산."""
    r, g, b = rgb
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def check_low_contrast(bg_rgb: tuple[int, int, int], text_rgb: tuple[int, int, int]) -> bool:
    """어두운 배경 위 어두운 글씨 및 밝은 배경 위 밝은 글씨 저대비 여부 판별."""
    bg_lum = compute_luminance(bg_rgb)
    text_lum = compute_luminance(text_rgb)
    return (bg_lum < 0.30 and text_lum < 0.50) or (bg_lum > 0.80 and text_lum > 0.75)


class DeliveryGateScanner:
    """4대 오피스 납품 전 5단계 품질 게이트(Gate 1~5) 순수 로직 스캐너."""

    @staticmethod
    def scan_text_leaks(text: str) -> list[dict[str, str]]:
        if not text:
            return []
        leaks: list[dict[str, str]] = []
        txt = str(text)

        for m in _MUSTACHE_RE.findall(txt):
            leaks.append({"code": "unreplaced_token", "token": f"{{{{{m}}}}}", "message": f"미치환 템플릿 변수 '{{{{{m}}}}}' 발견"})

        # {{...}}를 제외한 영역에서 단일 중괄호 {var} 탐색 (API 경로/수식 오탐 방지를 위해 warning 등급)
        txt_no_mustache = _MUSTACHE_RE.sub("", txt)
        for m in _SINGLE_BRACE_VAR_RE.findall(txt_no_mustache):
            leaks.append(
                {
                    "code": "unreplaced_brace_token",
                    "severity": "warning",
                    "token": f"{{{m}}}",
                    "message": f"미치환 변수 추정 토큰 '{{{m}}}' 발견 (코드/수식이 아니라면 병합 확인 필요)",
                }
            )

        for m in _TODO_SENTINEL_RE.findall(txt):
            leaks.append({"code": "placeholder_sentinel", "token": m, "message": f"임시 플레이스홀더 '{m}' 잔존"})

        if _SHELL_DOLLAR_RE.search(txt):
            leaks.append({"code": "shell_escape_dollar", "token": r"\$", "message": r"쉘 이스케이프 잔재 `\$` 발견"})

        # Windows 파일 경로 제외 후 리터럴 \n, \t 탐지
        txt_no_paths = _WIN_PATH_RE.sub("", txt)
        if _LITERAL_ESC_RE.search(txt_no_paths):
            leaks.append({"code": "literal_escape_sequence", "token": r"\n/\t", "message": r"리터럴 이스케이프 문자열(`\n` 또는 `\t`) 노출"})

        if _MD_BOLD_RE.search(txt) or _MD_HEAD_RE.search(txt) or _MD_CHECK_RE.search(txt):
            leaks.append({"code": "markdown_syntax_leak", "token": "** / ##", "message": "오피스 본문에 마크다운 서식 기호(**, ##, [ ])가 그대로 노출됨"})

        return leaks

    @staticmethod
    def scan_excel_cell(cell_ref: str, display_text: Any, formula: Any, value: Any) -> list[dict[str, Any]]:
        findings: list[dict[str, Any]] = []
        disp = str(display_text or "").strip()

        if disp.startswith("###"):
            findings.append(
                {
                    "gate": "gate_2",
                    "cell": cell_ref,
                    "code": "excel_cell_truncated",
                    "message": f"셀 {cell_ref}의 열 너비가 부족하여 '{disp}'(###)로 잘렸습니다.",
                }
            )

        val_err = EXCEL_CVERR_MAP.get(value) if isinstance(value, int) else None
        err_hit = val_err or next((e for e in EXCEL_ERROR_STRINGS if e in disp or e in str(value or "")), None)
        if err_hit:
            findings.append(
                {
                    "gate": "gate_2",
                    "cell": cell_ref,
                    "code": "excel_formula_error",
                    "token": err_hit,
                    "message": f"셀 {cell_ref}에 Excel 수식 오류({err_hit})가 있습니다.",
                }
            )

        up_disp = disp.upper()
        if "IMBALANCED" in up_disp or "CF !=" in up_disp:
            findings.append(
                {
                    "gate": "gate_4",
                    "cell": cell_ref,
                    "code": "financial_imbalance",
                    "message": f"셀 {cell_ref}에서 재무제표 불균형({disp})이 감지되었습니다.",
                }
            )
        return findings

    @classmethod
    def evaluate(cls, doc_type: str, snapshot: dict[str, Any]) -> dict[str, Any]:
        kind = doc_type.lower().strip()
        if kind in ("powerpoint", "pptx"):
            kind = "ppt"
        elif kind in ("docx", "doc"):
            kind = "word"
        elif kind in ("xlsx", "xls", "sheet"):
            kind = "excel"
        elif kind in ("hwpx", "hancom"):
            kind = "hwp"

        gates: dict[str, dict[str, Any]] = {
            f"gate_{i}": {"passed": True, "issues": [], "warnings": []} for i in range(1, 6)
        }
        suggestions: list[str] = []

        # --- Gate 1: 구조 / 스키마 ---
        if kind == "ppt":
            slides = snapshot.get("slides", [])
            if not slides:
                gates["gate_1"]["issues"].append({"code": "empty_presentation", "message": "슬라이드가 1장도 없습니다."})
            for sl in slides:
                idx = sl.get("slide_index", 0)
                if int(sl.get("shapes_count", 0)) == 0:
                    gates["gate_1"]["issues"].append({"slide": idx, "code": "empty_slide", "message": f"슬라이드 {idx}가 비어 있습니다."})
                elif not str(sl.get("all_text", "")).strip() and not sl.get("has_media_or_chart", False):
                    gates["gate_1"]["issues"].append({"slide": idx, "code": "blank_content_slide", "message": f"슬라이드 {idx}에 텍스트/차트/표 콘텐츠가 없습니다."})
                for sp_issue in sl.get("spatial_warnings", []):
                    gates["gate_1"]["warnings"].append({"slide": idx, **sp_issue})
        elif kind in ("word", "hwp"):
            paras = [str(p or "") for p in snapshot.get("paragraphs", [])]
            if not any(p.strip() for p in paras) and not snapshot.get("tables_count", 0):
                gates["gate_1"]["issues"].append({"code": "empty_document", "message": "문서 본문이 비어 있습니다."})
            paras_for_empty = [str(p or "") for p in (snapshot.get("body_paragraphs") or paras)]
            consec_empty = 0
            for p_i, p in enumerate(paras_for_empty, start=1):
                if not p.strip():
                    consec_empty += 1
                    if consec_empty == 3:
                        gates["gate_1"]["issues"].append({"paragraph": p_i, "code": "consecutive_empty_paragraphs", "message": f"문단 {p_i} 부근에 빈 줄이 3개 이상 연속 누적되었습니다."})
                else:
                    consec_empty = 0
        elif kind == "excel":
            sheets = snapshot.get("sheets", [])
            if not sheets:
                gates["gate_1"]["issues"].append({"code": "empty_workbook", "message": "표시된 시트가 없습니다."})
            elif all(not s.get("has_content", False) for s in sheets):
                gates["gate_1"]["issues"].append({"code": "empty_sheets", "message": "모든 시트가 비어 있습니다."})

        # --- Gate 2: 토큰 & 이스케이프 누수 + Excel 수식/### 오류 ---
        for blk in snapshot.get("text_blocks", []):
            loc = blk.get("location", "unknown")
            for leak in cls.scan_text_leaks(blk.get("text", "")):
                if leak.get("severity") == "warning":
                    gates["gate_2"]["warnings"].append({"location": loc, **leak})
                else:
                    gates["gate_2"]["issues"].append({"location": loc, **leak})

        if kind == "excel":
            for sh in snapshot.get("sheets", []):
                s_name = sh.get("name", "Sheet")
                for cell_info in sh.get("cells", []):
                    for f in cls.scan_excel_cell(
                        f"{s_name}!{cell_info.get('cell', '')}",
                        cell_info.get("display", ""),
                        cell_info.get("formula", ""),
                        cell_info.get("value", ""),
                    ):
                        g_key = f.pop("gate")
                        gates[g_key]["issues"].append(f)

        # --- Gate 3: 필수 구성 요소 ---
        if kind == "ppt":
            slides = snapshot.get("slides", [])
            for sl in slides:
                idx = int(sl.get("slide_index", 0))
                if idx >= 2 and not str(sl.get("notes_text", "")).strip():
                    gates["gate_3"]["issues"].append({"slide": idx, "code": "missing_speaker_notes", "message": f"콘텐츠 슬라이드 {idx}에 발표자 노트(Speaker Notes)가 누락되었습니다."})
        elif kind == "word":
            h_count = int(snapshot.get("heading_count", 0))
            if h_count >= 3:
                if not snapshot.get("has_toc", False):
                    gates["gate_3"]["issues"].append({"code": "missing_toc", "message": f"제목이 {h_count}개 이상인 문서에 목차(TOC)가 없습니다."})
                if not snapshot.get("has_page_field", False):
                    gates["gate_3"]["issues"].append({"code": "missing_page_numbers", "message": "머리글/바닥글에 동적 페이지 번호(PAGE)가 없습니다."})
        elif kind == "excel":
            for sh in snapshot.get("sheets", []):
                s_name = str(sh.get("name", ""))
                if FinancialModelClassifier.classify_zone(s_name) == "Outputs":
                    if not sh.get("print_area"):
                        gates["gate_3"]["warnings"].append({"sheet": s_name, "code": "missing_print_area", "message": f"Outputs 시트 '{s_name}'에 Print_Area가 설정되지 않았습니다."})

        # --- Gate 4: 시각 대비 & 수리 무결성 ---
        for cp in snapshot.get("contrast_pairs", []):
            bg = tuple(cp.get("bg_rgb", (255, 255, 255)))
            fg = tuple(cp.get("text_rgb", (0, 0, 0)))
            if check_low_contrast(bg, fg):
                kind_msg = "어두운 배경 위 어두운 글씨" if compute_luminance(bg) < 0.5 else "밝은 배경 위 밝은 글씨"
                gates["gate_4"]["issues"].append(
                    {
                        "location": cp.get("location", ""),
                        "code": "low_contrast",
                        "bg_rgb": bg,
                        "text_rgb": fg,
                        "message": f"{cp.get('location', '요소')}에서 {kind_msg} 저대비(배경{bg}, 글씨{fg})가 탐지되었습니다.",
                    }
                )

        if kind == "excel":
            for hv in snapshot.get("calc_hardcodes", []):
                gates["gate_4"]["issues"].append({"code": "calc_hardcode", **hv})

        # --- Gate 5: 도메인 특화 게이트 (Morph / Word Forms) ---
        if kind == "ppt":
            morph_slides = snapshot.get("morph_slides", [])
            has_morph = any(
                s.get("transition_effect") in (3954, "morph")
                or any(classify_morph_shape_prefix(sh.get("name", "")) != "other" for sh in s.get("shapes", []))
                for s in morph_slides
            )
            if has_morph and len(morph_slides) >= 2:
                m_res = MorphAuditor.audit_deck(morph_slides)
                for iss in m_res.get("issues", []):
                    gates["gate_5"]["issues"].append(iss)
                for warn in m_res.get("warnings", []):
                    gates["gate_5"]["warnings"].append(warn)
        elif kind == "word":
            controls = snapshot.get("controls", [])
            paras = [str(p or "") for p in snapshot.get("paragraphs", [])]
            has_underscore = any(_UNDERSCORE_FIELD_RE.search(p) for p in paras)
            if controls or has_underscore:
                f_res = scan_form_antipatterns(
                    paragraphs=paras,
                    controls=controls,
                    protection_type=snapshot.get("protection_type"),
                    legacy_formfields=snapshot.get("legacy_formfields"),
                )
                for leak in f_res.get("underscore_leaks", []):
                    gates["gate_5"]["issues"].append({"code": "underscore_fake_field", **leak})
                for ci in f_res.get("control_issues", []):
                    gates["gate_5"]["issues"].append(ci)

        total_issues = 0
        total_warnings = 0
        authenticity_report: dict[str, Any] | None = None
        if kind == "word":
            gates["gate_6"] = {"passed": True, "issues": [], "warnings": []}
            auth_prof = profile_from_word_snapshot(snapshot)
            authenticity_report = evaluate_docx_authenticity(auth_prof, archetype="auto")
            for tell in authenticity_report.get("ai_tells", []):
                target_bucket = "issues" if tell.get("severity") == "critical" else "warnings"
                gates["gate_6"][target_bucket].append(tell)

        for g_name, g_val in gates.items():
            g_val["passed"] = len(g_val["issues"]) == 0
            total_issues += len(g_val["issues"])
            total_warnings += len(g_val["warnings"])

        if not gates["gate_2"]["passed"]:
            suggestions.append("미치환 `{{key}}` 토큰은 `office_merge_template_tokens`로 병합하고, Excel `###` 셀은 열 너비를 늘리세요.")
        if not gates["gate_4"]["passed"]:
            suggestions.append("저대비 도형의 글자 색상을 `#FFFFFF`로 밝게 조정하거나, Excel Calc 시트 하드코딩 숫자를 `Assumptions`로 이동하세요.")
        if not gates["gate_5"]["passed"]:
            suggestions.append("PPT Morph는 `ppt_clone_and_morph_slide`로 고스팅하고, Word 밑줄은 `word_insert_sdt_control`로 교체하세요.")
        if "gate_6" in gates and not gates["gate_6"]["passed"]:
            suggestions.append("AI 합성 징후(이모지/네온 저대비 음영/미설치 폰트/공백 가짜 들여쓰기)를 실무 KISA·비즈니스 표준 서식으로 정제하세요.")

        res_dict: dict[str, Any] = {
            "status": "PASS" if total_issues == 0 else "FAIL",
            "doc_type": kind,
            "total_issues": total_issues,
            "total_warnings": total_warnings,
            "gates": gates,
            "suggestions": suggestions,
        }
        if authenticity_report is not None:
            res_dict["authenticity"] = authenticity_report
        return res_dict


def _resolve_live_app(target_app: str) -> tuple[str, Any, Any]:
    """요청된 앱 또는 현재 열린 첫 번째 활성 오피스 앱 COM 객체를 반환."""
    req = (target_app or "auto").lower().strip()
    if req in ("word", "docx", "doc"):
        app, doc = com_manager.get_word(create_if_missing=False)
        return ("word", app, doc)
    if req in ("excel", "xlsx", "xls"):
        app, wb, _ = com_manager.get_excel(create_if_missing=False)
        return ("excel", app, wb)
    if req in ("ppt", "powerpoint", "pptx"):
        app, prs = com_manager.get_ppt(create_if_missing=False)
        return ("ppt", app, prs)
    if req in ("hwp", "hwpx", "hancom"):
        hwp = com_manager.get_hwp(create_if_missing=False)
        return ("hwp", hwp, hwp)

    for probe in ("ppt", "word", "excel", "hwp"):
        kind, app, doc = _resolve_live_app(probe)
        if app is not None and doc is not None:
            return (kind, app, doc)
    return ("none", None, None)


def _scan_ppt_shape(
    sh: Any,
    s_idx: int,
    slide_w: float,
    slide_bg_rgb: tuple[int, int, int],
    opaque_rects: list[tuple[float, float, float, float, tuple[int, int, int]]],
    texts: list[str],
    text_blocks: list[dict[str, str]],
    contrast_pairs: list[dict[str, Any]],
    spatial_boxes: list[BoundingBox],
    in_group: bool = False,
) -> bool:
    """PowerPoint 도형(그룹/표/차트 제목 포함)을 재귀 순회하며 스냅샷 메타를 수집."""
    has_media = False
    sh_name = str(getattr(sh, "Name", "Shape"))

    try:
        if getattr(sh, "Type", 0) == 6:  # msoGroup
            for g_i in range(1, sh.GroupItems.Count + 1):
                if _scan_ppt_shape(
                    sh.GroupItems(g_i),
                    s_idx,
                    slide_w,
                    slide_bg_rgb,
                    opaque_rects,
                    texts,
                    text_blocks,
                    contrast_pairs,
                    spatial_boxes,
                    in_group=True,
                ):
                    has_media = True
            return has_media
    except Exception:
        pass

    sh_l = float(getattr(sh, "Left", 0.0))
    sh_t = float(getattr(sh, "Top", 0.0))
    sh_w = float(getattr(sh, "Width", 0.0))
    sh_h = float(getattr(sh, "Height", 0.0))

    if getattr(sh, "HasChart", 0) or getattr(sh, "HasTable", 0) or getattr(sh, "Type", 0) == 13:
        has_media = True

    if getattr(sh, "HasChart", 0):
        try:
            if getattr(sh.Chart, "HasTitle", False):
                ct_txt = str(sh.Chart.ChartTitle.Text or "").replace("\r", " ").strip()
                if ct_txt:
                    texts.append(ct_txt)
                    text_blocks.append({"location": f"slide[{s_idx}]/{sh_name}/chart_title", "text": ct_txt})
        except Exception:
            pass

    own_bg_rgb: tuple[int, int, int] | None = None
    try:
        if getattr(sh.Fill, "Visible", 0) and getattr(sh.Fill, "Type", 0) == 1:
            own_bg_rgb = bgr_to_rgb(sh.Fill.ForeColor.RGB)
    except Exception:
        pass

    if getattr(sh, "HasTextFrame", 0) and getattr(sh.TextFrame, "HasText", 0):
        tr = sh.TextFrame.TextRange
        t = (tr.Text or "").replace("\r", "\n").strip()
        if t:
            texts.append(t)
            text_blocks.append({"location": f"slide[{s_idx}]/{sh_name}", "text": t})
            if sh_l < slide_w - 5.0:
                eff_bg = own_bg_rgb
                if eff_bg is None:
                    cx, cy = sh_l + sh_w / 2.0, sh_t + sh_h / 2.0
                    eff_bg = slide_bg_rgb
                    for rl, rt, rw, rh, r_rgb in reversed(opaque_rects):
                        if rl <= cx <= rl + rw and rt <= cy <= rt + rh:
                            eff_bg = r_rgb
                            break
                try:
                    fg_rgb = bgr_to_rgb(tr.Font.Color.RGB)
                    contrast_pairs.append(
                        {"location": f"slide[{s_idx}]/{sh_name}", "bg_rgb": eff_bg, "text_rgb": fg_rgb}
                    )
                except Exception:
                    pass

                if not in_group and classify_morph_shape_prefix(sh_name) != "scene" and sh_w > 0 and sh_h > 0:
                    try:
                        f_sz = float(tr.Font.Size)
                    except Exception:
                        f_sz = 14.0
                    spatial_boxes.append(
                        BoundingBox(id=sh_name, left=sh_l, top=sh_t, width=sh_w, height=sh_h, text=t, font_size=f_sz)
                    )

    if own_bg_rgb is not None and sh_l < slide_w - 5.0:
        opaque_rects.append((sh_l, sh_t, sh_w, sh_h, own_bg_rgb))

    if getattr(sh, "HasTable", 0):
        tbl = sh.Table
        for r in range(1, tbl.Rows.Count + 1):
            for c in range(1, tbl.Columns.Count + 1):
                ct = (tbl.Cell(r, c).Shape.TextFrame.TextRange.Text or "").replace("\r", " ").strip()
                if ct:
                    texts.append(ct)
                    text_blocks.append({"location": f"slide[{s_idx}]/{sh_name}[{r},{c}]", "text": ct})

    return has_media


def _extract_ppt_snapshot(prs: Any) -> dict[str, Any]:
    morph_slides = _read_presentation_morph_snapshot(prs)
    slides_meta: list[dict[str, Any]] = []
    text_blocks: list[dict[str, str]] = []
    contrast_pairs: list[dict[str, Any]] = []
    slide_w = float(getattr(prs.PageSetup, "SlideWidth", 960.0))

    for s_idx in range(1, prs.Slides.Count + 1):
        sl = prs.Slides(s_idx)
        texts: list[str] = []
        has_media = False
        opaque_rects: list[tuple[float, float, float, float, tuple[int, int, int]]] = []
        spatial_boxes: list[BoundingBox] = []

        # 슬라이드 기본 배경색 조회 (단색 채움 없으면 흰색 캔버스 간주)
        slide_bg_rgb: tuple[int, int, int] = (255, 255, 255)
        try:
            if getattr(sl.Background.Fill, "Type", 0) == 1 and getattr(sl.Background.Fill, "Visible", 0):
                slide_bg_rgb = bgr_to_rgb(sl.Background.Fill.ForeColor.RGB)
        except Exception:
            pass

        for sh_i in range(1, sl.Shapes.Count + 1):
            try:
                sh = sl.Shapes(sh_i)
                if _scan_ppt_shape(
                    sh,
                    s_idx,
                    slide_w,
                    slide_bg_rgb,
                    opaque_rects,
                    texts,
                    text_blocks,
                    contrast_pairs,
                    spatial_boxes,
                ):
                    has_media = True
            except Exception:
                continue

        spatial_warnings: list[dict[str, Any]] = []
        if spatial_boxes:
            try:
                sp_res = SpatialAuditor.audit_slide_elements(spatial_boxes, slide_index=s_idx)
                for iss in sp_res.get("issues", []):
                    if iss.get("code") in ("overlap", "text_overflow"):
                        spatial_warnings.append(iss)
            except Exception:
                pass

        notes_txt = ""
        try:
            for n_i in range(1, sl.NotesPage.Shapes.Count + 1):
                nsh = sl.NotesPage.Shapes(n_i)
                if getattr(nsh, "HasTextFrame", 0) and getattr(nsh.TextFrame, "HasText", 0):
                    nt = (nsh.TextFrame.TextRange.Text or "").strip()
                    if nt and not nt.isdigit():
                        notes_txt += nt + " "
        except Exception:
            pass

        slides_meta.append(
            {
                "slide_index": s_idx,
                "shapes_count": sl.Shapes.Count,
                "all_text": "\n".join(texts),
                "has_media_or_chart": has_media,
                "notes_text": notes_txt.strip(),
                "spatial_warnings": spatial_warnings,
            }
        )

    return {
        "slides": slides_meta,
        "text_blocks": text_blocks,
        "contrast_pairs": contrast_pairs,
        "morph_slides": morph_slides,
    }


def _extract_word_snapshot(doc: Any) -> dict[str, Any]:
    paragraphs: list[str] = []
    raw_paragraphs: list[str] = []
    body_paragraphs: list[str] = []
    text_blocks: list[dict[str, str]] = []
    fonts_seen: set[str] = set()
    fills_seen: set[str] = set()
    heading_count = 0
    dark_runs = 0
    color_samples = 0

    try:
        nst_fn = str(getattr(doc.Styles(-1).Font, "NameFarEast", "") or getattr(doc.Styles(-1).Font, "Name", "") or "").strip()
        if nst_fn:
            fonts_seen.add(nst_fn)
    except Exception:
        pass

    for p_i in range(1, min(doc.Paragraphs.Count + 1, 1000)):
        try:
            p = doc.Paragraphs(p_i)
            raw_txt = str(p.Range.Text or "")
            raw_no_cr = raw_txt.replace("\r", "").replace("\x07", "")
            raw_paragraphs.append(raw_no_cr)
            txt = raw_no_cr.strip()
            paragraphs.append(txt)
            in_table = "\x07" in raw_txt
            if not in_table:
                try:
                    in_table = bool(p.Range.Information(12))  # wdWithInTable
                except Exception:
                    pass
            if txt or not in_table:
                body_paragraphs.append(txt)
            if txt:
                text_blocks.append({"location": f"p[{p_i}]", "text": txt})
                if p_i <= 60:
                    try:
                        fn = str(getattr(p.Range.Font, "NameFarEast", "") or getattr(p.Range.Font, "Name", "") or "").strip()
                        if fn:
                            fonts_seen.add(fn)
                    except Exception:
                        pass
                    if not in_table:
                        try:
                            fc = int(p.Range.Font.Color)
                            if fc in (-16777216, -1, 9999999):
                                color_samples += 1
                                dark_runs += 1
                            else:
                                color_samples += 1
                                if compute_luminance(bgr_to_rgb(fc)) < 0.45:
                                    dark_runs += 1
                        except Exception:
                            pass
            try:
                st_name = str(p.Style.NameLocal or "")
                if "Heading" in st_name or "제목" in st_name:
                    heading_count += 1
            except Exception:
                pass
        except Exception:
            continue

    bordered_tables_count = 0
    padded_tables_count = 0
    neon_low_contrast_count = 0
    tables_total = int(getattr(doc.Tables, "Count", 0))
    for t_i in range(1, min(tables_total + 1, 20)):
        try:
            tbl = doc.Tables(t_i)
            if any(getattr(tbl.Borders(b_id), "LineStyle", 0) != 0 for b_id in (-1, -2, -3, -5)):
                bordered_tables_count += 1
            pad_sum = float(getattr(tbl, "TopPadding", 0.0) or 0.0) + float(getattr(tbl, "LeftPadding", 0.0) or 0.0)
            if pad_sum <= 0.5:
                try:
                    c11 = tbl.Cell(1, 1)
                    pad_sum = float(getattr(c11, "TopPadding", 0.0) or 0.0) + float(getattr(c11, "LeftPadding", 0.0) or 0.0)
                except Exception:
                    pass
            if pad_sum > 0.5:
                padded_tables_count += 1
            for r_i in range(1, min(int(tbl.Rows.Count) + 1, 4)):
                for c_i in range(1, min(int(tbl.Columns.Count) + 1, 7)):
                    try:
                        cell = tbl.Cell(r_i, c_i)
                        bg_bgr = int(cell.Shading.BackgroundPatternColor)
                        if bg_bgr not in (-16777216, -1, 9999999):
                            br, bg_g, bb = bgr_to_rgb(bg_bgr)
                            fill_hex = f"{br:02X}{bg_g:02X}{bb:02X}"
                            fills_seen.add(fill_hex)
                            fg_bgr = int(cell.Range.Font.Color)
                            if fg_bgr not in (-16777216, -1, 9999999):
                                fr, fg_g, fb = bgr_to_rgb(fg_bgr)
                                fg_hex = f"{fr:02X}{fg_g:02X}{fb:02X}"
                            else:
                                fg_hex = "000000"
                            if is_neon_low_contrast_fill(fill_hex, fg_hex):
                                neon_low_contrast_count += 1
                    except Exception:
                        continue
        except Exception:
            continue

    margins_pt: dict[str, float] = {}
    try:
        ps = doc.Sections(1).PageSetup
        margins_pt = {
            "top": round(float(ps.TopMargin), 1),
            "left": round(float(ps.LeftMargin), 1),
            "right": round(float(ps.RightMargin), 1),
            "bottom": round(float(ps.BottomMargin), 1),
        }
    except Exception:
        pass

    try:
        for sh_i in range(1, getattr(doc.Shapes, "Count", 0) + 1):
            sh = doc.Shapes(sh_i)
            if getattr(sh, "TextFrame", None) and getattr(sh.TextFrame, "HasText", 0):
                sh_txt = str(sh.TextFrame.TextRange.Text or "").replace("\r", "\n").strip()
                if sh_txt:
                    text_blocks.append({"location": f"shape[{sh_i}]", "text": sh_txt})
                    for ln in sh_txt.splitlines():
                        if ln.strip():
                            paragraphs.append(ln.strip())
                            raw_paragraphs.append(ln)
    except Exception:
        pass

    has_toc = False
    try:
        has_toc = doc.TablesOfContents.Count > 0
    except Exception:
        pass

    has_page_field = False
    has_hf_text = False
    try:
        for s_i in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(s_i)
            for hf_coll in (sec.Headers, sec.Footers):
                for idx in (1, 2, 3):
                    try:
                        rng = hf_coll(idx).Range
                        if str(rng.Text or "").replace("\r", "").strip():
                            has_hf_text = True
                        for f_i in range(1, rng.Fields.Count + 1):
                            if "PAGE" in str(rng.Fields(f_i).Code.Text or "").upper():
                                has_page_field = True
                                break
                    except Exception:
                        continue
            try:
                if sec.Footers(1).PageNumbers.Count > 0 or sec.Headers(1).PageNumbers.Count > 0:
                    has_page_field = True
            except Exception:
                pass
    except Exception:
        pass

    controls: list[dict[str, Any]] = []
    try:
        for i in range(1, doc.ContentControls.Count + 1):
            cc = doc.ContentControls(i)
            cc_type = int(getattr(cc, "Type", 1))
            ec = 0
            if cc_type in (3, 4):
                try:
                    ec = int(cc.DropdownListEntries.Count)
                except Exception:
                    pass
            date_fmt = ""
            if cc_type == 6:
                try:
                    date_fmt = str(cc.DateDisplayFormat or "")
                except Exception:
                    pass
            cc_in_tbl = False
            cc_fsize = 10.5
            try:
                cc_in_tbl = bool(cc.Range.Information(12))
            except Exception:
                pass
            try:
                sz = float(cc.Range.Font.Size)
                if 5.0 <= sz <= 72.0:
                    cc_fsize = sz
            except Exception:
                pass
            controls.append(
                {
                    "title": str(getattr(cc, "Title", "") or ""),
                    "tag": str(getattr(cc, "Tag", "") or ""),
                    "type": cc_type,
                    "entries_count": ec,
                    "date_format": date_fmt,
                    "in_table": cc_in_tbl,
                    "font_size": cc_fsize,
                }
            )
    except Exception:
        pass

    return {
        "paragraphs": paragraphs,
        "raw_paragraphs": raw_paragraphs,
        "body_paragraphs": body_paragraphs,
        "tables_count": tables_total,
        "bordered_tables_count": bordered_tables_count,
        "padded_tables_count": padded_tables_count,
        "text_blocks": text_blocks,
        "fonts": sorted(fonts_seen),
        "fills": sorted(fills_seen),
        "neon_low_contrast_count": neon_low_contrast_count,
        "dark_body_ratio": round(dark_runs / max(color_samples, 1), 3) if color_samples else 1.0,
        "margins_pt": margins_pt,
        "heading_count": heading_count,
        "has_toc": has_toc,
        "has_page_field": has_page_field,
        "has_header_footer": has_page_field or has_hf_text,
        "controls": controls,
        "protection_type": getattr(doc, "ProtectionType", WD_NO_PROTECTION),
        "contrast_pairs": [],
    }


def _extract_excel_snapshot(wb: Any) -> dict[str, Any]:
    sheets_meta: list[dict[str, Any]] = []
    text_blocks: list[dict[str, str]] = []
    calc_hardcodes: list[dict[str, Any]] = []
    contrast_pairs: list[dict[str, Any]] = []

    for i in range(1, wb.Sheets.Count + 1):
        sh = wb.Sheets(i)
        s_name = str(sh.Name)
        used = sh.UsedRange
        r0, c0 = int(used.Row), int(used.Column)
        formulas_2d = _normalize_com_value(used.Formula)
        values_2d = _normalize_com_value(used.Value)
        calc_hardcodes.extend(
            FinancialModelClassifier.audit_calc_hardcodes_2d(s_name, formulas_2d, values_2d, start_row=r0, start_col=c0)
        )

        cells_meta: list[dict[str, Any]] = []
        has_content = False
        scanned_nonempty = 0
        for r_off, row_vals in enumerate(values_2d):
            f_row = formulas_2d[r_off] if r_off < len(formulas_2d) else []
            for c_off, val in enumerate(row_vals):
                f_val = f_row[c_off] if c_off < len(f_row) else ""
                has_formula = bool(str(f_val or "").strip())
                if val is None and not has_formula:
                    continue
                has_content = True
                scanned_nonempty += 1
                if scanned_nonempty > 2500:
                    continue
                cell_addr = f"{_col_letter(c0 + c_off)}{r0 + r_off}"
                disp = ""
                try:
                    if isinstance(val, int) and val in EXCEL_CVERR_MAP:
                        disp = EXCEL_CVERR_MAP[val]
                    else:
                        cell_obj = sh.Cells(r0 + r_off, c0 + c_off)
                        disp = str(cell_obj.Text or "")
                        if r_off <= 1 and disp.strip():
                            try:
                                if int(getattr(cell_obj.Interior, "ColorIndex", -4142)) != -4142:
                                    contrast_pairs.append(
                                        {
                                            "location": f"{s_name}!{cell_addr}",
                                            "bg_rgb": bgr_to_rgb(cell_obj.Interior.Color),
                                            "text_rgb": bgr_to_rgb(cell_obj.Font.Color),
                                        }
                                    )
                            except Exception:
                                pass
                except Exception:
                    disp = str(val or "")
                if disp.strip() or (isinstance(val, int) and val in EXCEL_CVERR_MAP):
                    cells_meta.append({"cell": cell_addr, "display": disp, "formula": f_val, "value": val})
                    if isinstance(val, str):
                        text_blocks.append({"location": f"{s_name}!{cell_addr}", "text": val})

        try:
            for sh_i in range(1, getattr(sh.Shapes, "Count", 0) + 1):
                xsh = sh.Shapes(sh_i)
                try:
                    if getattr(xsh, "HasChart", 0) and getattr(xsh.Chart, "HasTitle", False):
                        ct = str(xsh.Chart.ChartTitle.Text or "").strip()
                        if ct:
                            text_blocks.append({"location": f"{s_name}!chart[{sh_i}]", "text": ct})
                except Exception:
                    pass
                try:
                    if getattr(xsh, "TextFrame2", None) and getattr(xsh.TextFrame2, "HasText", 0):
                        xt = str(xsh.TextFrame2.TextRange.Text or "").strip()
                        if xt:
                            text_blocks.append({"location": f"{s_name}!shape[{sh_i}]", "text": xt})
                except Exception:
                    pass
        except Exception:
            pass

        print_area = ""
        try:
            print_area = str(sh.PageSetup.PrintArea or "")
        except Exception:
            pass
        sheets_meta.append({"name": s_name, "has_content": has_content, "cells": cells_meta, "print_area": print_area})

    return {
        "sheets": sheets_meta,
        "text_blocks": text_blocks,
        "calc_hardcodes": calc_hardcodes,
        "contrast_pairs": contrast_pairs,
    }


def _extract_hwp_snapshot(hwp: Any) -> dict[str, Any]:
    raw_txt = ""
    try:
        raw_txt = hwp.GetTextFile("TEXT", "") or ""
    except Exception:
        pass
    paras = [line.strip() for line in raw_txt.replace("\r\n", "\n").split("\n")]
    blocks = [{"location": f"p[{idx}]", "text": p} for idx, p in enumerate(paras, start=1) if p]
    return {"paragraphs": paras, "tables_count": 0, "text_blocks": blocks}


def _replace_in_ppt_shape(sh: Any, token: str, val_str: str) -> int:
    """PowerPoint 도형(그룹/표/차트제목 포함) 내부의 `{{key}}` 토큰을 서식 유지 치환."""
    count = 0
    try:
        if getattr(sh, "Type", 0) == 6:  # msoGroup
            for g_i in range(1, sh.GroupItems.Count + 1):
                count += _replace_in_ppt_shape(sh.GroupItems(g_i), token, val_str)
            return count
    except Exception:
        pass

    try:
        if getattr(sh, "HasTextFrame", 0) and getattr(sh.TextFrame, "HasText", 0):
            tr = sh.TextFrame.TextRange
            cur_text = tr.Text or ""
            if token in cur_text:
                occ = cur_text.count(token)
                try:
                    tr.Replace(token, val_str)
                except Exception:
                    tr.Text = cur_text.replace(token, val_str)
                count += occ
    except Exception:
        pass

    try:
        if getattr(sh, "HasTable", 0):
            tbl = sh.Table
            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    ctr = tbl.Cell(r, c).Shape.TextFrame.TextRange
                    ct = ctr.Text or ""
                    if token in ct:
                        occ = ct.count(token)
                        try:
                            ctr.Replace(token, val_str)
                        except Exception:
                            ctr.Text = ct.replace(token, val_str)
                        count += occ
    except Exception:
        pass

    try:
        if getattr(sh, "HasChart", 0) and getattr(sh.Chart, "HasTitle", False):
            ct = str(sh.Chart.ChartTitle.Text or "")
            if token in ct:
                sh.Chart.ChartTitle.Text = ct.replace(token, val_str)
                count += ct.count(token)
    except Exception:
        pass

    return count


def _replace_in_excel_shape(xsh: Any, token: str, val_str: str) -> int:
    """Excel 도형(그룹/텍스트박스/차트제목 포함) 내부의 `{{key}}` 토큰을 치환."""
    count = 0
    try:
        if getattr(xsh, "Type", 0) == 6:  # msoGroup
            for g_i in range(1, xsh.GroupItems.Count + 1):
                count += _replace_in_excel_shape(xsh.GroupItems(g_i), token, val_str)
            return count
    except Exception:
        pass
    try:
        if getattr(xsh, "HasChart", 0) and getattr(xsh.Chart, "HasTitle", False):
            ct = str(xsh.Chart.ChartTitle.Text or "")
            if token in ct:
                xsh.Chart.ChartTitle.Text = ct.replace(token, val_str)
                count += ct.count(token)
    except Exception:
        pass
    try:
        if getattr(xsh, "TextFrame2", None) and getattr(xsh.TextFrame2, "HasText", 0):
            xt = str(xsh.TextFrame2.TextRange.Text or "")
            if token in xt:
                xsh.TextFrame2.TextRange.Text = xt.replace(token, val_str)
                count += xt.count(token)
    except Exception:
        pass
    return count


def register_delivery_gate_tools(mcp: MCPServer) -> None:
    """5-Gate 납품 감사, 전사 {{key}} 병합, Mark-and-Review 및 시각 스냅샷 도구 등록."""

    @mcp.tool()
    def office_run_delivery_gate(
        target_app: Literal["auto", "word", "excel", "ppt", "hwp"] = "auto",
    ) -> dict[str, Any]:
        """
        현재 활성화된 오피스 문서(Word, Excel, PPT, HWP)에 대해 납품 전 5단계 품질 게이트를 일괄 검증합니다:
        - Gate 1 (구조/스키마): 빈 슬라이드, 연속 3줄 빈 문단, 빈 시트 점검
        - Gate 2 (토큰/이스케이프 누수): `{{...}}`, `<TODO>`, `**`, `##`, `\\$`, Excel `###` 및 `#REF!`/`#DIV/0!`
        - Gate 3 (필수 구성요소): PPT 발표자 노트, Word 3+ 헤딩 목차/PAGE, Excel Outputs 인쇄 영역
        - Gate 4 (시각 대비 & 수리 무결성): 다크 배경 위 어두운 글씨 저대비, `IMBALANCED` 및 Calc 하드코딩
        - Gate 5 (도메인 특화): Morph 유령 도형 잔존, Word `____` 밑줄 가짜 양식 안티패턴
        """
        ensure_com()
        kind, app, doc = _resolve_live_app(target_app)
        if app is None or doc is None:
            return {"success": False, "error": f"열려 있는 {target_app} 문서를 찾을 수 없습니다."}

        try:
            if kind == "ppt":
                snapshot = _extract_ppt_snapshot(doc)
                doc_name = getattr(doc, "Name", "Presentation")
            elif kind == "word":
                snapshot = _extract_word_snapshot(doc)
                doc_name = getattr(doc, "Name", "Document")
            elif kind == "excel":
                snapshot = _extract_excel_snapshot(doc)
                doc_name = getattr(doc, "Name", "Workbook")
            else:
                snapshot = _extract_hwp_snapshot(doc)
                doc_name = getattr(doc, "Path", "") or "HWP"

            result = DeliveryGateScanner.evaluate(kind, snapshot)
            result["success"] = True
            result["document_name"] = doc_name
            return result
        except Exception as e:
            return {"success": False, "error": f"5-Gate 감사 실행 실패: {e}"}

    @mcp.tool()
    def office_merge_template_tokens(
        data: dict[str, Any],
        target_app: Literal["auto", "word", "excel", "ppt", "hwp"] = "auto",
        temporarily_unprotect: bool = True,
    ) -> dict[str, Any]:
        """
        현재 열린 템플릿 문서(Word, Excel, PPT, HWP)의 기존 폰트·색상·레이아웃을 보존한 채
        `{{key}}` 플레이스홀더를 `data` 딕셔너리 값으로 무손실 일괄 치환(Mail-Merge)합니다.
        """
        ensure_com()
        if not data:
            return {"success": False, "error": "병합할 data 딕셔너리가 비어 있습니다."}

        kind, app, doc = _resolve_live_app(target_app)
        if app is None or doc is None:
            return {"success": False, "error": f"열려 있는 {target_app} 문서를 찾을 수 없습니다."}

        token_pairs = [(f"{{{{{k}}}}}", str(v)) for k, v in data.items()]
        replaced_counts: dict[str, int] = {}

        try:
            if kind == "word":
                orig_prot = getattr(doc, "ProtectionType", WD_NO_PROTECTION)
                was_protected = orig_prot != WD_NO_PROTECTION
                if was_protected:
                    if not temporarily_unprotect:
                        return {"success": False, "error": "Word 문서가 양식 보호 상태여서 치환할 수 없습니다. temporarily_unprotect=True로 호출하세요."}
                    try:
                        doc.Unprotect("")
                    except Exception as pe:
                        return {"success": False, "error": f"비밀번호 등으로 보호된 Word 문서는 자동 해제할 수 없습니다: {pe}"}

                try:
                    with word_undo_group(app, "MCP 템플릿 토큰 병합"):
                        body_txt_snapshot = str(getattr(doc.Content, "Text", "") or "")
                        for token, val in token_pairs:
                            hits = 0
                            # 본문 + 본문 표 + SDT 콘텐츠
                            expected_occ = body_txt_snapshot.count(token)
                            f_main = doc.Content.Find
                            f_main.ClearFormatting()
                            f_main.Replacement.ClearFormatting()
                            if f_main.Execute(token, False, False, False, False, False, True, 1, False, val, 2):
                                hits += max(1, expected_occ)
                            # ContentControls 잔여 토큰(잠금 해제된 내부 텍스트) 백업 치환
                            try:
                                for cc_i in range(1, getattr(doc.ContentControls, "Count", 0) + 1):
                                    cc = doc.ContentControls(cc_i)
                                    if int(getattr(cc, "Type", 1)) in (0, 1, 3, 4):
                                        cc_txt = str(cc.Range.Text or "")
                                        if token in cc_txt:
                                            hits += cc_txt.count(token)
                                            cc.Range.Text = cc_txt.replace(token, val)
                            except Exception:
                                pass
                            # 도형 텍스트박스
                            for sh_i in range(1, getattr(doc.Shapes, "Count", 0) + 1):
                                try:
                                    sh = doc.Shapes(sh_i)
                                    if sh.TextFrame.HasText:
                                        sh_occ = str(sh.TextFrame.TextRange.Text or "").count(token)
                                        f_sh = sh.TextFrame.TextRange.Find
                                        f_sh.ClearFormatting()
                                        if f_sh.Execute(token, False, False, False, False, False, True, 1, False, val, 2):
                                            hits += max(1, sh_occ)
                                except Exception:
                                    continue
                            # 머리글 / 바닥글
                            for s_i in range(1, doc.Sections.Count + 1):
                                sec = doc.Sections(s_i)
                                for coll in (sec.Headers, sec.Footers):
                                    for idx in (1, 2, 3):
                                        try:
                                            hf_rng = coll(idx).Range
                                            hf_occ = str(hf_rng.Text or "").count(token)
                                            fh = hf_rng.Find
                                            fh.ClearFormatting()
                                            if fh.Execute(token, False, False, False, False, False, True, 1, False, val, 2):
                                                hits += max(1, hf_occ)
                                        except Exception:
                                            continue
                            if hits:
                                replaced_counts[token] = hits
                finally:
                    if was_protected:
                        try:
                            doc.Protect(orig_prot if orig_prot != WD_NO_PROTECTION else WD_ALLOW_ONLY_FORM_FIELDS, True, "")
                        except Exception:
                            pass
                com_manager.activate_app("word")

            elif kind == "ppt":
                with ppt_freeze_display(app):
                    for s_i in range(1, doc.Slides.Count + 1):
                        sl = doc.Slides(s_i)
                        for sh_i in range(1, sl.Shapes.Count + 1):
                            sh = sl.Shapes(sh_i)
                            for token, val in token_pairs:
                                cnt = _replace_in_ppt_shape(sh, token, val)
                                if cnt:
                                    replaced_counts[token] = replaced_counts.get(token, 0) + cnt
                        # 발표자 노트 내 토큰 치환
                        try:
                            for n_i in range(1, sl.NotesPage.Shapes.Count + 1):
                                nsh = sl.NotesPage.Shapes(n_i)
                                for token, val in token_pairs:
                                    cnt = _replace_in_ppt_shape(nsh, token, val)
                                    if cnt:
                                        replaced_counts[token] = replaced_counts.get(token, 0) + cnt
                        except Exception:
                            pass
                com_manager.activate_app("ppt")

            elif kind == "excel":
                with excel_freeze_display(app, pause_calculation=True):
                    for s_i in range(1, doc.Sheets.Count + 1):
                        sh = doc.Sheets(s_i)
                        used = sh.UsedRange
                        r0, c0 = int(used.Row), int(used.Column)
                        formulas_2d = _normalize_com_value(used.Formula)
                        for r_off, f_row in enumerate(formulas_2d):
                            for c_off, f_val in enumerate(f_row):
                                if isinstance(f_val, str) and "{{" in f_val and not f_val.startswith("="):
                                    new_val = f_val
                                    for token, val in token_pairs:
                                        if token in new_val:
                                            replaced_counts[token] = replaced_counts.get(token, 0) + new_val.count(token)
                                            new_val = new_val.replace(token, val)
                                    if new_val != f_val:
                                        c_obj = sh.Cells(r0 + r_off, c0 + c_off)
                                        if new_val.isdigit() and new_val.startswith("0") and len(new_val) > 1:
                                            try:
                                                c_obj.NumberFormat = "@"
                                            except Exception:
                                                pass
                                        c_obj.Value = new_val
                        try:
                            for sh_i in range(1, getattr(sh.Shapes, "Count", 0) + 1):
                                xsh = sh.Shapes(sh_i)
                                for token, val in token_pairs:
                                    cnt = _replace_in_excel_shape(xsh, token, val)
                                    if cnt:
                                        replaced_counts[token] = replaced_counts.get(token, 0) + cnt
                        except Exception:
                            pass
                com_manager.activate_app("excel")

            elif kind == "hwp":
                for token, val in token_pairs:
                    try:
                        pset = app.HParameterSet.HFindReplace
                        app.HAction.GetDefault("AllReplace", pset.HSet)
                        pset.FindString = token
                        pset.ReplaceString = val
                        pset.IgnoreMessage = 1
                        pset.FindRegExp = 0
                        if app.HAction.Execute("AllReplace", pset.HSet):
                            replaced_counts[token] = replaced_counts.get(token, 0) + 1
                    except Exception:
                        continue
                com_manager.activate_app("hwp")

            return {
                "success": True,
                "app": kind,
                "replaced_tokens_count": sum(replaced_counts.values()),
                "replaced_by_key": replaced_counts,
            }
        except Exception as e:
            return {"success": False, "error": f"템플릿 토큰 병합 실패: {e}"}

    @mcp.tool()
    def office_mark_review_item(
        note: str,
        target_app: Literal["auto", "word", "excel", "ppt"] = "auto",
        target_ref: str = "",
        find_text: str = "",
        color_hex: str = "#FFF59D",
        author: str = "AI Reviewer",
        slide_index: int | None = None,
    ) -> dict[str, Any]:
        """
        문서 본문을 곧바로 덮어쓰기 전에 수정 제안 대상(단어/문단/셀/도형)에 리뷰 메모(Comment)와
        강조 색상을 비파괴적으로 마킹(Mark-and-Review)하여 사용자의 검토 승인을 지원합니다.
        """
        ensure_com()
        kind, app, doc = _resolve_live_app(target_app)
        if app is None or doc is None or kind == "hwp":
            return {"success": False, "error": "열려 있는 Word, Excel 또는 PowerPoint 문서가 없습니다."}

        try:
            if kind == "word":
                if getattr(doc, "ProtectionType", WD_NO_PROTECTION) != WD_NO_PROTECTION:
                    return {
                        "success": False,
                        "error": "Word 문서가 양식 보호 상태이므로 리뷰 메모/하이라이트를 추가할 수 없습니다. 보호 적용 전 또는 해제 후 마킹하세요.",
                    }
                rng = None
                if find_text:
                    srch = doc.Content.Duplicate
                    if srch.Find.Execute(find_text):
                        rng = srch
                if rng is None:
                    rng = app.Selection.Range
                rng.HighlightColorIndex = 7  # wdYellow
                doc.Comments.Add(rng, f"[{author}] {note}")
                com_manager.activate_app("word")
                return {"success": True, "app": "word", "marked_text": (rng.Text or "")[:80], "note": note}

            if kind == "excel":
                cell = None
                if target_ref:
                    if "!" in target_ref:
                        s_name, c_addr = target_ref.split("!", 1)
                        cell = doc.Sheets(s_name.strip("'")).Range(c_addr)
                    else:
                        cell = doc.ActiveSheet.Range(target_ref)
                elif find_text:
                    cell = doc.ActiveSheet.UsedRange.Find(find_text)
                if cell is None:
                    cell = app.Selection
                cell.Interior.Color = hex_to_bgr(color_hex)
                try:
                    cell.ClearComments()
                except Exception:
                    pass
                try:
                    ct = getattr(cell, "CommentThreaded", None)
                    if ct is not None:
                        ct.Delete()
                except Exception:
                    pass
                cell.AddComment(f"[{author}] {note}")
                com_manager.activate_app("excel")
                return {"success": True, "app": "excel", "cell": getattr(cell, "Address", ""), "note": note}

            if kind == "ppt":
                parsed_slide = slide_index
                clean_ref = target_ref
                m_sl = re.match(r"^slide\[(\d+)\]/(.*)$", target_ref or "", re.IGNORECASE)
                if m_sl:
                    if parsed_slide is None:
                        parsed_slide = int(m_sl.group(1))
                    clean_ref = m_sl.group(2)

                if parsed_slide is not None and 1 <= int(parsed_slide) <= doc.Slides.Count:
                    slide = doc.Slides(int(parsed_slide))
                else:
                    slide = app.ActiveWindow.View.Slide
                pos_x, pos_y = 60.0, 60.0
                if clean_ref or find_text:
                    for sh_i in range(1, slide.Shapes.Count + 1):
                        sh = slide.Shapes(sh_i)
                        sh_name = str(getattr(sh, "Name", ""))
                        sh_txt = (
                            sh.TextFrame.TextRange.Text
                            if getattr(sh, "HasTextFrame", 0) and getattr(sh.TextFrame, "HasText", 0)
                            else ""
                        )
                        if (clean_ref and clean_ref in sh_name) or (find_text and find_text in sh_txt):
                            pos_x = float(sh.Left) + float(sh.Width) - 10.0
                            pos_y = float(sh.Top)
                            break
                initials = (author[:2] or "AI").upper()
                slide.Comments.Add(pos_x, pos_y, author, initials, note)
                com_manager.activate_app("ppt")
                return {"success": True, "app": "ppt", "slide_index": slide.SlideIndex, "position": (pos_x, pos_y), "note": note}

            return {"success": False, "error": f"지원하지 않는 앱: {kind}"}
        except Exception as e:
            return {"success": False, "error": f"리뷰 마킹 실패: {e}"}

    @mcp.tool()
    def office_export_visual_snapshot(
        target_app: Literal["auto", "word", "excel", "ppt"] = "auto",
        slide_or_page: int | None = None,
        output_dir: str = "",
        export_format: Literal["png", "pdf"] = "png",
        export_all: bool = False,
        png_width: int = 1920,
        png_height: int = 1080,
    ) -> dict[str, Any]:
        """
        현재 작업 중인 활성 문서의 실제 렌더링 화면을 PNG 또는 PDF 스냅샷으로 내보내어
        멀티모달 에이전트가 레이아웃/텍스트 잘림/겹침을 직접 눈으로 검증(Render -> See -> Fix)할 수 있게 합니다.
        (원본 편집 문서의 경로/열림 상태는 변경하지 않고 비파괴 Export합니다.)
        """
        ensure_com()
        kind, app, doc = _resolve_live_app(target_app)
        if app is None or doc is None or kind == "hwp":
            return {"success": False, "error": "열려 있는 Word, Excel 또는 PowerPoint 문서가 없습니다."}

        base_dir = Path(output_dir) if output_dir else (Path.cwd() / "scratch" / "snapshots")
        base_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        outputs: list[str] = []

        try:
            if kind == "ppt":
                if export_format == "pdf":
                    out_pdf = str((base_dir / f"ppt_snapshot_{stamp}.pdf").resolve())
                    doc.ExportAsFixedFormat(out_pdf, 2)  # ppFixedFormatTypePDF = 2
                    outputs.append(out_pdf)
                else:
                    if export_all:
                        target_indices = list(range(1, doc.Slides.Count + 1))
                    elif slide_or_page is not None:
                        target_indices = [int(slide_or_page)]
                    else:
                        try:
                            target_indices = [int(app.ActiveWindow.View.Slide.SlideIndex)]
                        except Exception:
                            target_indices = [1] if doc.Slides.Count >= 1 else []
                    for s_i in target_indices:
                        sl = doc.Slides(s_i)
                        out_png = str((base_dir / f"slide_{s_i}_{stamp}.png").resolve())
                        sl.Export(out_png, "PNG", int(png_width), int(png_height))
                        outputs.append(out_png)
                return {"success": True, "app": "ppt", "format": export_format, "output_paths": outputs}

            if kind == "word":
                out_pdf = str((base_dir / f"word_snapshot_{stamp}.pdf").resolve())
                doc.ExportAsFixedFormat(out_pdf, 17)  # wdExportFormatPDF = 17
                outputs.append(out_pdf)
                return {"success": True, "app": "word", "format": "pdf", "output_paths": outputs}

            if kind == "excel":
                out_pdf = str((base_dir / f"excel_snapshot_{stamp}.pdf").resolve())
                if export_all:
                    doc.ExportAsFixedFormat(0, out_pdf)  # 전체 시트 통합 PDF
                else:
                    doc.ActiveSheet.ExportAsFixedFormat(0, out_pdf)  # xlTypePDF = 0
                outputs.append(out_pdf)
                return {"success": True, "app": "excel", "format": "pdf", "output_paths": outputs}

            return {"success": False, "error": f"지원하지 않는 앱: {kind}"}
        except Exception as e:
            return {"success": False, "error": f"시각 스냅샷 내보내기 실패: {e}"}

    @mcp.tool()
    def word_compare_authenticity_gate(
        archetype: Literal["auto", "corporate_consent_form", "corporate_submission", "kisa_executive_report"] = "auto",
        reference_path: str = "",
        candidate_path: str = "",
    ) -> dict[str, Any]:
        """
        Gate 6: 현재 열려 있는 Word 문서(또는 candidate_path DOCX)를 기업 실무 양식/KISA 표준 아키타입 및
        원본 레퍼런스(reference_path)와 비교하여 구조 유사도(%), AI 합성 티(AI-Tell) 유무, 진위 판별(Real vs AI)을 수행합니다.
        """
        try:
            if candidate_path.strip():
                cand_p = Path(candidate_path.strip())
                if not cand_p.is_absolute():
                    cand_p = (Path.cwd() / cand_p).resolve()
                cand_prof = extract_docx_style_profile(cand_p)
                doc_name = cand_p.name
            else:
                ensure_com()
                word_app, doc = com_manager.get_word(create_if_missing=False)
                if word_app is None or doc is None:
                    return {"success": False, "error": "열려 있는 Word 문서가 없으며 candidate_path도 지정되지 않았습니다."}
                cand_prof = profile_from_word_snapshot(_extract_word_snapshot(doc))
                doc_name = getattr(doc, "Name", "Document")

            if reference_path.strip():
                ref_p = Path(reference_path.strip())
                if not ref_p.is_absolute():
                    ref_p = (Path.cwd() / ref_p).resolve()
                comp = compare_real_vs_generated(ref_p, cand_prof, archetype=archetype)
                return {
                    "success": True,
                    "document_name": doc_name,
                    "mode": "pairwise_reference_comparison",
                    **comp,
                }

            auth = evaluate_docx_authenticity(cand_prof, archetype=archetype)
            return {
                "success": True,
                "document_name": doc_name,
                "mode": "archetype_baseline",
                **auth,
            }
        except Exception as e:
            return {"success": False, "error": f"Gate 6 진위/유사도 평가 실패: {e}"}
