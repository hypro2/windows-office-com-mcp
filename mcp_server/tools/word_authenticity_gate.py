# -*- coding: utf-8 -*-
"""
mcp_server/tools/word_authenticity_gate.py
Gate 6: 실무 기업·KISA 문서 스타일 프로파일 기반 유사도 평가 및
AI 합성 징후(AI-Tell) vs 인간 실무 문서 판별 게이트.
"""

from __future__ import annotations

import math
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    import winreg
except ImportError:  # pragma: no cover
    winreg = None

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_TAG = f"{{{W_NS}}}"

_AI_EMOJI_RE = re.compile(r"[💡ℹ️⚠️📌✅🚀🔥🎯📊]")
_SPACE_FAKED_INDENT_RE = re.compile(r"^\s{2,}[○\-\*•]")

# 실무 표준 PDF/DOCX 및 현대화 KISA 보고서의 검증된 팔레트 (RGB Hex)
AUTHENTIC_FILL_HEXES: set[str] = {
    "FFFFFF", "F7FAF5", "C5DFB3", "A9D08E", "E2EFDA", "E6EEF7",
    "E7E6E6", "F2F2F2", "D9E1F2", "F4F8F7", "F4F7FA", "F2F8F6",
    "FFFBEB", "F8FAFC", "F4F6F9", "1F3864", "0F766E", "2E608A",
    "404040", "002060", "E9ECEF", "EBF3FA",
}

_SAFE_COMMON_FONT_TOKENS = (
    "맑은 고딕", "malgun", "times", "calibri", "consolas", "arial", "cambria",
    "바탕", "batang", "돋움", "dotum", "굴림", "gulim", "나눔", "nanum",
    "함초롬", "휴먼명조", "hy헤드라인", "pretendard", "noto",
)

_INSTALLED_WIN_FONTS: set[str] | None = None


def _is_font_installed(font_name: str) -> bool:
    norm = (font_name or "").strip().lower()
    if not norm:
        return True
    if any(tok in norm for tok in _SAFE_COMMON_FONT_TOKENS):
        return True
    global _INSTALLED_WIN_FONTS
    if _INSTALLED_WIN_FONTS is None:
        keys: set[str] = set()
        if winreg is not None:
            subkey = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(root, subkey) as hk:
                        idx = 0
                        while True:
                            try:
                                val_name, _val_data, _ = winreg.EnumValue(hk, idx)
                                clean_name = re.sub(r"\s*\(.*?\)\s*$", "", str(val_name)).strip().lower()
                                if clean_name:
                                    keys.add(clean_name)
                                idx += 1
                            except OSError:
                                break
                except OSError:
                    continue
        _INSTALLED_WIN_FONTS = keys
    if not _INSTALLED_WIN_FONTS:
        return False
    base_token = re.sub(r"\s+(bold|regular|ttf|light|medium)\b", "", norm).strip()
    return any(norm in k or (base_token and base_token in k) for k in _INSTALLED_WIN_FONTS)

CORPORATE_KISA_CUE_RE = re.compile(
    r"(귀중|결재|담\s*당|책임자|본부장|팀장|부\(점\)장|서명|확인|\(주\)|주식회사|"
    r"현\s*행|개\s*정|비\s*고|제\s*\d+\s*조|부\s*칙|\[참고\]|\[안내\]|\[유의사항\]|\[핵심\s*요약\]|"
    r"\(인\)|신청일|신청자|동의|신용정보|금융거래)"
)

_CORPUS_BASELINE_CACHE: dict[str, dict[str, Any]] | None = None


def _hex_to_rgb(hex_str: str) -> tuple[int, int, int] | None:
    s = (hex_str or "").strip().lstrip("#").upper()
    if len(s) != 6:
        return None
    try:
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except ValueError:
        return None


def _rel_luminance(rgb: tuple[int, int, int]) -> float:
    def _chan(c: int) -> float:
        v = c / 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = rgb
    return 0.2126 * _chan(r) + 0.7152 * _chan(g) + 0.0722 * _chan(b)


def wcag_contrast_ratio(bg_hex: str, fg_hex: str) -> float:
    """두 RGB Hex 색상의 WCAG 명암비(1.0~21.0) 반환."""
    bg = _hex_to_rgb(bg_hex)
    fg = _hex_to_rgb(fg_hex)
    if bg is None or fg is None:
        return 21.0
    l1 = _rel_luminance(bg)
    l2 = _rel_luminance(fg)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def is_neon_low_contrast_fill(fill_hex: str, text_hex: str = "FFFFFF") -> bool:
    """밝은 연두/네온 배경 위 흰색 글씨 저대비(<3.5:1) 판별."""
    norm = (fill_hex or "").strip().lstrip("#").upper()
    fg = (text_hex or "FFFFFF").strip().lstrip("#").upper()
    if norm in ("", "AUTO", "FFFFFF", "CLEAR"):
        return False
    rgb = _hex_to_rgb(norm)
    fg_rgb = _hex_to_rgb(fg)
    if rgb is None or fg_rgb is None:
        return False
    # 밝은 채색 셀에 흰 글씨(또는 밝은 글씨)인 경우 명암비 검사
    if _rel_luminance(fg_rgb) > 0.75 and wcag_contrast_ratio(norm, fg) < 3.5:
        return True
    return False


def _wattr(elem: ET.Element, name: str) -> str:
    return elem.attrib.get(f"{W_TAG}{name}", elem.attrib.get(name, ""))


def extract_docx_style_profile(docx_path: str | Path) -> dict[str, Any]:
    """OOXML(.docx ZIP)을 COM 없이 즉시 파싱하여 실무 문서 스타일 프로파일을 추출."""
    p = Path(docx_path)
    if not p.exists():
        raise FileNotFoundError(f"DOCX 파일이 존재하지 않습니다: {p}")

    fonts: set[str] = set()
    fills: set[str] = set()
    colors: set[str] = set()
    sizes_pt: list[float] = []
    paragraphs: list[str] = []
    body_paragraphs: list[str] = []
    margins_pt: dict[str, float] = {}

    tables_count = 0
    bordered_tables_count = 0
    padded_tables_count = 0
    sdt_total = 0
    sdt_in_table = 0
    sdt_style_leaks = 0
    neon_low_contrast_count = 0
    dark_runs = 0
    total_colored_runs = 0
    has_header_footer = False
    has_page_field = False

    with zipfile.ZipFile(p, "r") as zf:
        names = set(zf.namelist())
        xml_targets = [n for n in ("word/styles.xml", "word/document.xml") if n in names]
        hf_targets = sorted(n for n in names if re.match(r"word/(header|footer)\d*\.xml$", n))
        if hf_targets:
            has_header_footer = True

        for xname in xml_targets + hf_targets:
            root = ET.fromstring(zf.read(xname))

            for rf in root.iter(f"{W_TAG}rFonts"):
                for attr in ("ascii", "eastAsia", "hAnsi"):
                    val = _wattr(rf, attr).strip()
                    if val and not val.startswith("+"):
                        fonts.add(val)

            for instr in root.iter(f"{W_TAG}instrText"):
                if "PAGE" in (instr.text or "").upper():
                    has_page_field = True
            for fld in root.iter(f"{W_TAG}fldSimple"):
                if "PAGE" in _wattr(fld, "instr").upper():
                    has_page_field = True

            if xname == "word/document.xml":
                for pgm in root.iter(f"{W_TAG}pgMar"):
                    for side in ("top", "left", "right", "bottom"):
                        raw_tw = _wattr(pgm, side)
                        if raw_tw.isdigit():
                            margins_pt[side] = round(int(raw_tw) / 20.0, 1)

                # 문서 내 표 및 표 밖 문단/SDT 분석
                table_para_ids: set[int] = set()
                table_sdt_ids: set[int] = set()

                for tbl in root.iter(f"{W_TAG}tbl"):
                    tables_count += 1
                    has_bdr = any(
                        _wattr(b, "val") not in ("", "none", "nil")
                        for b in tbl.iter()
                        if b.tag.split("}")[-1] in ("top", "bottom", "left", "right", "insideH", "insideV")
                    )
                    if has_bdr:
                        bordered_tables_count += 1

                    has_pad = any(
                        _wattr(cm, "w") not in ("", "0")
                        for tag_name in (f"{W_TAG}tcMar", f"{W_TAG}tblCellMar")
                        for cm_parent in tbl.iter(tag_name)
                        for cm in cm_parent
                    )
                    if has_pad:
                        padded_tables_count += 1

                    for tp in tbl.iter(f"{W_TAG}p"):
                        table_para_ids.add(id(tp))
                    for tsdt in tbl.iter(f"{W_TAG}sdt"):
                        table_sdt_ids.add(id(tsdt))

                    # 셀 단위 채움 + 글자색 저대비 점검 (기본 AUTO 글자색은 검정 000000)
                    for tc in tbl.iter(f"{W_TAG}tc"):
                        cell_fill = ""
                        for shd in tc.iter(f"{W_TAG}shd"):
                            fval = _wattr(shd, "fill").upper()
                            if fval and fval not in ("AUTO", "CLEAR"):
                                fills.add(fval)
                                if not cell_fill:
                                    cell_fill = fval
                        cell_fg = ""
                        for col in tc.iter(f"{W_TAG}color"):
                            cval = _wattr(col, "val").upper()
                            if cval and cval != "AUTO":
                                cell_fg = cval
                                break
                        if cell_fill and is_neon_low_contrast_fill(cell_fill, cell_fg or "000000"):
                            neon_low_contrast_count += 1

                for sdt in root.iter(f"{W_TAG}sdt"):
                    sdt_total += 1
                    if id(sdt) in table_sdt_ids:
                        sdt_in_table += 1
                    # SDT 내부 텍스트가 비정상 거대 폰트(>=14pt)나 네온색으로 오염되었는지 점검
                    for rpr in sdt.iter(f"{W_TAG}rPr"):
                        sz_el = rpr.find(f"{W_TAG}sz")
                        if sz_el is not None:
                            sval = _wattr(sz_el, "val")
                            if sval.isdigit() and int(sval) >= 28:
                                sdt_style_leaks += 1
                                break

                for pel in root.iter(f"{W_TAG}p"):
                    txt = "".join((t.text or "") for t in pel.iter(f"{W_TAG}t")).strip()
                    paragraphs.append(txt)
                    if id(pel) not in table_para_ids:
                        body_paragraphs.append(txt)
                    for rpr in pel.iter(f"{W_TAG}rPr"):
                        sz_el = rpr.find(f"{W_TAG}sz")
                        if sz_el is not None:
                            sval = _wattr(sz_el, "val")
                            if sval.isdigit():
                                sizes_pt.append(int(sval) / 2.0)
                        col_el = rpr.find(f"{W_TAG}color")
                        if col_el is not None:
                            cval = _wattr(col_el, "val").upper()
                            if cval and cval != "AUTO":
                                colors.add(cval)
                                total_colored_runs += 1
                                rgb = _hex_to_rgb(cval)
                                if rgb is not None and _rel_luminance(rgb) < 0.35:
                                    dark_runs += 1
                            else:
                                total_colored_runs += 1
                                dark_runs += 1

    # 문단 내 공백 가짜 들여쓰기 탐지 (raw 텍스트 앞 공백 검사 위해 재스캔 없이 body/all 검사)
    space_faked_indent_hits = 0
    emoji_hits = 0
    corporate_cues_count = 0
    with zipfile.ZipFile(p, "r") as zf:
        if "word/document.xml" in zf.namelist():
            root = ET.fromstring(zf.read("word/document.xml"))
            for pel in root.iter(f"{W_TAG}p"):
                raw_line = "".join((t.text or "") for t in pel.iter(f"{W_TAG}t"))
                if _SPACE_FAKED_INDENT_RE.match(raw_line):
                    space_faked_indent_hits += 1
                emoji_hits += len(_AI_EMOJI_RE.findall(raw_line))
                corporate_cues_count += len(CORPORATE_KISA_CUE_RE.findall(raw_line))

    uninstalled_fonts = sorted(f for f in fonts if not _is_font_installed(f))
    sorted_sizes = sorted(sizes_pt)
    median_body_pt = sorted_sizes[len(sorted_sizes) // 2] if sorted_sizes else 10.5
    dark_body_ratio = round(dark_runs / max(total_colored_runs, 1), 3) if total_colored_runs else 1.0

    return {
        "source": str(p),
        "fonts": sorted(fonts),
        "uninstalled_fonts": uninstalled_fonts,
        "fills": sorted(fills),
        "colors": sorted(colors),
        "median_body_pt": median_body_pt,
        "margins_pt": margins_pt,
        "tables_count": tables_count,
        "bordered_tables_count": bordered_tables_count,
        "padded_tables_count": padded_tables_count,
        "sdt_total": sdt_total,
        "sdt_in_table": sdt_in_table,
        "sdt_style_leaks": sdt_style_leaks,
        "neon_low_contrast_count": neon_low_contrast_count,
        "emoji_hits": emoji_hits,
        "space_faked_indent_hits": space_faked_indent_hits,
        "dark_body_ratio": dark_body_ratio,
        "has_header_footer": has_header_footer or has_page_field,
        "corporate_cues_count": corporate_cues_count,
        "paragraphs": paragraphs,
        "body_paragraphs": body_paragraphs,
        "is_minimal_mock": False,
    }


def profile_from_word_snapshot(snapshot: dict[str, Any]) -> dict[str, Any]:
    """실시간 COM _extract_word_snapshot() 결과를 Gate 6 공통 스타일 프로파일로 정규화."""
    raw_paras = [str(x or "") for x in snapshot.get("raw_paragraphs", snapshot.get("paragraphs", []))]
    paras = [str(x or "") for x in snapshot.get("paragraphs", [])]
    body_paras = [str(x or "") for x in snapshot.get("body_paragraphs", paras)]
    controls = snapshot.get("controls", [])

    emoji_hits = int(snapshot.get("emoji_hits", sum(len(_AI_EMOJI_RE.findall(p)) for p in raw_paras)))
    space_faked = int(
        snapshot.get("space_faked_indent_hits", sum(1 for p in raw_paras if _SPACE_FAKED_INDENT_RE.match(p)))
    )
    cues = int(
        snapshot.get("corporate_cues_count", sum(len(CORPORATE_KISA_CUE_RE.findall(p)) for p in raw_paras))
    )

    fonts = sorted({str(f).strip() for f in snapshot.get("fonts", []) if str(f).strip()})
    uninstalled = sorted(f for f in fonts if not _is_font_installed(f))
    fills = sorted({str(f).strip().lstrip("#").upper() for f in snapshot.get("fills", []) if str(f).strip()})

    sdt_total = int(snapshot.get("sdt_total", len(controls)))
    sdt_in_table = int(
        snapshot.get("sdt_in_table", sum(1 for c in controls if c.get("in_table", False)))
    )
    sdt_style_leaks = int(
        snapshot.get(
            "sdt_style_leaks",
            sum(1 for c in controls if float(c.get("font_size", 10.5) or 10.5) >= 14.0),
        )
    )

    is_minimal_mock = "fonts" not in snapshot and "margins_pt" not in snapshot and not raw_paras

    return {
        "source": snapshot.get("source", "live_com_word"),
        "fonts": fonts,
        "uninstalled_fonts": uninstalled,
        "fills": fills,
        "colors": sorted({str(c).strip().lstrip("#").upper() for c in snapshot.get("colors", []) if str(c).strip()}),
        "median_body_pt": float(snapshot.get("median_body_pt", 10.5)),
        "margins_pt": dict(snapshot.get("margins_pt", {})),
        "tables_count": int(snapshot.get("tables_count", 0)),
        "bordered_tables_count": int(snapshot.get("bordered_tables_count", snapshot.get("tables_count", 0))),
        "padded_tables_count": int(snapshot.get("padded_tables_count", snapshot.get("tables_count", 0))),
        "sdt_total": sdt_total,
        "sdt_in_table": sdt_in_table,
        "sdt_style_leaks": sdt_style_leaks,
        "neon_low_contrast_count": int(snapshot.get("neon_low_contrast_count", 0)),
        "emoji_hits": emoji_hits,
        "space_faked_indent_hits": space_faked,
        "dark_body_ratio": float(snapshot.get("dark_body_ratio", 1.0)),
        "has_header_footer": bool(snapshot.get("has_header_footer", snapshot.get("has_page_field", False))),
        "corporate_cues_count": cues,
        "paragraphs": paras,
        "body_paragraphs": body_paras,
        "is_minimal_mock": is_minimal_mock,
    }


def load_corpus_archetype_baselines(corpus_dir: str | Path | None = None) -> dict[str, dict[str, Any]]:
    """기업 실무 양식 및 KISA 기준 3대 아키타입 기준선 로드."""
    global _CORPUS_BASELINE_CACHE
    if _CORPUS_BASELINE_CACHE is not None and corpus_dir is None:
        return _CORPUS_BASELINE_CACHE

    corp_profiles: list[dict[str, Any]] = []
    if corpus_dir is not None:
        target_dir = Path(corpus_dir)
        if target_dir.exists():
            for dp in sorted(target_dir.glob("*.docx")):
                try:
                    corp_profiles.append(extract_docx_style_profile(dp))
                except Exception:
                    continue

    avg_pt = (
        round(sum(p["median_body_pt"] for p in corp_profiles) / len(corp_profiles), 1)
        if corp_profiles
        else 10.5
    )

    baselines: dict[str, dict[str, Any]] = {
        "corporate_consent_form": {
            "label": "기업 실무 동의서/약정서 (Sage-Grid + 결재란 + 서명)",
            "expected_pt": 10.0,
            "margin_range": (24.0, 90.0),
            "preferred_fills": {"C5DFB3", "E2EFDA", "A9D08E", "F7FAF5", "F2F2F2", "FFFFFF"},
            "min_tables": 1,
        },
        "corporate_submission": {
            "label": "기업 제출양식/신구조문대비표 (Corporate Standard)",
            "expected_pt": avg_pt,
            "margin_range": (28.0, 95.0),
            "preferred_fills": {"E6EEF7", "E7E6E6", "F2F2F2", "C5DFB3", "FFFFFF"},
            "min_tables": 1,
            "corpus_files_count": len(corp_profiles),
        },
        "kisa_executive_report": {
            "label": "KISA 1310 공공·비즈니스 보고서 및 기술/학술 표준",
            "expected_pt": 10.5,
            "margin_range": (28.0, 95.0),
            "preferred_fills": {"1F3864", "0F766E", "2E608A", "F4F8F7", "F4F7FA", "F2F8F6", "F8FAFC", "F2F2F2", "FFFFFF"},
            "min_tables": 1,
        },
    }
    if corpus_dir is None:
        _CORPUS_BASELINE_CACHE = baselines
    return baselines


def infer_archetype(profile: dict[str, Any]) -> str:
    joined = "\n".join(profile.get("paragraphs", []))
    fills = set(profile.get("fills", []))
    if "C5DFB3" in fills or "E2EFDA" in fills or ("귀중" in joined and ("동의" in joined or "서명" in joined)):
        return "corporate_consent_form"
    if "E6EEF7" in fills or "신구조문" in joined or ("현행" in joined and "개정" in joined) or "제출양식" in joined:
        return "corporate_submission"
    return "kisa_executive_report"


def detect_ai_tells(profile: dict[str, Any]) -> list[dict[str, str]]:
    """인간 실무 문서에서 절대 나타나지 않는 대표적인 AI 작성 티(AI-Tell) 결함을 검출."""
    tells: list[dict[str, str]] = []
    if profile.get("uninstalled_fonts"):
        fonts_joined = " ".join(profile.get("fonts", [])).lower()
        has_safe_sys = any(
            k in fonts_joined
            for k in ("맑은 고딕", "malgun", "times", "calibri", "consolas", "바탕", "batang", "돋움", "dotum", "나눔", "nanum")
        )
        if not has_safe_sys:
            tells.append({
                "code": "uninstalled_font_fallback",
                "severity": "critical",
                "message": f"시스템 미설치 글꼴({', '.join(profile['uninstalled_fonts'])}) 참조로 대체 폰트 렌더링이 발생합니다.",
            })
    if int(profile.get("emoji_hits", 0)) > 0:
        tells.append({
            "code": "ai_emoji_prefix",
            "severity": "critical",
            "message": f"공식 비즈니스/금융 문서에 부적절한 이모지({profile['emoji_hits']}건)가 포함되어 있습니다.",
        })
    if int(profile.get("neon_low_contrast_count", 0)) > 0:
        tells.append({
            "code": "neon_low_contrast",
            "severity": "critical",
            "message": "네온/밝은 배경 + 흰색 글씨 등 저대비 원색 음영이 감지되었습니다.",
        })
    if int(profile.get("space_faked_indent_hits", 0)) > 0:
        tells.append({
            "code": "space_faked_indent",
            "severity": "critical",
            "message": f"문단 들여쓰기 대신 공백 문자로 가짜 계층({profile['space_faked_indent_hits']}건)을 흉내 냈습니다.",
        })
    sdt_total = int(profile.get("sdt_total", 0))
    sdt_in_tbl = int(profile.get("sdt_in_table", 0))
    if sdt_total >= 2 and sdt_in_tbl == 0 and int(profile.get("tables_count", 0)) <= 1:
        tells.append({
            "code": "unboxed_floating_form_lines",
            "severity": "warning",
            "message": f"SDT 양식 컨트롤({sdt_total}개)이 표 격자 없이 본문 줄에 흩어져 있어 실무 양식 완성도가 떨어집니다.",
        })
    if int(profile.get("sdt_style_leaks", 0)) > 0:
        tells.append({
            "code": "cross_paragraph_style_leak",
            "severity": "critical",
            "message": f"이전 제목/강조 서식(>=14pt)이 SDT 또는 본문 입력 필드({profile['sdt_style_leaks']}건)로 전이되었습니다.",
        })
    return tells


def evaluate_docx_authenticity(profile: dict[str, Any], archetype: str = "auto") -> dict[str, Any]:
    """문서 프로파일의 실무 유사도(0~100%)와 AI-Tell 감점/판정 결과를 반환."""
    baselines = load_corpus_archetype_baselines()
    resolved_arch = infer_archetype(profile) if archetype == "auto" else archetype
    spec = baselines.get(resolved_arch, baselines["kisa_executive_report"])

    # 1. 타이포그래피 청결도 (/25)
    fonts = profile.get("fonts", [])
    uninstalled = profile.get("uninstalled_fonts", [])
    has_cleartype_fallback = any(
        k in " ".join(fonts).lower()
        for k in ("맑은 고딕", "malgun", "times", "calibri", "consolas", "바탕", "batang", "돋움", "dotum", "나눔", "nanum")
    )
    typo_score = 15.0 if (not uninstalled or has_cleartype_fallback) else 0.0
    if len(fonts) <= 3:
        typo_score += 10.0
    elif len(fonts) <= 6:
        typo_score += 7.0
    else:
        typo_score += 4.0

    # 2. 색상 및 명암비 충실도 (/25)
    neon_cnt = int(profile.get("neon_low_contrast_count", 0))
    fills = set(profile.get("fills", []))
    dark_ratio = float(profile.get("dark_body_ratio", 1.0))
    palette_score = 12.0 if neon_cnt == 0 else 0.0
    palette_score += 7.0 if dark_ratio >= 0.85 else (4.0 if dark_ratio >= 0.65 else 1.0)
    if not fills or fills.issubset(AUTHENTIC_FILL_HEXES):
        palette_score += 6.0
    else:
        matched_cnt = len(fills & AUTHENTIC_FILL_HEXES)
        palette_score += round(6.0 * (matched_cnt / max(len(fills), 1)), 1)

    # 3. 표 및 여백/머리글 위생 (/25)
    t_cnt = int(profile.get("tables_count", 0))
    b_cnt = int(profile.get("bordered_tables_count", t_cnt))
    p_cnt = int(profile.get("padded_tables_count", t_cnt))
    if t_cnt > 0:
        b_ratio = b_cnt / t_cnt
        p_ratio = p_cnt / t_cnt
        table_score = round(8.0 * min(b_ratio, 1.0) + 7.0 * min(p_ratio, 1.0), 1)
    else:
        table_score = 8.0  # 순수 텍스트/학술 문서

    margins = profile.get("margins_pt", {})
    m_lo, m_hi = spec["margin_range"]
    if margins:
        l_mar = float(margins.get("left", 72.0))
        margin_score = 5.0 if m_lo <= l_mar <= m_hi else 2.0
    else:
        margin_score = 5.0
    hf_score = 5.0 if profile.get("has_header_footer", False) else 2.5
    layout_score = table_score + margin_score + hf_score

    # 4. 계층 들여쓰기 및 양식 구조 (/15)
    space_faked = int(profile.get("space_faked_indent_hits", 0))
    hier_score = 8.0 if space_faked == 0 else 0.0
    sdt_total = int(profile.get("sdt_total", 0))
    sdt_in_tbl = int(profile.get("sdt_in_table", 0))
    if sdt_total > 0:
        hier_score += round(7.0 * (sdt_in_tbl / sdt_total), 1)
    elif len([x for x in profile.get("paragraphs", []) if str(x).strip()]) >= 3:
        hier_score += 7.0
    else:
        hier_score += 4.0

    # 5. 실무 관용 표현 및 AI 이모지 배제 (/10)
    emojis = int(profile.get("emoji_hits", 0))
    cues = int(profile.get("corporate_cues_count", 0))
    idiom_score = (5.0 if emojis == 0 else 0.0) + (5.0 if cues >= 2 else (3.5 if cues == 1 or len(profile.get("paragraphs", [])) >= 3 else 2.0))

    raw_score = round(min(100.0, typo_score + palette_score + layout_score + hier_score + idiom_score), 1)
    ai_tells = detect_ai_tells(profile)
    penalty = sum(15.0 if t["severity"] == "critical" else 8.0 for t in ai_tells)
    final_score = round(max(0.0, raw_score - penalty), 1)

    if ai_tells:
        verdict = "AI_TELL_DETECTED"
    elif final_score >= 80.0:
        verdict = "AUTHENTIC_PASS"
    else:
        verdict = "LOW_SIMILARITY"

    return {
        "archetype": resolved_arch,
        "archetype_label": spec["label"],
        "raw_similarity_score": raw_score,
        "ai_tell_penalty": penalty,
        "similarity_score": final_score,
        "verdict": verdict,
        "passed": verdict == "AUTHENTIC_PASS",
        "ai_tells": ai_tells,
        "subscores": {
            "typography": typo_score,
            "palette_contrast": palette_score,
            "table_layout": layout_score,
            "hierarchy_structure": hier_score,
            "idiom_authenticity": idiom_score,
        },
    }


def _profile_feature_vector(prof: dict[str, Any]) -> list[float]:
    t_cnt = max(int(prof.get("tables_count", 0)), 1)
    margins = prof.get("margins_pt", {})
    m_left = float(margins.get("left", 56.7))
    return [
        min(float(prof.get("median_body_pt", 10.5)) / 12.0, 1.5),
        min(m_left / 85.0, 1.5),
        min(int(prof.get("tables_count", 0)) / 5.0, 1.2),
        float(prof.get("dark_body_ratio", 1.0)),
        min(int(prof.get("bordered_tables_count", 0)) / t_cnt, 1.0) if prof.get("tables_count", 0) else 0.8,
        min(int(prof.get("padded_tables_count", 0)) / t_cnt, 1.0) if prof.get("tables_count", 0) else 0.8,
        1.0 if int(prof.get("emoji_hits", 0)) == 0 else 0.0,
        1.0 if int(prof.get("space_faked_indent_hits", 0)) == 0 else 0.0,
        1.0 if not prof.get("uninstalled_fonts") else 0.0,
    ]


def _cosine_similarity_pct(v1: list[float], v2: list[float]) -> float:
    dot = sum(a * b for a, b in zip(v1, v2))
    n1 = math.sqrt(sum(a * a for a in v1))
    n2 = math.sqrt(sum(b * b for b in v2))
    if n1 == 0.0 or n2 == 0.0:
        return 0.0
    return round(100.0 * (dot / (n1 * n2)), 1)


def compare_real_vs_generated(
    real_path: str | Path,
    generated: str | Path | dict[str, Any],
    archetype: str = "auto",
) -> dict[str, Any]:
    """실제 레퍼런스 원본 DOCX vs COM 생성 문서를 블라인드 비교해 진위(Genuine vs AI)와 구조 유사도를 판정."""
    real_prof = extract_docx_style_profile(real_path)
    if isinstance(generated, (str, Path)):
        gen_prof = extract_docx_style_profile(generated)
    elif "is_minimal_mock" in generated and "neon_low_contrast_count" in generated:
        gen_prof = generated
    else:
        gen_prof = profile_from_word_snapshot(generated)

    resolved_arch = archetype if archetype != "auto" else infer_archetype(real_prof)
    real_eval = evaluate_docx_authenticity(real_prof, archetype=resolved_arch)
    gen_eval = evaluate_docx_authenticity(gen_prof, archetype=resolved_arch)

    vec_real = _profile_feature_vector(real_prof)
    vec_gen = _profile_feature_vector(gen_prof)
    cosine_pct = _cosine_similarity_pct(vec_real, vec_gen)

    # 팔레트 겹침도 반영
    r_fills = set(real_prof.get("fills", []))
    g_fills = set(gen_prof.get("fills", []))
    if r_fills or g_fills:
        # 둘 다 공식 팔레트에 속하면 구조 동질성 가점
        both_auth = r_fills.issubset(AUTHENTIC_FILL_HEXES) and g_fills.issubset(AUTHENTIC_FILL_HEXES)
        pal_sim = 100.0 if both_auth else round(100.0 * len(r_fills & g_fills) / max(len(r_fills | g_fills), 1), 1)
    else:
        pal_sim = 100.0

    structural_similarity_pct = round(0.75 * cosine_pct + 0.25 * pal_sim - gen_eval["ai_tell_penalty"] * 0.5, 1)
    structural_similarity_pct = max(0.0, min(100.0, structural_similarity_pct))

    pair_delta = round(gen_eval["similarity_score"] - real_eval["similarity_score"], 1)
    passed = (
        gen_eval["verdict"] == "AUTHENTIC_PASS"
        and structural_similarity_pct >= 80.0
        and len(gen_eval["ai_tells"]) == 0
    )

    return {
        "archetype": resolved_arch,
        "real_source": str(real_path),
        "generated_source": str(gen_prof.get("source", "live_com_word")),
        "real_authenticity_score": real_eval["similarity_score"],
        "generated_authenticity_score": gen_eval["similarity_score"],
        "structural_similarity_pct": structural_similarity_pct,
        "pair_delta": pair_delta,
        "ai_tells": gen_eval["ai_tells"],
        "comparison_verdict": "INDISTINGUISHABLE" if passed else "NEEDS_REFINEMENT",
        "fake_detector_label": "REAL_GRADE_PARITY" if passed else "DETECTED_AI_SYNTHETIC",
        "passed": passed,
        "real_eval": real_eval,
        "generated_eval": gen_eval,
    }
