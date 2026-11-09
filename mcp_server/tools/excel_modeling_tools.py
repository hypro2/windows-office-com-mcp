# -*- coding: utf-8 -*-
"""
mcp_server/tools/excel_modeling_tools.py
iOfficeAI(OfficeCLI) officecli-financial-model & officecli-data-dashboard 규격 기반 Excel 실시간 COM 도구:
1. 3-Zone 재무 모델 스캐폴더 (Inputs 노랑 #FFC000 / Calc 파랑 #4472C4 / Outputs 초록 #70AD47 + 대차균형 & 현금검증)
2. CFO 표준 5색 자동 코딩 & Calc 시트 하드코딩 숫자 감사 (`FinancialModelClassifier`)
3. 데이터 행 수 적응형 대시보드 예산·열너비 계산기 (`DashboardBudgetCalculator`)
"""

from __future__ import annotations

import math
import re
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr
from .excel_tools import _normalize_com_value, excel_freeze_display

INPUT_SHEETS = {"assumptions", "inputs", "drivers"}
CALC_SHEETS = {"p&l", "balance sheet", "cash flow", "fcf", "wacc", "dcf", "debt", "arr", "s&u"}
OUTPUT_SHEETS = {"summary", "dashboard", "sensitivity", "returns", "exit"}

ZONE_TAB_COLORS = {
    "Inputs": "#FFC000",   # 노랑
    "Calc": "#4472C4",     # 파랑
    "Outputs": "#70AD47",  # 초록
}

CFO_FONT_COLORS = {
    "numeric_input": "#0000FF",       # 파란 글씨: 하드코딩 입력 상수
    "same_sheet_formula": "#000000",  # 검은 글씨: 동일 시트 수식
    "cross_sheet": "#008000",         # 초록 글씨: 시트 간 참조 (!)
    "external_link": "#FF0000",       # 빨간 글씨: 외부 워크북 링크 ([)
}

_QUOTED_STR_RE = re.compile(r'"[^"]*"')


def _col_letter(col_idx: int) -> str:
    """1-based 열 번호 -> Excel 열 문자 (1->A, 27->AA)."""
    out = []
    n = int(col_idx)
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out)) or "A"


class FinancialModelClassifier:
    """3-Zone 분류 및 CFO 표준 5색 코딩/하드코딩 감사 순수 로직."""

    @staticmethod
    def classify_zone(sheet_name: str) -> str:
        low = (sheet_name or "").strip().lower()
        if low in INPUT_SHEETS:
            return "Inputs"
        if low in CALC_SHEETS:
            return "Calc"
        if low in OUTPUT_SHEETS:
            return "Outputs"
        return "Other"

    @staticmethod
    def classify_cell(formula: Any, value: Any) -> str:
        f_str = str(formula or "").strip()
        if f_str.startswith("="):
            # 따옴표 내부 리터럴 문자열(예: "CF != BS CASH")의 '!'나 '[' 오탐 방지
            stripped = _QUOTED_STR_RE.sub("", f_str)
            if "[" in stripped:
                return "external_link"
            if "!" in stripped:
                return "cross_sheet"
            return "same_sheet_formula"

        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return "numeric_input"
        return "text_or_empty"

    @classmethod
    def audit_calc_hardcodes_2d(
        cls,
        sheet_name: str,
        formulas_2d: list[list[Any]],
        values_2d: list[list[Any]],
        start_row: int = 1,
        start_col: int = 1,
    ) -> list[dict[str, Any]]:
        """Calc Zone 시트 내 수식이 아닌 하드코딩 숫자 상수 적발."""
        if cls.classify_zone(sheet_name) != "Calc":
            return []
        violations: list[dict[str, Any]] = []
        for r_off, row_vals in enumerate(values_2d):
            f_row = formulas_2d[r_off] if r_off < len(formulas_2d) else []
            for c_off, val in enumerate(row_vals):
                f_val = f_row[c_off] if c_off < len(f_row) else ""
                if cls.classify_cell(f_val, val) == "numeric_input":
                    cell_addr = f"{_col_letter(start_col + c_off)}{start_row + r_off}"
                    violations.append(
                        {
                            "sheet": sheet_name,
                            "cell": cell_addr,
                            "value": val,
                            "message": f"Calc 시트 '{sheet_name}!{cell_addr}'에 하드코딩 숫자({val})가 있습니다. Assumptions 시트로 분리하세요.",
                        }
                    )
        return violations


class DashboardBudgetCalculator:
    """iOfficeAI officecli-data-dashboard 규격 기반 스케일링 & 열너비/차트폭 계산기."""

    @staticmethod
    def scale_complexity(row_count: int) -> dict[str, Any]:
        n = max(0, int(row_count))
        if n < 10:
            return {
                "tier": "<10",
                "kpis": "1-2",
                "recommended_kpis": 2,
                "charts": "1",
                "recommended_charts": 1,
                "sparklines": "skip",
                "cf_rules": "0-1",
                "recommended_cf_rules": 0,
                "preset": "minimal",
            }
        if n < 50:
            return {
                "tier": "10-50",
                "kpis": "2-3",
                "recommended_kpis": 3,
                "charts": "2",
                "recommended_charts": 2,
                "sparklines": "if_time_series",
                "cf_rules": "1-2",
                "recommended_cf_rules": 1,
                "preset": "dashboard",
            }
        if n < 200:
            return {
                "tier": "50-200",
                "kpis": "3-5",
                "recommended_kpis": 4,
                "charts": "2-3",
                "recommended_charts": 3,
                "sparklines": "if_time_series",
                "cf_rules": "2-3",
                "recommended_cf_rules": 2,
                "preset": "dashboard",
            }
        return {
            "tier": "200+",
            "kpis": "3-5",
            "recommended_kpis": 5,
            "charts": "3",
            "recommended_charts": 3,
            "sparklines": "if_time_series",
            "cf_rules": "3-4",
            "recommended_cf_rules": 3,
            "preset": "dashboard",
        }

    @staticmethod
    def kpi_column_width(sample_value: str | int, min_width: int = 22) -> int:
        """24pt 대형 KPI 숫자 `###` 방지 공식: max(22, ceil((visible_chars + 2) * 1.3))."""
        visible_chars = int(sample_value) if isinstance(sample_value, int) else len(str(sample_value or ""))
        return max(min_width, math.ceil((visible_chars + 2) * 1.3))

    @staticmethod
    def chart_column_span(title: str) -> int:
        """차트 제목 잘림 방지 최소 열 수 공식: ceil(len(title) * 0.18)."""
        return math.ceil(len(str(title or "")) * 0.18)


def _bgr_luminance(bgr_val: Any) -> float:
    """COM COLORREF (BGR) 정수 밝기(0.0~1.0) 반환. 판단 불가 시 1.0(밝음)."""
    try:
        if bgr_val is None or isinstance(bgr_val, bool):
            return 1.0
        val = int(bgr_val) & 0xFFFFFF
        r, g, b = val & 0xFF, (val >> 8) & 0xFF, (val >> 16) & 0xFF
        return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0
    except Exception:
        return 1.0


def _ensure_sheet(wb: Any, name: str, tab_hex: str) -> Any:
    """워크북 내 시트를 찾거나 생성하고 탭 색상을 설정 (생성 시 항상 맨 뒤 추가)."""
    for i in range(1, wb.Sheets.Count + 1):
        sh = wb.Sheets(i)
        if str(sh.Name).strip().lower() == name.strip().lower():
            sh.Name = name
            try:
                sh.Tab.Color = hex_to_bgr(tab_hex)
            except Exception:
                pass
            return sh
    sh = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
    sh.Name = name
    try:
        sh.Tab.Color = hex_to_bgr(tab_hex)
    except Exception:
        pass
    return sh


def _style_ib_sheet(
    excel_app: Any,
    sh: Any,
    end_col: str,
    max_data_row: int,
    subtotal_rows: tuple[int, ...] = (),
    total_rows: tuple[int, ...] = (),
    check_rows: tuple[int, ...] = (),
    print_area: str = "",
    landscape: bool = False,
) -> None:
    """IB 표준 재무제표 타이포그래피·합계 이중선·Check 행·여백/인쇄 영역 일괄 정돈."""
    try:
        sh.Range(f"A1:{end_col}{max(max_data_row, 2)}").Font.Name = "맑은 고딕"
    except Exception:
        pass
    try:
        t_cell = sh.Range("A1")
        t_cell.Font.Size = 13
        t_cell.Font.Bold = True
        t_cell.Font.Color = hex_to_bgr("#1F3864")
        t_cell.VerticalAlignment = -4108
        sh.Rows(1).RowHeight = 26
        sh.Rows(2).RowHeight = 22
        sh.Range(f"A2:{end_col}2").Font.Size = 10
        sh.Range(f"A2:{end_col}2").VerticalAlignment = -4108
    except Exception:
        pass

    for r in range(3, max_data_row + 1):
        try:
            sh.Rows(r).RowHeight = 20
            rng = sh.Range(f"A{r}:{end_col}{r}")
            rng.Font.Size = 10
            rng.VerticalAlignment = -4108
            b_btm = rng.Borders(9)
            b_btm.LineStyle = 1
            b_btm.Weight = 1  # xlHairline
            b_btm.Color = hex_to_bgr("#E2E8F0")
        except Exception:
            pass

    for r in subtotal_rows:
        try:
            rng = sh.Range(f"A{r}:{end_col}{r}")
            rng.Font.Bold = True
            rng.Interior.Color = hex_to_bgr("#F8FAFC")
            b_top = rng.Borders(8)
            b_top.LineStyle = 1
            b_top.Weight = 2
            b_top.Color = hex_to_bgr("#94A3B8")
            b_btm = rng.Borders(9)
            b_btm.LineStyle = 1
            b_btm.Weight = 2
            b_btm.Color = hex_to_bgr("#CBD5E1")
        except Exception:
            pass

    for r in total_rows:
        try:
            rng = sh.Range(f"A{r}:{end_col}{r}")
            rng.Font.Bold = True
            rng.Interior.Color = hex_to_bgr("#EFF6FF")
            b_top = rng.Borders(8)
            b_top.LineStyle = 1
            b_top.Weight = 2
            b_top.Color = hex_to_bgr("#1F3864")
            b_btm = rng.Borders(9)
            b_btm.LineStyle = -4119  # xlDouble
            b_btm.Color = hex_to_bgr("#1F3864")
            sh.Rows(r).RowHeight = 22
        except Exception:
            pass

    for r in check_rows:
        try:
            rng = sh.Range(f"A{r}:{end_col}{r}")
            rng.Font.Bold = True
            rng.Interior.Color = hex_to_bgr("#ECFDF5")
            sh.Range(f"B{r}:{end_col}{r}").HorizontalAlignment = -4152  # xlRight
            b_top = rng.Borders(8)
            b_top.LineStyle = 1
            b_top.Weight = 2
            b_top.Color = hex_to_bgr("#A7F3D0")
            b_btm = rng.Borders(9)
            b_btm.LineStyle = 1
            b_btm.Weight = 2
            b_btm.Color = hex_to_bgr("#10B981")
            sh.Rows(r).RowHeight = 22
        except Exception:
            pass

    try:
        ps = sh.PageSetup
        if print_area:
            ps.PrintArea = print_area
        margin_pt = float(excel_app.InchesToPoints(0.4))
        ps.LeftMargin = margin_pt
        ps.RightMargin = margin_pt
        ps.TopMargin = margin_pt
        ps.BottomMargin = margin_pt
        ps.CenterHorizontally = True
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        if landscape:
            ps.Orientation = 2  # xlLandscape
    except Exception:
        pass


def _freeze_cell(excel_app: Any, sheet: Any, row_idx: int, col_idx: int) -> None:
    """시트의 지정 셀 기준(SplitRow/SplitColumn) 틀 고정."""
    try:
        sheet.Activate()
        win = excel_app.ActiveWindow
        win.FreezePanes = False
        win.SplitRow = max(0, row_idx - 1)
        win.SplitColumn = max(0, col_idx - 1)
        win.FreezePanes = True
    except Exception:
        pass


def register_excel_modeling_tools(mcp: MCPServer) -> None:
    """Excel 3-Zone 재무 모델링 및 적응형 대시보드 MCP 도구 등록."""

    @mcp.tool()
    def excel_scaffold_financial_model(
        model_type: Literal["3-statement", "dcf", "lbo"] = "3-statement",
        start_year: int = 2024,
        years: int = 4,
        apply_sample_formulas: bool = True,
    ) -> dict[str, Any]:
        """
        iOfficeAI 3-Zone 아키텍처(Inputs 노랑 / Calc 파랑 / Outputs 초록) 기반으로
        재무 모델 통합문서 스켈레톤 및 살아있는 수식을 현재 Excel 창에 즉시 구축합니다.

        - `3-statement`: Assumptions -> P&L -> Balance Sheet(대차균형 검증 포함) -> Cash Flow(현금대조 검증 포함) -> Summary
        - `dcf`: Assumptions -> FCF -> WACC -> DCF -> Sensitivity
        - `lbo`: Assumptions -> S&U -> Debt -> P&L -> Cash Flow -> Returns
        """
        ensure_com()
        excel_app, wb, _ = com_manager.get_excel(create_if_missing=True, visible=True)
        if excel_app is None or wb is None:
            return {"success": False, "error": "Excel 연결 실패"}

        years_cnt = max(2, min(10, int(years)))
        year_labels = [f"{start_year}A"] + [f"{start_year + i}E" for i in range(1, years_cnt)]

        zone_map = {
            "3-statement": [
                ("Assumptions", "Inputs", (2, 2)),
                ("P&L", "Calc", (3, 2)),
                ("Balance Sheet", "Calc", (3, 2)),
                ("Cash Flow", "Calc", (3, 2)),
                ("Summary", "Outputs", (2, 2)),
            ],
            "dcf": [
                ("Assumptions", "Inputs", (2, 2)),
                ("FCF", "Calc", (3, 2)),
                ("WACC", "Calc", (3, 2)),
                ("DCF", "Calc", (3, 2)),
                ("Sensitivity", "Outputs", (2, 2)),
            ],
            "lbo": [
                ("Assumptions", "Inputs", (2, 2)),
                ("S&U", "Calc", (3, 2)),
                ("Debt", "Calc", (3, 2)),
                ("P&L", "Calc", (3, 2)),
                ("Cash Flow", "Calc", (3, 2)),
                ("Returns", "Outputs", (2, 2)),
            ],
        }
        plan = zone_map[model_type]

        try:
            with excel_freeze_display(excel_app, pause_calculation=True):
                created_sheets: dict[str, Any] = {}
                for sh_name, zone, _freeze in plan:
                    sh = _ensure_sheet(wb, sh_name, ZONE_TAB_COLORS[zone])
                    created_sheets[sh_name] = sh

                # 빈 기본 시트(Sheet1 등) 제거 및 plan 순서대로 시트 탭 정확히 정렬
                plan_names_low = {p[0].lower() for p in plan}
                for sh_i in range(wb.Sheets.Count, 0, -1):
                    cand = wb.Sheets(sh_i)
                    if str(cand.Name).strip().lower() not in plan_names_low and wb.Sheets.Count > len(plan):
                        try:
                            if int(excel_app.WorksheetFunction.CountA(cand.UsedRange)) == 0:
                                cand.Delete()
                        except Exception:
                            pass
                for idx_pos, (sh_name, _zone, _freeze) in enumerate(plan):
                    sh_obj = created_sheets[sh_name]
                    try:
                        if idx_pos == 0:
                            sh_obj.Move(wb.Sheets(1))
                        else:
                            sh_obj.Move(None, created_sheets[plan[idx_pos - 1][0]])
                    except Exception:
                        pass

                if model_type == "3-statement" and apply_sample_formulas:
                    end_col = _col_letter(1 + years_cnt)
                    sh_a = created_sheets["Assumptions"]
                    sh_a.Range("A1").Value = "Key Model Assumptions (Inputs Zone)"
                    sh_a.Range(f"A2:{end_col}2").Value = [["Driver"] + year_labels]
                    sh_a.Range("A3:B3").Value = [["Base Revenue ($K)", 10000]]
                    sh_a.Range("A4").Value = "Revenue Growth"
                    sh_a.Range("A5").Value = "Gross Margin"
                    sh_a.Range("A6").Value = "OpEx % of Revenue"
                    sh_a.Range("A7").Value = "Tax Rate"
                    sh_a.Range("A8").Value = "D&A % of Revenue"
                    sh_a.Range("A9").Value = "CapEx % of Revenue"
                    sh_a.Range("A10:B10").Value = [["Y1 Opening Cash ($K)", 2000]]
                    sh_a.Range("A11").Value = "Paid-In Capital ($K)"
                    sh_a.Range("A12").Value = "Total Debt ($K)"
                    sh_a.Range("A13:B13").Value = [["Y1 Opening Net PP&E ($K)", 4000]]

                    for c_i in range(2, 2 + years_cnt):
                        col = _col_letter(c_i)
                        sh_a.Range(f"{col}4").Value = 0.0 if c_i == 2 else 0.12
                        sh_a.Range(f"{col}5").Value = 0.60
                        sh_a.Range(f"{col}6").Value = 0.32
                        sh_a.Range(f"{col}7").Value = 0.22
                        sh_a.Range(f"{col}8").Value = 0.04
                        sh_a.Range(f"{col}9").Value = 0.05
                        sh_a.Range(f"{col}11").Value = 5000
                        sh_a.Range(f"{col}12").Value = 1000
                    # 입력 셀 일관된 소프트 크림(#FEF9C3) 음영 + 파란 글씨 + 퍼센트/천단위 콤마 서식
                    for inp_rng in ("B3", f"B4:{end_col}9", "B10", f"B11:{end_col}12", "B13"):
                        sh_a.Range(inp_rng).Interior.Color = hex_to_bgr("#FEF9C3")
                    sh_a.Range(f"B3:{end_col}13").Font.Color = hex_to_bgr(CFO_FONT_COLORS["numeric_input"])
                    sh_a.Range(f"B4:{end_col}9").NumberFormat = "0.0%"
                    for amt_rng in ("B3", "B10", f"B11:{end_col}12", "B13"):
                        sh_a.Range(amt_rng).NumberFormat = '#,##0;(#,##0);"-"'
                    sh_a.Range("A:A").ColumnWidth = 28

                    # P&L (Calc Zone)
                    sh_pl = created_sheets["P&L"]
                    sh_pl.Range("A1").Value = "Income Statement (P&L)"
                    sh_pl.Range(f"A2:{end_col}2").Value = [["Line Item"] + year_labels]
                    pl_labels = [
                        (3, "Revenue"), (4, "COGS"), (5, "Gross Profit"), (6, "OpEx"),
                        (7, "EBITDA"), (8, "D&A"), (9, "EBIT"), (10, "EBT"),
                        (11, "Tax"), (12, "Net Income"),
                    ]
                    for r_i, lbl in pl_labels:
                        sh_pl.Range(f"A{r_i}").Value = lbl
                    for idx_c in range(years_cnt):
                        c_num = 2 + idx_c
                        col = _col_letter(c_num)
                        prev_col = _col_letter(c_num - 1)
                        sh_pl.Range(f"{col}3").Formula = "=Assumptions!B3" if idx_c == 0 else f"={prev_col}3*(1+Assumptions!{col}4)"
                        sh_pl.Range(f"{col}4").Formula = f"=-{col}3*(1-Assumptions!{col}5)"
                        sh_pl.Range(f"{col}5").Formula = f"={col}3+{col}4"
                        sh_pl.Range(f"{col}6").Formula = f"=-{col}3*Assumptions!{col}6"
                        sh_pl.Range(f"{col}7").Formula = f"={col}5+{col}6"
                        sh_pl.Range(f"{col}8").Formula = f"=-{col}3*Assumptions!{col}8"
                        sh_pl.Range(f"{col}9").Formula = f"={col}7+{col}8"
                        sh_pl.Range(f"{col}10").Formula = f"={col}9"
                        sh_pl.Range(f"{col}11").Formula = f"=-MAX(0,{col}10*Assumptions!{col}7)"
                        sh_pl.Range(f"{col}12").Formula = f"={col}10+{col}11"
                    sh_pl.Range(f"B3:{end_col}12").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_pl.Range("A:A").ColumnWidth = 26

                    # Cash Flow (Calc Zone) - Net PP&E 및 BS Cash가 참조할 행 구성
                    sh_cf = created_sheets["Cash Flow"]
                    sh_cf.Range("A1").Value = "Cash Flow Statement"
                    sh_cf.Range(f"A2:{end_col}2").Value = [["Line Item"] + year_labels]
                    cf_labels = [
                        (3, "Opening Cash"), (4, "Net Income"), (5, "Plus: D&A"),
                        (6, "Operating Cash Flow"), (7, "CapEx (Investing CF)"),
                        (8, "Financing Cash Flow"), (9, "Ending Cash"),
                        (11, "CF - BS Cash Recon"),
                    ]
                    for r_i, lbl in cf_labels:
                        sh_cf.Range(f"A{r_i}").Value = lbl
                    for idx_c in range(years_cnt):
                        c_num = 2 + idx_c
                        col = _col_letter(c_num)
                        prev_col = _col_letter(c_num - 1)
                        sh_cf.Range(f"{col}3").Formula = "=Assumptions!B10" if idx_c == 0 else f"={prev_col}9"
                        sh_cf.Range(f"{col}4").Formula = f"='P&L'!{col}12"
                        sh_cf.Range(f"{col}5").Formula = f"=-'P&L'!{col}8"
                        sh_cf.Range(f"{col}6").Formula = f"={col}4+{col}5"
                        sh_cf.Range(f"{col}7").Formula = f"=-'P&L'!{col}3*Assumptions!{col}9"
                        sh_cf.Range(f"{col}8").Formula = f"=Assumptions!{col}12-Assumptions!{col}12"
                        sh_cf.Range(f"{col}9").Formula = f"={col}3+{col}6+{col}7+{col}8"
                        sh_cf.Range(f"{col}11").Formula = f'=IF(ABS({col}9-\'Balance Sheet\'!{col}5)<0.01,"OK","CF != BS CASH: "&ROUND({col}9-\'Balance Sheet\'!{col}5,0))'
                    sh_cf.Range(f"B3:{end_col}9").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_cf.Range("A:A").ColumnWidth = 26

                    # Balance Sheet (Calc Zone)
                    sh_bs = created_sheets["Balance Sheet"]
                    sh_bs.Range("A1").Value = "Balance Sheet"
                    sh_bs.Range(f"A2:{end_col}2").Value = [["Line Item"] + year_labels]
                    bs_labels = [
                        (5, "Cash"), (6, "Net PP&E"), (10, "Total Assets"),
                        (15, "Total Liabilities (Debt)"), (16, "Paid-In Capital"),
                        (17, "Retained Earnings"), (18, "Total Equity"), (19, "Balance Check"),
                    ]
                    for r_i, lbl in bs_labels:
                        sh_bs.Range(f"A{r_i}").Value = lbl
                    for idx_c in range(years_cnt):
                        c_num = 2 + idx_c
                        col = _col_letter(c_num)
                        prev_col = _col_letter(c_num - 1)
                        sh_bs.Range(f"{col}5").Formula = f"='Cash Flow'!{col}9"
                        sh_bs.Range(f"{col}6").Formula = (
                            f"=Assumptions!B13-'Cash Flow'!{col}7-'Cash Flow'!{col}5"
                            if idx_c == 0
                            else f"={prev_col}6-'Cash Flow'!{col}7-'Cash Flow'!{col}5"
                        )
                        sh_bs.Range(f"{col}10").Formula = f"=SUM({col}5:{col}6)"
                        sh_bs.Range(f"{col}15").Formula = f"=Assumptions!{col}12"
                        sh_bs.Range(f"{col}16").Formula = f"=Assumptions!{col}11"
                        sh_bs.Range(f"{col}17").Formula = (
                            f"='P&L'!{col}12" if idx_c == 0 else f"={prev_col}17+'P&L'!{col}12"
                        )
                        sh_bs.Range(f"{col}18").Formula = f"=SUM({col}16:{col}17)"
                        sh_bs.Range(f"{col}19").Formula = f'=IF(ABS({col}10-{col}15-{col}18)<0.01,"OK","IMBALANCED: "&ROUND({col}10-{col}15-{col}18,0))'
                    sh_bs.Range(f"B5:{end_col}18").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_bs.Range("A:A").ColumnWidth = 26

                    # Summary (Outputs Zone)
                    sh_sum = created_sheets["Summary"]
                    sh_sum.Range("A1").Value = "Executive Financial Summary (Outputs Zone)"
                    sh_sum.Range(f"A2:{end_col}2").Value = [["Metric"] + year_labels]
                    sh_sum.Range("A3").Value = "Revenue"
                    sh_sum.Range("A4").Value = "EBITDA"
                    sh_sum.Range("A5").Value = "Net Income"
                    sh_sum.Range("A6").Value = "Ending Cash"
                    sh_sum.Range("A7").Value = "BS Balance Status"
                    for idx_c in range(years_cnt):
                        col = _col_letter(2 + idx_c)
                        sh_sum.Range(f"{col}3").Formula = f"='P&L'!{col}3"
                        sh_sum.Range(f"{col}4").Formula = f"='P&L'!{col}7"
                        sh_sum.Range(f"{col}5").Formula = f"='P&L'!{col}12"
                        sh_sum.Range(f"{col}6").Formula = f"='Cash Flow'!{col}9"
                        sh_sum.Range(f"{col}7").Formula = f"='Balance Sheet'!{col}19"
                    sh_sum.Range(f"B3:{end_col}6").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_sum.Range(f"B3:{end_col}7").Font.Color = hex_to_bgr(CFO_FONT_COLORS["cross_sheet"])
                    sh_sum.Range("A:A").ColumnWidth = 24
                    sh_sum.Range(f"B:{end_col}").ColumnWidth = 18
                    for sh_c in (sh_a, sh_pl, sh_cf, sh_bs):
                        sh_c.Range(f"B:{end_col}").ColumnWidth = 16
                    is_wide = years_cnt >= 5
                    _style_ib_sheet(excel_app, sh_a, end_col, 13, print_area=f"$A$1:${end_col}$14", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_pl, end_col, 12, subtotal_rows=(5, 7, 9), total_rows=(12,), print_area=f"$A$1:${end_col}$13", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_cf, end_col, 11, subtotal_rows=(6,), total_rows=(9,), check_rows=(11,), print_area=f"$A$1:${end_col}$12", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_bs, end_col, 19, subtotal_rows=(10, 15), total_rows=(18,), check_rows=(19,), print_area=f"$A$1:${end_col}$20", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_sum, end_col, 7, total_rows=(5, 6), check_rows=(7,), print_area=f"$A$1:${end_col}$9", landscape=is_wide)
                elif model_type == "dcf" and apply_sample_formulas:
                    end_col = _col_letter(1 + years_cnt)
                    sh_a = created_sheets["Assumptions"]
                    sh_a.Range("A1").Value = "Key DCF Assumptions (Inputs Zone)"
                    sh_a.Range(f"A2:{end_col}2").Value = [["Driver"] + year_labels]
                    a_rows = [
                        ["Base Revenue ($K)", 10000],
                        ["Revenue Growth"],
                        ["EBIT Margin"],
                        ["Cash Tax Rate"],
                        ["D&A % of Revenue"],
                        ["CapEx % of Revenue"],
                        ["Delta NWC % of Rev Delta"],
                        ["Terminal Growth Rate (TGR)", 0.025],
                        ["Exit EBITDA Multiple (x)", 10.0],
                        ["Net Debt ($K)", 3000],
                        ["Diluted Shares (K)", 1000],
                        ["Risk-Free Rate", 0.04],
                        ["Equity Risk Premium", 0.055],
                        ["Target Beta", 1.1],
                        ["Pre-Tax Cost of Debt", 0.06],
                        ["Target Equity Weight", 0.8],
                    ]
                    for r_off, row_vals in enumerate(a_rows, start=3):
                        sh_a.Range(f"A{r_off}").Value = row_vals[0]
                        if len(row_vals) > 1:
                            sh_a.Range(f"B{r_off}").Value = row_vals[1]
                    for c_i in range(2, 2 + years_cnt):
                        col = _col_letter(c_i)
                        sh_a.Range(f"{col}4:{col}9").Value = [
                            [0.0 if c_i == 2 else 0.12], [0.25], [0.22], [0.04], [0.05], [0.08]
                        ]
                    for inp_rng in (f"B4:{end_col}9", "B3", "B10:B18"):
                        sh_a.Range(inp_rng).Interior.Color = hex_to_bgr("#FEF9C3")
                    sh_a.Range(f"B3:{end_col}18").Font.Color = hex_to_bgr(CFO_FONT_COLORS["numeric_input"])
                    sh_a.Range(f"B4:{end_col}9").NumberFormat = "0.0%"
                    for r_pct in (10, 14, 15, 17, 18):
                        sh_a.Range(f"B{r_pct}").NumberFormat = "0.0%"
                    for r_num in (3, 12, 13):
                        sh_a.Range(f"B{r_num}").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_a.Range("B11").NumberFormat = '0.0"x"'
                    sh_a.Range("B16").NumberFormat = "0.00"
                    sh_a.Range("A:A").ColumnWidth = 28
                    sh_a.Range(f"B:{end_col}").ColumnWidth = 16

                    # WACC (Calc Zone)
                    sh_w = created_sheets["WACC"]
                    sh_w.Range("A1").Value = "Weighted Average Cost of Capital (WACC)"
                    sh_w.Range("A2:B2").Value = [["Parameter", "Value"]]
                    wacc_rows = [
                        ["Risk-Free Rate", "=Assumptions!B14"],
                        ["Equity Risk Premium", "=Assumptions!B15"],
                        ["Target Beta", "=Assumptions!B16"],
                        ["Cost of Equity (Ke)", "=B3+B5*B4"],
                        ["After-Tax Cost of Debt", "=Assumptions!B17*(1-Assumptions!B6)"],
                        ["Equity Weight (%)", "=Assumptions!B18"],
                        ["Debt Weight (%)", "=1-B8"],
                        ["Blended WACC", "=B6*B8+B7*B9"],
                    ]
                    sh_w.Range("A3:B10").Formula = wacc_rows
                    sh_w.Range("B3:B4").NumberFormat = "0.0%"
                    sh_w.Range("B5").NumberFormat = "0.00"
                    sh_w.Range("B6:B10").NumberFormat = "0.00%"
                    sh_w.Range("A:A").ColumnWidth = 30
                    sh_w.Range("B:B").ColumnWidth = 18

                    # FCF (Calc Zone)
                    sh_f = created_sheets["FCF"]
                    sh_f.Range("A1").Value = "Unlevered Free Cash Flow (UFCF)"
                    sh_f.Range(f"A2:{end_col}2").Value = [["Line Item"] + year_labels]
                    fcf_lbls = [
                        (3, "Revenue"), (4, "EBIT"), (5, "Less: Cash Taxes"), (6, "NOPAT"),
                        (7, "Plus: D&A"), (8, "Less: CapEx"), (9, "Less: Delta NWC"),
                        (10, "Unlevered FCF (UFCF)"), (11, "Discount Period (Yrs)"), (12, "Present Value of UFCF"),
                    ]
                    for r_i, lbl in fcf_lbls:
                        sh_f.Range(f"A{r_i}").Value = lbl
                    for idx_c in range(years_cnt):
                        col = _col_letter(2 + idx_c)
                        prev_col = _col_letter(1 + idx_c)
                        sh_f.Range(f"{col}3:{col}12").Formula = [
                            ["=Assumptions!B3" if idx_c == 0 else f"={prev_col}3*(1+Assumptions!{col}4)"],
                            [f"={col}3*Assumptions!{col}5"],
                            [f"=-MAX(0,{col}4*Assumptions!{col}6)"],
                            [f"={col}4+{col}5"],
                            [f"={col}3*Assumptions!{col}7"],
                            [f"=-{col}3*Assumptions!{col}8"],
                            [f"=-({col}3-Assumptions!B3)*Assumptions!{col}9" if idx_c == 0 else f"=-({col}3-{prev_col}3)*Assumptions!{col}9"],
                            [f"={col}6+{col}7+{col}8+{col}9"],
                            [f"=MAX(0,COLUMNS($B$2:{col}2)-1)"],
                            [f"=IF({col}11=0,0,{col}10/((1+WACC!$B$10)^{col}11))"],
                        ]
                    sh_f.Range(f"B3:{end_col}10").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_f.Range(f"B12:{end_col}12").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_f.Range("A:A").ColumnWidth = 26
                    sh_f.Range(f"B:{end_col}").ColumnWidth = 16

                    # DCF (Calc Zone)
                    sh_d = created_sheets["DCF"]
                    sh_d.Range("A1").Value = "Discounted Cash Flow Valuation"
                    sh_d.Range("A2:B2").Value = [["Valuation Component", "Value"]]
                    dcf_rows = [
                        ["Sum of PV of UFCF ($K)", f"=SUM(FCF!C12:{end_col}12)"],
                        ["Terminal Year UFCF ($K)", f"=FCF!{end_col}10"],
                        ["Terminal Value (Gordon, $K)", "=B4*(1+Assumptions!B10)/(WACC!$B$10-Assumptions!B10)"],
                        ["Terminal Discount Factor", f"=1/((1+WACC!$B$10)^FCF!{end_col}11)"],
                        ["PV of Terminal Value ($K)", "=B5*B6"],
                        ["Enterprise Value ($K)", "=B3+B7"],
                        ["Less: Net Debt ($K)", "=-Assumptions!B12"],
                        ["Implied Equity Value ($K)", "=B8+B9"],
                        ["Implied Share Price ($/sh)", "=B10/Assumptions!B13"],
                    ]
                    sh_d.Range("A3:B11").Formula = dcf_rows
                    sh_d.Range("B3:B10").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_d.Range("B6").NumberFormat = "0.000"
                    sh_d.Range("B11").NumberFormat = "$#,##0.00"
                    sh_d.Range("A:A").ColumnWidth = 30
                    sh_d.Range("B:B").ColumnWidth = 18

                    # Sensitivity (Outputs Zone)
                    sh_sens = created_sheets["Sensitivity"]
                    sh_sens.Range("A1").Value = "DCF Sensitivity Matrix: Implied Equity Value ($K) - WACC vs TGR"
                    sh_sens.Range("A2").Value = "WACC / TGR"
                    sh_sens.Range("B2:F2").Formula = [
                        ["=C2-0.005", "=D2-0.005", "=Assumptions!B10", "=D2+0.005", "=E2+0.005"]
                    ]
                    sh_sens.Range("B2:F2").NumberFormat = "0.0%"
                    sh_sens.Range("A3:A7").Formula = [
                        ["=A4-0.005"], ["=A5-0.005"], ["=WACC!$B$10"], ["=A5+0.005"], ["=A6+0.005"]
                    ]
                    sh_sens.Range("A3:A7").NumberFormat = "0.00%"
                    for r_s in range(3, 8):
                        row_f = [
                            f"=DCF!$B$3*(1+WACC!$B$10)/(1+$A{r_s})+(DCF!$B$4*(1+{_col_letter(c_s)}$2)/MAX(0.01,$A{r_s}-{_col_letter(c_s)}$2))/((1+$A{r_s})^FCF!${end_col}$11)-Assumptions!$B$12"
                            for c_s in range(2, 7)
                        ]
                        sh_sens.Range(f"B{r_s}:F{r_s}").Formula = [row_f]
                    sh_sens.Range("B3:F7").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_sens.Range("A:A").ColumnWidth = 22
                    sh_sens.Range("B:F").ColumnWidth = 18

                    is_wide = years_cnt >= 5
                    _style_ib_sheet(excel_app, sh_a, end_col, 18, print_area=f"$A$1:${end_col}$19", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_f, end_col, 12, subtotal_rows=(6, 10), total_rows=(12,), print_area=f"$A$1:${end_col}$13", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_w, "B", 10, subtotal_rows=(6, 7), total_rows=(10,), print_area="$A$1:$B$11", landscape=False)
                    _style_ib_sheet(excel_app, sh_d, "B", 11, subtotal_rows=(8,), total_rows=(10, 11), print_area="$A$1:$B$12", landscape=False)
                    _style_ib_sheet(excel_app, sh_sens, "F", 7, print_area="$A$1:$F$8", landscape=True)
                    try:
                        sh_sens.Range("A3:A7").Font.Bold = True
                        sh_sens.Range("A3:A7").Interior.Color = hex_to_bgr("#F1F5F9")
                        sh_sens.Range("A3:A7").HorizontalAlignment = -4108
                        sh_sens.Range("D5").Interior.Color = hex_to_bgr("#D1FAE5")
                        sh_sens.Range("D5").Font.Bold = True
                        for b_id in (7, 8, 9, 10, 11, 12):
                            b_obj = sh_sens.Range("A2:F7").Borders(b_id)
                            b_obj.LineStyle = 1
                            b_obj.Weight = 2
                            b_obj.Color = hex_to_bgr("#CBD5E1")
                    except Exception:
                        pass

                elif model_type == "lbo" and apply_sample_formulas:
                    end_col = _col_letter(1 + years_cnt)
                    sh_a = created_sheets["Assumptions"]
                    sh_a.Range("A1").Value = "Key LBO Assumptions (Inputs Zone)"
                    sh_a.Range(f"A2:{end_col}2").Value = [["Driver"] + year_labels]
                    lbo_a = [
                        ["Entry EBITDA ($K)", 2800],
                        ["Entry Multiple (x)", 8.0],
                        ["Exit Multiple (x)", 8.5],
                        ["Senior Debt % of EV", 0.6],
                        ["Transaction Fees % of EV", 0.02],
                        ["Debt Interest Rate", 0.07],
                        ["Mandatory Amort %", 0.05],
                        ["Base Revenue ($K)", 10000],
                    ]
                    for r_off, row_v in enumerate(lbo_a, start=3):
                        sh_a.Range(f"A{r_off}").Value = row_v[0]
                        sh_a.Range(f"B{r_off}").Value = row_v[1]
                    for r_off, lbl in ((11, "Revenue Growth"), (12, "EBITDA Margin"), (13, "D&A % of Rev"), (14, "CapEx % of Rev"), (15, "Tax Rate")):
                        sh_a.Range(f"A{r_off}").Value = lbl
                    for c_i in range(2, 2 + years_cnt):
                        col = _col_letter(c_i)
                        sh_a.Range(f"{col}11:{col}15").Value = [
                            [0.0 if c_i == 2 else 0.10], [0.28], [0.04], [0.04], [0.22]
                        ]
                    for inp_rng in ("B3:B10", f"B11:{end_col}15"):
                        sh_a.Range(inp_rng).Interior.Color = hex_to_bgr("#FEF9C3")
                    sh_a.Range(f"B3:{end_col}15").Font.Color = hex_to_bgr(CFO_FONT_COLORS["numeric_input"])
                    sh_a.Range("B3").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_a.Range("B10").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_a.Range("B4:B5").NumberFormat = '0.0"x"'
                    for rp in (6, 7, 8, 9):
                        sh_a.Range(f"B{rp}").NumberFormat = "0.0%"
                    sh_a.Range(f"B11:{end_col}15").NumberFormat = "0.0%"
                    sh_a.Range("A:A").ColumnWidth = 28
                    sh_a.Range(f"B:{end_col}").ColumnWidth = 16

                    # S&U (Calc Zone)
                    sh_su = created_sheets["S&U"]
                    sh_su.Range("A1").Value = "Sources & Uses of Funds"
                    sh_su.Range("A2:B2").Value = [["Item", "Amount ($K)"]]
                    su_rows = [
                        ["Entry Enterprise Value", "=Assumptions!B3*Assumptions!B4"],
                        ["Transaction Fees", "=B3*Assumptions!B7"],
                        ["Total Uses", "=B3+B4"],
                        ["", ""],
                        ["Senior Term Debt", "=B3*Assumptions!B6"],
                        ["Sponsor Equity (Plug)", "=B5-B7"],
                        ["Total Sources", "=B7+B8"],
                        ["", ""],
                        ["S&U Balance Check", '=IF(ABS(B9-B5)<0.01,"OK","IMBALANCED: "&ROUND(B9-B5,0))'],
                    ]
                    sh_su.Range("A3:B11").Formula = su_rows
                    sh_su.Range("B3:B9").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_su.Range("A:A").ColumnWidth = 28
                    sh_su.Range("B:B").ColumnWidth = 18

                    # Debt (Calc Zone - 먼저 P&L/CF 참조 셀 이름 확보 후 P&L/CF 채움)
                    sh_pl = created_sheets["P&L"]
                    sh_pl.Range("A1").Value = "LBO Income Statement (P&L)"
                    sh_pl.Range(f"A2:{end_col}2").Value = [["Line Item"] + year_labels]
                    for r_i, lbl in ((3, "Revenue"), (4, "EBITDA"), (5, "D&A"), (6, "EBIT"), (7, "Interest Expense"), (8, "EBT"), (9, "Taxes"), (10, "Net Income")):
                        sh_pl.Range(f"A{r_i}").Value = lbl
                    for idx_c in range(years_cnt):
                        col = _col_letter(2 + idx_c)
                        prev_col = _col_letter(1 + idx_c)
                        sh_pl.Range(f"{col}3:{col}10").Formula = [
                            ["=Assumptions!B10" if idx_c == 0 else f"={prev_col}3*(1+Assumptions!{col}11)"],
                            [f"={col}3*Assumptions!{col}12"],
                            [f"=-{col}3*Assumptions!{col}13"],
                            [f"={col}4+{col}5"],
                            [f"=-Debt!{col}5"],
                            [f"={col}6+{col}7"],
                            [f"=-MAX(0,{col}8*Assumptions!{col}15)"],
                            [f"={col}8+{col}9"],
                        ]
                    sh_pl.Range(f"B3:{end_col}10").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_pl.Range("A:A").ColumnWidth = 26
                    sh_pl.Range(f"B:{end_col}").ColumnWidth = 16

                    sh_cf = created_sheets["Cash Flow"]
                    sh_cf.Range("A1").Value = "LBO Levered Cash Flow & Debt Sweep"
                    sh_cf.Range(f"A2:{end_col}2").Value = [["Line Item"] + year_labels]
                    for r_i, lbl in ((3, "Net Income"), (4, "Plus: D&A"), (5, "Less: CapEx"), (6, "FCF Before Debt Sweep"), (7, "Mandatory + Sweep Debt Paydown"), (8, "Net Cash Flow"), (9, "Ending Cash")):
                        sh_cf.Range(f"A{r_i}").Value = lbl
                    for idx_c in range(years_cnt):
                        col = _col_letter(2 + idx_c)
                        prev_col = _col_letter(1 + idx_c)
                        sh_cf.Range(f"{col}3:{col}9").Formula = [
                            [f"='P&L'!{col}10"],
                            [f"=-'P&L'!{col}5"],
                            [f"=-'P&L'!{col}3*Assumptions!{col}14"],
                            ["='P&L'!B3*0" if idx_c == 0 else f"={col}3+{col}4+{col}5"],
                            [f"=Debt!{col}4"],
                            [f"={col}6+{col}7"],
                            [f"={col}8" if idx_c == 0 else f"={prev_col}9+{col}8"],
                        ]
                    sh_cf.Range(f"B3:{end_col}9").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_cf.Range("A:A").ColumnWidth = 30
                    sh_cf.Range(f"B:{end_col}").ColumnWidth = 16

                    sh_debt = created_sheets["Debt"]
                    sh_debt.Range("A1").Value = "Senior Debt Amortization Schedule"
                    sh_debt.Range(f"A2:{end_col}2").Value = [["Debt Tranche"] + year_labels]
                    for r_i, lbl in ((3, "Beginning Senior Debt"), (4, "Debt Paydown (Sweep)"), (5, "Interest Expense"), (6, "Ending Senior Debt")):
                        sh_debt.Range(f"A{r_i}").Value = lbl
                    for idx_c in range(years_cnt):
                        col = _col_letter(2 + idx_c)
                        prev_col = _col_letter(1 + idx_c)
                        sh_debt.Range(f"{col}3:{col}6").Formula = [
                            ["='S&U'!B7" if idx_c == 0 else f"={prev_col}6"],
                            ["='S&U'!B7*0" if idx_c == 0 else f"=-MIN({col}3,MAX('S&U'!$B$7*Assumptions!$B$9,'Cash Flow'!{col}6))"],
                            ["='S&U'!B7*0" if idx_c == 0 else f"={col}3*Assumptions!$B$8"],
                            [f"={col}3+{col}4"],
                        ]
                    sh_debt.Range(f"B3:{end_col}6").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_debt.Range("A:A").ColumnWidth = 28
                    sh_debt.Range(f"B:{end_col}").ColumnWidth = 16

                    sh_ret = created_sheets["Returns"]
                    sh_ret.Range("A1").Value = "Sponsor Returns Summary (Outputs Zone)"
                    sh_ret.Range("A2:B2").Value = [["Metric", "Value"]]
                    ret_rows = [
                        ["Exit Year EBITDA ($K)", f"='P&L'!{end_col}4"],
                        ["Exit Enterprise Value ($K)", "=B3*Assumptions!B5"],
                        ["Less: Exit Net Debt ($K)", f"=-(Debt!{end_col}6-'Cash Flow'!{end_col}9)"],
                        ["Exit Sponsor Equity ($K)", "=B4+B5"],
                        ["Initial Sponsor Equity ($K)", "='S&U'!B8"],
                        ["Multiple on Invested Capital (MOIC)", "=B6/B7"],
                        ["Implied Sponsor IRR (%)", f"=(B6/B7)^(1/MAX(1,COLUMNS('P&L'!B2:{end_col}2)-1))-1"],
                    ]
                    sh_ret.Range("A3:B9").Formula = ret_rows
                    sh_ret.Range("B3:B7").NumberFormat = '#,##0;(#,##0);"-"'
                    sh_ret.Range("B8").NumberFormat = '0.00"x"'
                    sh_ret.Range("B9").NumberFormat = "0.0%"
                    sh_ret.Range("A:A").ColumnWidth = 32
                    sh_ret.Range("B:B").ColumnWidth = 20

                    is_wide = years_cnt >= 5
                    _style_ib_sheet(excel_app, sh_a, end_col, 15, print_area=f"$A$1:${end_col}$16", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_su, "B", 11, total_rows=(5, 9), check_rows=(11,), print_area="$A$1:$B$12", landscape=False)
                    _style_ib_sheet(excel_app, sh_debt, end_col, 6, total_rows=(6,), print_area=f"$A$1:${end_col}$7", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_pl, end_col, 10, subtotal_rows=(4, 6), total_rows=(10,), print_area=f"$A$1:${end_col}$11", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_cf, end_col, 9, subtotal_rows=(6,), total_rows=(9,), print_area=f"$A$1:${end_col}$10", landscape=is_wide)
                    _style_ib_sheet(excel_app, sh_ret, "B", 9, subtotal_rows=(6,), total_rows=(8, 9), print_area="$A$1:$B$11", landscape=False)
                else:
                    # apply_sample_formulas=False 경량 스캐폴딩
                    for sh_name, sh in created_sheets.items():
                        sh.Range("A1").Value = f"{model_type.upper()} - {sh_name}"
                        sh.Range(f"A2:{_col_letter(1 + years_cnt)}2").Value = [["Item"] + year_labels]
                        sh.Range("A:A").ColumnWidth = 26
                        sh.Range(f"B:{_col_letter(1 + years_cnt)}").ColumnWidth = 16

                # 모든 시트 헤더 Fill + 틀 고정 + 인쇄 FitToPage + Outputs 활성화
                for sh_name, zone, (fr, fc) in plan:
                    sh = created_sheets[sh_name]
                    if apply_sample_formulas and model_type == "dcf" and sh_name == "Sensitivity":
                        end_c = "F"
                    elif apply_sample_formulas and (
                        (model_type == "dcf" and sh_name in ("WACC", "DCF"))
                        or (model_type == "lbo" and sh_name in ("S&U", "Returns"))
                    ):
                        end_c = "B"
                    else:
                        end_c = _col_letter(1 + years_cnt)
                    try:
                        hdr = sh.Range(f"A2:{end_c}2")
                        hdr.Interior.Color = hex_to_bgr("#1F3864")
                        hdr.Font.Color = hex_to_bgr("#FFFFFF")
                        hdr.Font.Bold = True
                        sh.Range("A2").HorizontalAlignment = -4131  # xlLeft
                        sh.Range(f"B2:{end_c}2").HorizontalAlignment = -4152  # xlRight
                        if not apply_sample_formulas:
                            sh.PageSetup.Zoom = False
                            sh.PageSetup.FitToPagesWide = 1
                            sh.PageSetup.FitToPagesTall = 1
                            if years_cnt >= 5:
                                sh.PageSetup.Orientation = 2  # xlLandscape
                    except Exception:
                        pass
                    _freeze_cell(excel_app, sh, fr, fc)

            # 계산 실행 및 Outputs 탭 활성화
            excel_app.Calculate()
            out_sheet_name = plan[-1][0]
            created_sheets[out_sheet_name].Activate()
            com_manager.activate_app("excel")

            return {
                "success": True,
                "workbook_name": wb.Name,
                "model_type": model_type,
                "sheets_created": [p[0] for p in plan],
                "active_output_sheet": out_sheet_name,
            }
        except Exception as e:
            return {"success": False, "error": f"3-Zone 재무 모델 스캐폴딩 실패: {e}"}

    @mcp.tool()
    def excel_apply_cfo_color_coding(
        sheet_name: str | None = None,
        highlight_assumption_rows: list[int] | None = None,
        audit_calc_hardcodes: bool = True,
    ) -> dict[str, Any]:
        """
        지정된 시트(생략 시 통합문서 전체 시트)에 CFO 표준 5색 코딩을 적용하고,
        Calc Zone 시트에 하드코딩된 숫자 상수가 있는지 감사(Audit)합니다:
        - 파랑(#0000FF): 하드코딩 숫자 입력
        - 검정(#000000): 동일 시트 수식
        - 초록(#008000): 시트 간 참조(!) 수식
        - 빨강(#FF0000): 외부 워크북 링크([) 수식
        """
        ensure_com()
        excel_app, wb, active_sheet = com_manager.get_excel(create_if_missing=False)
        if excel_app is None or wb is None:
            return {"success": False, "error": "열려 있는 Excel 워크북이 없습니다."}

        counts = {"numeric_input": 0, "same_sheet_formula": 0, "cross_sheet": 0, "external_link": 0}
        violations: list[dict[str, Any]] = []
        sheets_processed: list[str] = []

        try:
            target_sheets = [wb.Sheets(sheet_name)] if sheet_name else [wb.Sheets(i) for i in range(1, wb.Sheets.Count + 1)]
            with excel_freeze_display(excel_app, pause_calculation=False):
                for sh in target_sheets:
                    s_name = str(sh.Name)
                    sheets_processed.append(s_name)
                    used = sh.UsedRange
                    r_start = int(used.Row)
                    c_start = int(used.Column)
                    formulas_2d = _normalize_com_value(used.Formula)
                    values_2d = _normalize_com_value(used.Value)
                    color_buckets: dict[str, list[str]] = {k: [] for k in CFO_FONT_COLORS}

                    if audit_calc_hardcodes:
                        violations.extend(
                            FinancialModelClassifier.audit_calc_hardcodes_2d(
                                s_name, formulas_2d, values_2d, start_row=r_start, start_col=c_start
                            )
                        )

                    for r_off, row_vals in enumerate(values_2d):
                        r_abs = r_start + r_off
                        # 어두운 배경 행/셀(네이비 헤더, 코호트 다크 히트맵 등)의 흰 글씨 보호
                        row_fill = None
                        try:
                            last_col_let = _col_letter(c_start + max(0, len(row_vals) - 1))
                            row_rng = sh.Range(f"{_col_letter(c_start)}{r_abs}:{last_col_let}{r_abs}")
                            row_idx_val = row_rng.Interior.ColorIndex
                            if row_idx_val == -4142:
                                row_fill = 16777215  # xlColorIndexNone -> 밝은 배경 확정
                            elif row_idx_val is not None:
                                row_fill = row_rng.Interior.Color
                                if row_fill is not None and _bgr_luminance(row_fill) < 0.45:
                                    continue
                        except Exception:
                            pass

                        f_row = formulas_2d[r_off] if r_off < len(formulas_2d) else []
                        for c_off, val in enumerate(row_vals):
                            f_val = f_row[c_off] if c_off < len(f_row) else ""
                            cat = FinancialModelClassifier.classify_cell(f_val, val)
                            if cat in CFO_FONT_COLORS:
                                if row_fill is None:
                                    try:
                                        cell_obj = sh.Cells(r_abs, c_start + c_off)
                                        if cell_obj.Interior.ColorIndex != -4142:
                                            cell_fill = cell_obj.Interior.Color
                                            if cell_fill is not None and _bgr_luminance(cell_fill) < 0.45:
                                                continue
                                    except Exception:
                                        pass
                                counts[cat] += 1
                                cell_addr = f"{_col_letter(c_start + c_off)}{r_abs}"
                                color_buckets[cat].append(cell_addr)

                    for cat, addrs in color_buckets.items():
                        bgr_color = hex_to_bgr(CFO_FONT_COLORS[cat])
                        for i in range(0, len(addrs), 25):
                            chunk_ref = ",".join(addrs[i : i + 25])
                            try:
                                sh.Range(chunk_ref).Font.Color = bgr_color
                            except Exception:
                                for single_addr in addrs[i : i + 25]:
                                    try:
                                        sh.Range(single_addr).Font.Color = bgr_color
                                    except Exception:
                                        pass

                    if highlight_assumption_rows and FinancialModelClassifier.classify_zone(s_name) == "Inputs":
                        last_col = _col_letter(c_start + max(0, len(values_2d[0]) - 1 if values_2d else 0))
                        for hr in highlight_assumption_rows:
                            try:
                                sh.Range(f"B{hr}:{last_col}{hr}").Interior.Color = hex_to_bgr("#FEF9C3")
                            except Exception:
                                pass

            com_manager.activate_app("excel")
            return {
                "success": True,
                "workbook_name": wb.Name,
                "sheets_processed": sheets_processed,
                "color_counts": counts,
                "calc_hardcode_violations_count": len(violations),
                "calc_hardcode_violations": violations[:50],
            }
        except Exception as e:
            return {"success": False, "error": f"CFO 5색 코딩 적용 실패: {e}"}

    @mcp.tool()
    def excel_plan_adaptive_dashboard(
        row_count: int | None = None,
        data_sheet_name: str | None = None,
        kpi_sample_values: list[str] | None = None,
        chart_titles: list[str] | None = None,
        apply_column_widths_to_sheet: str | None = None,
        apply_row_height: bool = False,
    ) -> dict[str, Any]:
        """
        데이터 크기(`row_count`)별 적정 대시보드 복잡도(KPI/차트/스파크라인/조건부서식)를 계산하고,
        24pt KPI 숫자의 `###` 잘림을 막는 동적 열 너비(`max(22, ceil((chars+2)*1.3))`) 및
        차트 제목 잘림 방지 최소 열 폭(`ceil(len(title)*0.18)`)을 산출·적용합니다.
        """
        detected_rows = row_count
        excel_app = None
        wb = None

        if detected_rows is None or apply_column_widths_to_sheet:
            ensure_com()
            excel_app, wb, active_sh = com_manager.get_excel(create_if_missing=False)
            if detected_rows is None:
                if wb is not None:
                    src = wb.Sheets(data_sheet_name) if data_sheet_name else active_sh
                    detected_rows = max(0, int(src.UsedRange.Rows.Count) - 1)
                else:
                    detected_rows = 25

        complexity = DashboardBudgetCalculator.scale_complexity(int(detected_rows))
        samples = kpi_sample_values or ["₩1,958,414,250", "104.2%", "+18.5%"]
        titles = chart_titles or ["Monthly Revenue vs Target", "Department Attrition Rate"]

        kpi_widths = [
            {"sample": str(s), "visible_chars": len(str(s)), "recommended_width": DashboardBudgetCalculator.kpi_column_width(str(s))}
            for s in samples
        ]
        chart_budgets = [
            {"title": t, "title_length": len(t), "min_chart_cols": DashboardBudgetCalculator.chart_column_span(t)}
            for t in titles
        ]

        applied_cols: list[str] = []
        if apply_column_widths_to_sheet and wb is not None:
            try:
                target_sh = wb.Sheets(apply_column_widths_to_sheet)
                for idx, item in enumerate(kpi_widths, start=1):
                    col_let = _col_letter(idx)
                    target_sh.Columns(idx).ColumnWidth = item["recommended_width"]
                    applied_cols.append(f"{col_let}={item['recommended_width']}")
                if apply_row_height and complexity["sparklines"] != "skip":
                    max_r = max(3, int(target_sh.UsedRange.Row + target_sh.UsedRange.Rows.Count - 1))
                    target_sh.Range(f"3:{min(max_r, 50)}").RowHeight = 22
                if len(kpi_widths) >= 5:
                    try:
                        target_sh.PageSetup.Orientation = 2  # xlLandscape
                    except Exception:
                        pass
            except Exception as e:
                return {"success": False, "error": f"시트 열 너비 적용 실패: {e}"}

        return {
            "success": True,
            "row_count": int(detected_rows),
            "complexity_scaling": complexity,
            "kpi_column_widths": kpi_widths,
            "chart_width_budgets": chart_budgets,
            "sparkline_min_row_height_pt": 22 if complexity["sparklines"] != "skip" else None,
            "applied_columns": applied_cols,
        }
