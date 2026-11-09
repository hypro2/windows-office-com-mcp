# -*- coding: utf-8 -*-
"""
mcp_server/tools/excel_tools.py
MS Excel 실시간 시트/선택영역 읽기(Markdown 변환), 고속 데이터 쓰기,
KISA 1310 테두리(겉바속촉·겉바속점)·데이터 정제 및 비즈니스 표 서식 XLSX 도구.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Literal

from mcp_server.compat import MCPServer
import logging
from contextlib import contextmanager

logger = logging.getLogger("mcp_server.excel_tools")


@contextmanager
def excel_freeze_display(excel_app: Any, pause_calculation: bool = True):
    """Excel 화면 갱신 억제 및 수동 계산 전환으로 대량 작업 속도를 10배 가속."""
    prev_screen_updating = True
    prev_display_alerts = True
    prev_calculation = -4105  # xlCalculationAutomatic
    try:
        if excel_app:
            prev_screen_updating = getattr(excel_app, "ScreenUpdating", True)
            prev_display_alerts = getattr(excel_app, "DisplayAlerts", True)
            prev_calculation = getattr(excel_app, "Calculation", -4105)

            excel_app.ScreenUpdating = False
            excel_app.DisplayAlerts = False
            if pause_calculation:
                excel_app.Calculation = -4135  # xlCalculationManual
    except Exception:
        pass
    try:
        yield
    finally:
        try:
            if excel_app:
                excel_app.ScreenUpdating = prev_screen_updating
                excel_app.DisplayAlerts = prev_display_alerts
                if pause_calculation:
                    excel_app.Calculation = prev_calculation
                    excel_app.Calculate()
        except Exception:
            pass

from ..com_manager import com_manager, ensure_com, hex_to_bgr
from .parsers import calculate_korean_age, strip_markdown_syntax


def _2d_to_markdown(rows: list[list[Any]]) -> str:
    if not rows:
        return ""
    col_cnt = max((len(r) for r in rows), default=0)
    if col_cnt == 0:
        return ""
    lines: list[str] = []
    for idx, r in enumerate(rows):
        vals = [str(r[c] if c < len(r) and r[c] is not None else "").replace("\n", " ").strip() for c in range(col_cnt)]
        lines.append("| " + " | ".join(vals) + " |")
        if idx == 0:
            lines.append("| " + " | ".join(["---"] * col_cnt) + " |")
    return "\n".join(lines)


def _normalize_com_value(val: Any) -> list[list[Any]]:
    if val is None:
        return []
    if isinstance(val, tuple):
        if val and isinstance(val[0], tuple):
            return [[cell for cell in row] for row in val]
        return [[cell for cell in val]]
    return [[val]]


def _apply_excel_borders(rng: Any, border_preset: str) -> None:
    """
    Excel Range에 테두리 프리셋 적용:
    - 7: xlEdgeLeft, 8: xlEdgeTop, 9: xlEdgeBottom, 10: xlEdgeRight
    - 11: xlInsideVertical, 12: xlInsideHorizontal
    - LineStyle: 1(xlContinuous), -4118(xlDot), -4115(xlDash)
    - Weight: 1(xlHairline), 2(xlThin), -4138(xlMedium), 4(xlThick)
    """
    preset = border_preset.lower().strip()
    if preset in ("겉바속촉", "outer_medium_inner_thin"):
        # KISA 1310 '겉바속촉': 바깥 굵은 실선(-4138), 안쪽 얇은 실선(2)
        for edge in (7, 8, 9, 10):
            rng.Borders(edge).LineStyle = 1
            rng.Borders(edge).Weight = -4138
        for inside in (11, 12):
            try:
                rng.Borders(inside).LineStyle = 1
                rng.Borders(inside).Weight = 2
            except Exception:
                pass
    elif preset in ("겉바속점", "outer_medium_inner_dot"):
        # KISA 1310 '겉바속점': 바깥 굵은 실선(-4138), 안쪽 점선(-4118)
        for edge in (7, 8, 9, 10):
            rng.Borders(edge).LineStyle = 1
            rng.Borders(edge).Weight = -4138
        for inside in (11, 12):
            try:
                rng.Borders(inside).LineStyle = -4118
                rng.Borders(inside).Weight = 2
            except Exception:
                pass
    elif preset == "ib_hair":
        # 비즈니스/IB 보고서 표: 바깥 얇은 실선(2), 안쪽 가장 가는 점선(xlHairline=1)
        for edge in (7, 8, 9, 10):
            rng.Borders(edge).LineStyle = 1
            rng.Borders(edge).Weight = 2
        for inside in (11, 12):
            try:
                rng.Borders(inside).LineStyle = 1
                rng.Borders(inside).Weight = 1  # xlHairline
            except Exception:
                pass
    elif preset in ("all_thin", "실선"):
        for b_idx in (7, 8, 9, 10, 11, 12):
            try:
                rng.Borders(b_idx).LineStyle = 1
                rng.Borders(b_idx).Weight = 2
            except Exception:
                pass
    elif preset in ("none", "없음"):
        rng.Borders.LineStyle = -4142  # xlNone


def register_excel_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def excel_read_data(
        range_address: str | None = None,
        sheet_name: str | None = None,
        use_selection: bool = False,
        max_rows: int = 200,
    ) -> dict[str, Any]:
        """
        현재 열려 있는 MS Excel 워크북의 셀 데이터를 읽어 2차원 배열과 Markdown 표로 반환합니다.
        - use_selection=True: 사용자가 마우스로 드래그/선택한 셀 영역(Selection)을 읽어옵니다.
        - range_address="A1:G30": 지정한 셀 범위를 읽어옵니다 (생략 시 시트 전체 UsedRange).
        """
        ensure_com()
        excel_app, wb, active_sheet = com_manager.get_excel(create_if_missing=False)
        if excel_app is None or wb is None:
            return {"success": False, "error": "현재 열려 있는 MS Excel 워크북이 없습니다."}

        try:
            sheet = wb.Sheets(sheet_name) if sheet_name else active_sheet
            if use_selection:
                rng = excel_app.Selection
                addr = rng.Address
            elif range_address:
                rng = sheet.Range(range_address)
                addr = rng.Address
            else:
                rng = sheet.UsedRange
                addr = rng.Address

            rows_2d = _normalize_com_value(rng.Value)
            truncated = len(rows_2d) > max_rows
            rows_slice = rows_2d[:max_rows]

            return {
                "success": True,
                "workbook_name": wb.Name,
                "sheet_name": sheet.Name,
                "sheets_list": [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)],
                "range_address": addr,
                "row_count": len(rows_2d),
                "col_count": max((len(r) for r in rows_2d), default=0),
                "truncated": truncated,
                "markdown_table": _2d_to_markdown(rows_slice),
                "data": rows_slice,
            }
        except Exception as e:
            return {"success": False, "error": f"Excel 데이터 읽기 실패: {e}"}

    @mcp.tool()
    def excel_write_data(
        data: list[list[Any]],
        start_cell: str = "A1",
        sheet_name: str | None = None,
        auto_fit: bool = True,
        table_style: Literal["navy", "modern_teal", "ib_steel_blue", "겉바속촉", "none"] = "navy",
    ) -> dict[str, Any]:
        """
        MS Excel 시트의 지정한 시작 셀(start_cell)부터 2차원 배열 데이터를 한 번에 작성하고,
        원하는 표 서식(비즈니스 네이비 헤더, 모던 틸 헤더, IB 스틸블루 헤더, KISA 1310 겉바속촉 테두리 등)을 자동 적용합니다.
        """
        ensure_com()
        if not data or not data[0]:
            return {"success": False, "error": "작성할 데이터(data)가 비어 있습니다."}

        excel_app, wb, active_sheet = com_manager.get_excel(create_if_missing=True, visible=True)
        if excel_app is None or wb is None:
            return {"success": False, "error": "Excel 연결 실패"}

        try:
            if sheet_name:
                try:
                    sheet = wb.Sheets(sheet_name)
                except Exception:
                    sheet = wb.Sheets.Add()
                    sheet.Name = sheet_name[:31]
            else:
                sheet = active_sheet

            n_rows = len(data)
            n_cols = max(len(r) for r in data)
            cleaned_data = [
                [strip_markdown_syntax(str(c)) if isinstance(c, str) else c for c in r]
                for r in data
            ]
            padded = [list(r) + [""] * (n_cols - len(r)) for r in cleaned_data]

            start_rng = sheet.Range(start_cell)
            end_rng = sheet.Cells(start_rng.Row + n_rows - 1, start_rng.Column + n_cols - 1)
            target_rng = sheet.Range(start_rng, end_rng)
            target_rng.Value = padded

            # 서식 적용
            if table_style != "none":
                hdr_start = sheet.Cells(start_rng.Row, start_rng.Column)
                hdr_end = sheet.Cells(start_rng.Row, start_rng.Column + n_cols - 1)
                hdr_rng = sheet.Range(hdr_start, hdr_end)

                if table_style == "navy":
                    hdr_rng.Interior.Color = hex_to_bgr("#1F3864")
                    hdr_rng.Font.Color = hex_to_bgr("#FFFFFF")
                    hdr_rng.Font.Bold = True
                    target_rng.Font.Name = "맑은 고딕"
                    _apply_excel_borders(target_rng, "겉바속촉")
                elif table_style == "modern_teal":
                    hdr_rng.Interior.Color = hex_to_bgr("#0F766E")
                    hdr_rng.Font.Color = hex_to_bgr("#FFFFFF")
                    hdr_rng.Font.Bold = True
                    target_rng.Font.Name = "맑은 고딕"
                    _apply_excel_borders(target_rng, "겉바속촉")
                elif table_style == "ib_steel_blue":
                    hdr_rng.Interior.Color = hex_to_bgr("#2E608A")
                    hdr_rng.Font.Color = hex_to_bgr("#FFFFFF")
                    hdr_rng.Font.Bold = True
                    target_rng.Font.Name = "맑은 고딕"
                    _apply_excel_borders(target_rng, "ib_hair")
                elif table_style == "겉바속촉":
                    hdr_rng.Font.Bold = True
                    _apply_excel_borders(target_rng, "겉바속촉")

            if auto_fit:
                target_rng.Columns.AutoFit()

            com_manager.activate_app("excel")

            return {
                "success": True,
                "workbook_name": wb.Name,
                "sheet_name": sheet.Name,
                "written_range": target_rng.Address,
                "rows": n_rows,
                "cols": n_cols,
                "table_style": table_style,
            }
        except Exception as e:
            return {"success": False, "error": f"Excel 데이터 쓰기 실패: {e}"}

    @mcp.tool()
    def excel_format_range(
        range_address: str | None = None,
        use_selection: bool = True,
        sheet_name: str | None = None,
        style_preset: Literal["navy_header", "teal_header", "ib_steel_blue_header", "ib_section_bar"] | None = None,
        border_preset: Literal["겉바속촉", "겉바속점", "ib_hair", "all_thin", "none"] | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        bg_color_hex: str | None = None,
        font_color_hex: str | None = None,
        align: Literal["left", "center", "right"] | None = None,
        number_format: str | None = None,
        auto_fit: bool = False,
        activate_window: bool = True,
    ) -> dict[str, Any]:
        """
        MS Excel에서 현재 마우스로 선택한 영역(use_selection=True) 또는 지정한 셀 범위(range_address)에
        배경색, 글꼴, 정렬, 숫자 서식(#,##0 등), 그리고 테두리 프리셋(겉바속촉, 겉바속점, ib_hair, all_thin)을 실시간 적용합니다.
        """
        ensure_com()
        excel_app, wb, active_sheet = com_manager.get_excel(create_if_missing=False)
        if excel_app is None or wb is None:
            return {"success": False, "error": "열려 있는 Excel 워크북이 없습니다."}

        try:
            sheet = wb.Sheets(sheet_name) if sheet_name else active_sheet
            rng = sheet.Range(range_address) if range_address else excel_app.Selection

            if style_preset == "navy_header":
                rng.Interior.Color = hex_to_bgr("#1F3864")
                rng.Font.Color = hex_to_bgr("#FFFFFF")
                rng.Font.Bold = True
                rng.HorizontalAlignment = -4108  # xlCenter
            elif style_preset == "teal_header":
                rng.Interior.Color = hex_to_bgr("#0F766E")
                rng.Font.Color = hex_to_bgr("#FFFFFF")
                rng.Font.Bold = True
                rng.HorizontalAlignment = -4108
            elif style_preset == "ib_steel_blue_header":
                rng.Interior.Color = hex_to_bgr("#2E608A")
                rng.Font.Color = hex_to_bgr("#FFFFFF")
                rng.Font.Bold = True
                rng.HorizontalAlignment = -4108
            elif style_preset == "ib_section_bar":
                rng.Interior.Color = hex_to_bgr("#000000")
                rng.Font.Color = hex_to_bgr("#FFFFFF")
                rng.Font.Bold = True
                rng.Font.Size = 12

            if font_name:
                rng.Font.Name = font_name
            if font_size is not None:
                rng.Font.Size = float(font_size)
            if bold is not None:
                rng.Font.Bold = bool(bold)
            if bg_color_hex:
                rng.Interior.Color = hex_to_bgr(bg_color_hex)
            if font_color_hex:
                rng.Font.Color = hex_to_bgr(font_color_hex)
            if align:
                align_map = {"left": -4131, "center": -4108, "right": -4152}
                if align.lower() in align_map:
                    rng.HorizontalAlignment = align_map[align.lower()]
            if number_format:
                rng.NumberFormat = number_format
            if border_preset:
                _apply_excel_borders(rng, border_preset)
            if auto_fit:
                rng.Columns.AutoFit()

            if activate_window:
                com_manager.activate_app("excel")

            return {"success": True, "formatted_range": rng.Address}
        except Exception as e:
            return {"success": False, "error": f"Excel 서식 적용 실패: {e}"}

    @mcp.tool()
    def excel_clean_selection(
        action: Literal["trim_whitespace", "mask_rrn", "calc_korean_age"] = "trim_whitespace",
    ) -> dict[str, Any]:
        """
        KISA 범피스 1310의 엑셀 데이터 정제 유틸리티를 현재 선택된 영역(Selection)에 실시간 수행합니다:
        - "trim_whitespace" (공백정리): 앞뒤 공백 및 중복 줄바꿈/탭을 깔끔하게 정리
        - "mask_rrn" (주민암호화): 주민등록번호 뒷자리(7자리 중 뒤 6자리)를 '******'로 마스킹
        - "calc_korean_age" (만나이계산): 선택된 주민등록번호 셀 옆 열에 만 나이를 자동 계산해 기입
        """
        ensure_com()
        excel_app, wb, _ = com_manager.get_excel(create_if_missing=False)
        if excel_app is None or wb is None:
            return {"success": False, "error": "열려 있는 Excel 워크북이 없습니다."}

        try:
            sel = excel_app.Selection
            rows_cnt = sel.Rows.Count
            cols_cnt = sel.Columns.Count
            updated = 0
            for r in range(1, rows_cnt + 1):
                for c in range(1, cols_cnt + 1):
                    cell_obj = sel.Cells(r, c)
                    val = cell_obj.Value
                    if val is None:
                        continue
                    s_val = str(val)
                    if action == "trim_whitespace":
                        cleaned = re.sub(r"\n+", "\n", s_val.strip())
                        cleaned = re.sub(r"[ \t]+", " ", cleaned)
                        cell_obj.Value = cleaned
                        updated += 1
                    elif action == "mask_rrn":
                        masked = re.sub(r"(\d{6})[- ]?([1-4])\d{6}", r"\1-\2******", s_val)
                        cell_obj.Value = masked
                        updated += 1
                    elif action == "calc_korean_age":
                        age = calculate_korean_age(s_val)
                        if age >= 0:
                            cell_obj.Offset(1, 2).Value = age  # COM Offset(1,2) is same row, next col
                            updated += 1
            return {"success": True, "action": action, "cells_processed": updated}
        except Exception as e:
            return {"success": False, "error": f"Excel 데이터 정제 오류: {e}"}

    @mcp.tool()
    def excel_save_workbook(file_path: str | None = None) -> dict[str, Any]:
        """현재 활성 Excel 워크북을 저장합니다. file_path 지정 시 다른 이름으로 저장합니다."""
        ensure_com()
        excel_app, wb, _ = com_manager.get_excel(create_if_missing=False)
        if excel_app is None or wb is None:
            return {"success": False, "error": "열려 있는 Excel 워크북이 없습니다."}
        try:
            if file_path:
                abs_path = str(Path(file_path).resolve())
                Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
                prev_alerts = excel_app.DisplayAlerts
                try:
                    excel_app.DisplayAlerts = False
                    wb.SaveAs(abs_path)
                finally:
                    excel_app.DisplayAlerts = prev_alerts
                return {"success": True, "path": abs_path}
            wb.Save()
            return {"success": True, "path": wb.FullName}
        except Exception as e:
            return {"success": False, "error": f"Excel 저장 실패: {e}"}


    @mcp.tool()
    def excel_run_macro(
        macro_name: str,
        args: list[Any] | None = None,
        workbook_name: str | None = None,
    ) -> dict[str, Any]:
        """
        현재 열려 있는 Excel 워크북의 VBA 매크로(서브루틴 또는 함수)를 실행합니다.
        (ExcelMcp 및 vba-mcp 벤치마킹)

        - macro_name: 실행할 매크로 이름 (예: "Module1.ProcessData", "CalculateTax", "AutoSummary")
        - args: 매크로에 전달할 인자 리스트 (생략 가능)
        - workbook_name: 특정 워크북 지정 시 해당 워크북 내 매크로 실행 (생략 시 활성 워크북)
        """
        ensure_com()
        excel_app, wb, _ = com_manager.get_excel(create_if_missing=False)
        if excel_app is None:
            return {"success": False, "error": "열려 있는 Excel 인스턴스가 없습니다."}

        try:
            target_macro = macro_name
            if workbook_name:
                wb_clean = Path(workbook_name).name
                if "!" not in target_macro:
                    target_macro = f"'{wb_clean}'!{target_macro}"

            call_args = args or []
            res = excel_app.Run(target_macro, *call_args)

            norm_res = res
            if isinstance(res, tuple):
                norm_res = list(res)

            return {
                "success": True,
                "macro_name": target_macro,
                "args_count": len(call_args),
                "result": norm_res,
                "message": f"매크로 '{target_macro}'가 성공적으로 실행되었습니다.",
            }
        except Exception as e:
            return {"success": False, "error": f"VBA 매크로 실행 실패: {e}"}

    @mcp.tool()
    def excel_manage_sheets(
        action: Literal["list", "add", "delete", "rename", "copy", "activate"],
        sheet_name: str | None = None,
        new_name: str | None = None,
        before_sheet: str | None = None,
        after_sheet: str | None = None,
    ) -> dict[str, Any]:
        """
        Excel 워크북의 시트 라이프사이클을 관리합니다 (목록조회, 추가, 삭제, 이름변경, 복제, 활성화).

        - action: "list" | "add" | "delete" | "rename" | "copy" | "activate"
        - sheet_name: 대상 시트 이름 (add 시에는 새로 생성할 시트 이름, list 시에는 생략 가능)
        - new_name: rename 시 변경할 새 시트 이름 (최대 31자)
        - before_sheet: add/copy 시 기준이 될 앞쪽 시트 이름
        - after_sheet: add/copy 시 기준이 될 뒤쪽 시트 이름
        """
        ensure_com()
        excel_app, wb, active_sheet = com_manager.get_excel(create_if_missing=False)
        if excel_app is None or wb is None:
            return {"success": False, "error": "열려 있는 Excel 워크북이 없습니다."}

        try:
            # 1. 시트 목록 조회 (list)
            if action == "list":
                sheets_info = []
                for i in range(1, wb.Sheets.Count + 1):
                    sh = wb.Sheets(i)
                    sheets_info.append({
                        "index": i,
                        "name": sh.Name,
                        "visible": sh.Visible == -1,
                        "is_active": (active_sheet is not None and sh.Name == active_sheet.Name),
                    })
                return {
                    "success": True,
                    "action": "list",
                    "workbook_name": wb.Name,
                    "total_sheets": wb.Sheets.Count,
                    "sheets": sheets_info,
                }

            # 2. 시트 추가 (add)
            elif action == "add":
                kwargs = {}
                if before_sheet:
                    kwargs["Before"] = wb.Sheets(before_sheet)
                elif after_sheet:
                    kwargs["After"] = wb.Sheets(after_sheet)
                else:
                    kwargs["After"] = wb.Sheets(wb.Sheets.Count)

                new_sh = wb.Sheets.Add(**kwargs)
                if sheet_name:
                    clean_n = re.sub(r'[:\\/?*\[\]]', '_', str(sheet_name))[:31]
                    new_sh.Name = clean_n
                new_sh.Activate()
                return {
                    "success": True,
                    "action": "add",
                    "sheet_name": new_sh.Name,
                    "total_sheets": wb.Sheets.Count,
                    "message": f"시트 '{new_sh.Name}'가 성공적으로 추가되었습니다.",
                }

            # 3. 시트 삭제 (delete)
            elif action == "delete":
                if not sheet_name:
                    return {"success": False, "error": "삭제할 시트 이름(sheet_name)을 지정해야 합니다."}
                if wb.Sheets.Count <= 1:
                    return {"success": False, "error": "워크북에는 최소 1개 이상의 시트가 남아 있어야 하므로 삭제할 수 없습니다."}

                target_sh = wb.Sheets(sheet_name)
                del_name = target_sh.Name

                prev_alerts = excel_app.DisplayAlerts
                try:
                    excel_app.DisplayAlerts = False
                    target_sh.Delete()
                finally:
                    excel_app.DisplayAlerts = prev_alerts

                return {
                    "success": True,
                    "action": "delete",
                    "deleted_sheet": del_name,
                    "total_sheets": wb.Sheets.Count,
                    "message": f"시트 '{del_name}'가 안전하게 삭제되었습니다.",
                }

            # 4. 시트 이름 변경 (rename)
            elif action == "rename":
                if not sheet_name or not new_name:
                    return {"success": False, "error": "기존 시트명(sheet_name)과 새 시트명(new_name)을 모두 지정해야 합니다."}

                target_sh = wb.Sheets(sheet_name)
                clean_n = re.sub(r'[:\\/?*\[\]]', '_', str(new_name))[:31]
                target_sh.Name = clean_n
                return {
                    "success": True,
                    "action": "rename",
                    "old_name": sheet_name,
                    "new_name": clean_n,
                    "message": f"시트명이 '{sheet_name}'에서 '{clean_n}'으로 변경되었습니다.",
                }

            # 5. 시트 복제 (copy)
            elif action == "copy":
                if not sheet_name:
                    return {"success": False, "error": "복제할 원본 시트 이름(sheet_name)을 지정해야 합니다."}

                target_sh = wb.Sheets(sheet_name)
                kwargs = {}
                if before_sheet:
                    kwargs["Before"] = wb.Sheets(before_sheet)
                elif after_sheet:
                    kwargs["After"] = wb.Sheets(after_sheet)
                else:
                    kwargs["After"] = target_sh

                target_sh.Copy(**kwargs)
                copied_sh = excel_app.ActiveSheet
                if new_name:
                    clean_n = re.sub(r'[:\\/?*\[\]]', '_', str(new_name))[:31]
                    copied_sh.Name = clean_n

                return {
                    "success": True,
                    "action": "copy",
                    "source_sheet": sheet_name,
                    "copied_sheet": copied_sh.Name,
                    "total_sheets": wb.Sheets.Count,
                    "message": f"시트 '{sheet_name}'가 '{copied_sh.Name}'으로 복제되었습니다.",
                }

            # 6. 시트 활성화 (activate)
            elif action == "activate":
                if not sheet_name:
                    return {"success": False, "error": "활성화할 시트 이름(sheet_name)을 지정해야 합니다."}
                sh = wb.Sheets(sheet_name)
                sh.Activate()
                com_manager.activate_app("excel")
                return {
                    "success": True,
                    "action": "activate",
                    "sheet_name": sh.Name,
                    "message": f"시트 '{sh.Name}'가 활성화되었습니다.",
                }

            else:
                return {"success": False, "error": f"지원하지 않는 action: '{action}'"}

        except Exception as e:
            return {"success": False, "error": f"시트 관리 작업({action}) 실패: {e}"}

    @mcp.tool()
    def excel_evaluate_formula(
        formula: str,
        sheet_name: str | None = None,
        calculate_workbook: bool = False,
    ) -> dict[str, Any]:
        """
        Excel 수식을 실시간으로 평가(Evaluate)하여 결과값을 반환하거나 전체 워크북을 재계산합니다.
        셀에 수식을 직접 적지 않고도 복잡한 계산 결과(SUM, VLOOKUP, INDEX/MATCH, 통계치 등)를 즉시 계산할 수 있습니다.

        - formula: 엑셀 수식 (예: "SUM(A1:A50)", "AVERAGE(B2:B100)", "COUNTIF(C:C, \">100\")", "TODAY()+30")
        - sheet_name: 수식 실행 컨텍스트 대상 시트 (생략 시 현재 활성 시트)
        - calculate_workbook: True 시 수식 평가 전 전체 워크북 강제 재계산(Calculate) 수행
        """
        ensure_com()
        excel_app, wb, active_sheet = com_manager.get_excel(create_if_missing=False)
        if excel_app is None or wb is None:
            return {"success": False, "error": "열려 있는 Excel 워크북이 없습니다."}

        try:
            if calculate_workbook:
                excel_app.Calculate()

            sheet = wb.Sheets(sheet_name) if sheet_name else active_sheet

            clean_formula = formula.strip()
            eval_str = clean_formula[1:] if clean_formula.startswith("=") else clean_formula

            res = sheet.Evaluate(eval_str) if sheet else excel_app.Evaluate(eval_str)

            if isinstance(res, tuple):
                res_norm = _normalize_com_value(res)
            else:
                res_norm = res

            return {
                "success": True,
                "formula": formula,
                "sheet_name": sheet.Name if sheet else wb.Name,
                "result": res_norm,
            }
        except Exception as e:
            return {"success": False, "error": f"수식 평가 실패: {e}"}
