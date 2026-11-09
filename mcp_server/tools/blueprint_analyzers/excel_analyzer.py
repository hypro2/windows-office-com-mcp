# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_analyzers/excel_analyzer.py
MS Excel (XLSX) 시트 구조, 셀 서식, 수식, 열 너비 역공학 분석기.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com
from .base_analyzer import BaseDocumentAnalyzer

logger = logging.getLogger("mcp_server.excel_analyzer")


class ExcelDocumentAnalyzer(BaseDocumentAnalyzer):
    doc_type: str = "excel"

    def analyze_file(self, file_path: Path) -> dict[str, Any]:
        """XLSX 파일을 열어 시트 구조, 스타일 및 데이터 슬롯 추출 (DRM 친화적 COM 개방)"""
        ensure_com()
        excel_app = com_manager.get_or_create_excel_app(visible=True)
        try:
            excel_app.DisplayAlerts = False
        except Exception:
            pass

        abs_path = str(file_path.resolve())

        wb = None
        try:
            # 사내 DRM 복호화 후킹 통과 및 외부 링크 업데이트 경고 억제
            wb = excel_app.Workbooks.Open(
                abs_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )
            return self._extract_from_workbook(wb, file_path.stem)
        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def analyze_active_window(self) -> dict[str, Any]:
        """현재 모니터에 열려 있는 활성 Excel 워크북 분석"""
        ensure_com()
        excel_app, wb = com_manager.get_excel(create_if_missing=False)
        if not excel_app or not wb:
            raise RuntimeError("현재 활성화된 Excel 워크북 창을 찾을 수 없습니다.")

        name = getattr(wb, "Name", "ActiveWorkbook")
        return self._extract_from_workbook(wb, Path(name).stem)

    def _extract_from_workbook(self, wb: Any, source_name: str) -> dict[str, Any]:
        pages: list[dict[str, Any]] = []
        fonts: set[str] = set()
        colors: list[int] = []

        for sh_idx in range(1, wb.Sheets.Count + 1):
            sheet = wb.Sheets(sh_idx)
            sh_name = sheet.Name
            used = sheet.UsedRange

            start_row = used.Row
            start_col = used.Column
            num_rows = min(100, used.Rows.Count)  # 최대 100행 샘플링
            num_cols = min(30, used.Columns.Count) # 최대 30열 샘플링

            # 열 너비 추출
            col_widths = []
            for c in range(start_col, start_col + num_cols):
                try:
                    w = round(float(sheet.Columns(c).ColumnWidth), 1)
                    col_widths.append(w)
                except Exception:
                    col_widths.append(8.5)

            # 셀 스타일 및 슬롯 추출
            elements = []
            merged_ranges = set()

            for r_idx in range(start_row, start_row + num_rows):
                for c_idx in range(start_col, start_col + num_cols):
                    cell = sheet.Cells(r_idx, c_idx)
                    addr = cell.Address.replace("$", "")

                    # 병합 셀 검사
                    is_merged = bool(getattr(cell, "MergeCells", False))
                    merge_addr = ""
                    if is_merged:
                        merge_addr = cell.MergeArea.Address.replace("$", "")
                        if merge_addr in merged_ranges and addr != merge_addr.split(":")[0]:
                            continue  # 병합의 첫 셀만 분석
                        merged_ranges.add(merge_addr)

                    val = cell.Value
                    formula = str(cell.Formula) if cell.HasFormula else ""
                    val_str = str(val).strip() if val is not None else ""

                    f_name = str(cell.Font.Name)
                    fonts.add(f_name)
                    f_size = round(float(cell.Font.Size), 1)
                    f_bold = bool(cell.Font.Bold)
                    f_color = int(cell.Font.Color)

                    bg_color = None
                    try:
                        if cell.Interior.ColorIndex != -4142:  # xlNone = -4142
                            bg_color = int(cell.Interior.Color)
                            colors.append(bg_color)
                    except Exception:
                        pass

                    # 고정 헤더인지(배경색이 있거나 굵은 폰트이면서 텍스트가 있는 경우)
                    is_header = bool((bg_color is not None or f_bold) and val_str and not formula)

                    if val_str or bg_color is not None or is_merged or formula:
                        elements.append({
                            "element_id": f"cell_{addr}",
                            "cell_address": addr,
                            "merge_area": merge_addr if is_merged else None,
                            "row": r_idx,
                            "col": c_idx,
                            "text": val_str,
                            "formula": formula,
                            "font_name": f_name,
                            "font_size": f_size,
                            "bold": f_bold,
                            "font_color_bgr": f_color,
                            "bg_color_bgr": bg_color,
                            "is_static": is_header or bool(formula),
                            "slot_id": f"slot_{sh_name}_{addr}" if not is_header and not formula and val_str else None,
                            "label": f"[{sh_name}] {addr} 데이터" if not is_header and not formula and val_str else "",
                        })

            pages.append({
                "page_index": sh_idx,
                "title": sh_name,
                "start_row": start_row,
                "start_col": start_col,
                "num_rows": num_rows,
                "num_cols": num_cols,
                "col_widths": col_widths,
                "elements": elements,
            })

        return {
            "doc_type": "excel",
            "source_name": source_name,
            "canvas": {"sheets_count": len(pages)},
            "design_tokens": {
                "fonts": list(fonts)[:5],
                "sample_colors_bgr": colors[:8],
            },
            "pages": pages,
        }

