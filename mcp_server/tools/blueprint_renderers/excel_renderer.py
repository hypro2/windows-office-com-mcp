# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_renderers/excel_renderer.py
Excel 청사진을 데스크톱 COM으로 100% 동일하게 실시간 재생(Playback)하는 렌더러.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com
from ..excel_tools import excel_freeze_display

logger = logging.getLogger("mcp_server.excel_renderer")


class ExcelBlueprintRenderer:
    """Excel 사용자 청사진 실시간 COM 렌더러"""

    @classmethod
    def render(
        cls,
        blueprint: dict[str, Any],
        payload: dict[str, Any],
        output_path: str | Path | None = None,
    ) -> dict[str, Any]:
        ensure_com()
        excel_app = com_manager.get_or_create_excel_app(visible=True)
        wb = excel_app.Workbooks.Add()

        pages = blueprint.get("pages", [])
        total_cells = 0

        with excel_freeze_display(excel_app, pause_calculation=True):
            # 필요한 시트 수 확보
            while wb.Sheets.Count < len(pages):
                wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))

            for p_idx, page in enumerate(pages, start=1):
                sheet = wb.Sheets(p_idx)
                sheet.Name = page.get("title", f"Sheet{p_idx}")[:31]

                # 열 너비 복원
                col_widths = page.get("col_widths", [])
                start_col = page.get("start_col", 1)
                for i, w in enumerate(col_widths):
                    try:
                        sheet.Columns(start_col + i).ColumnWidth = float(w)
                    except Exception:
                        pass

                # 셀 서식 및 데이터 복제
                elements = page.get("elements", [])
                for el in elements:
                    r = int(el.get("row", 1))
                    c = int(el.get("col", 1))
                    cell = sheet.Cells(r, c)

                    # 병합 처리
                    merge_area = el.get("merge_area")
                    if merge_area:
                        try:
                            sheet.Range(merge_area).Merge()
                        except Exception:
                            pass

                    # 서식 복원
                    if el.get("font_name"):
                        cell.Font.Name = el["font_name"]
                    if el.get("font_size"):
                        cell.Font.Size = float(el["font_size"])
                    if el.get("bold") is not None:
                        cell.Font.Bold = -1 if el["bold"] else 0
                    if el.get("font_color_bgr"):
                        cell.Font.Color = int(el["font_color_bgr"])
                    if el.get("bg_color_bgr") is not None:
                        cell.Interior.Color = int(el["bg_color_bgr"])

                    # 값 / 수식 주입
                    slot_id = el.get("slot_id")
                    if slot_id and slot_id in payload:
                        cell.Value = payload[slot_id]
                    elif el.get("formula"):
                        cell.Formula = el["formula"]
                    elif el.get("text"):
                        cell.Value = el["text"]

                    total_cells += 1

        out_file = None
        if output_path:
            out_file = str(Path(output_path).resolve())
            wb.SaveAs(out_file)
            logger.info(f"Excel 청사진 문서 저장 완료: {out_file}")

        try:
            excel_app.Visible = True
        except Exception:
            pass

        return {
            "success": True,
            "doc_type": "excel",
            "template_name": blueprint.get("template_name", ""),
            "sheets_count": len(pages),
            "cells_styled_count": total_cells,
            "output_path": out_file,
            "message": f"Excel 워크북({len(pages)}개 시트, {total_cells}개 서식 셀)이 성공적으로 복제 생성되었습니다.",
        }

