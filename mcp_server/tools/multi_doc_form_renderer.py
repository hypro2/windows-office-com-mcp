# -*- coding: utf-8 -*-
"""
mcp_server/tools/multi_doc_form_renderer.py
283종 표준 참조 서식(form_catalog/form_knowledge)을 Word뿐만 아니라
한글(HWP), 엑셀(Excel), 파워포인트(PPT) 4대 오피스 데스크톱 창에 실시간 주입하는 멀티 타깃 렌더러.
"""

from __future__ import annotations

import logging
from typing import Any

from ..com_manager import com_manager, ensure_com, rgb_to_bgr

logger = logging.getLogger("mcp_server.multi_doc_forms")


class MultiDocFormRenderer:
    """4대 오피스 공용 서식 렌더러 브릿지"""

    @classmethod
    def detect_active_target(cls, doc_type: str = "auto") -> str:
        """
        현재 열려 있는 오피스 창을 감지하여 타깃 결정:
        'auto'인 경우 Word -> HWP -> Excel -> PPT 순으로 활성 창 탐색
        """
        dt = doc_type.lower().strip()
        if dt not in ("auto", "", "all"):
            if dt in ("word", "docx"):
                return "word"
            if dt in ("hwp", "hwpx", "hancom"):
                return "hwp"
            if dt in ("excel", "xlsx", "sheet"):
                return "excel"
            if dt in ("ppt", "pptx", "powerpoint"):
                return "ppt"

        # 1. Word 확인
        try:
            w_app, w_doc = com_manager.get_word(create_if_missing=False)
            if w_app and w_doc:
                return "word"
        except Exception:
            pass

        # 2. HWP 확인
        try:
            hwp = com_manager.get_hwp(create_if_missing=False)
            if hwp:
                return "hwp"
        except Exception:
            pass

        # 3. Excel 확인
        try:
            xl_app, xl_wb, _ = com_manager.get_excel(create_if_missing=False)
            if xl_app and xl_wb:
                return "excel"
        except Exception:
            pass

        # 4. PPT 확인
        try:
            ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
            if ppt_app and prs:
                return "ppt"
        except Exception:
            pass

        # 아무것도 안 열려 있으면 Word 기본 반환
        return "word"

    # =========================================================================
    # 1. HWP 렌더링
    # =========================================================================
    @classmethod
    def render_to_hwp(
        cls,
        form_name: str,
        meta: dict[str, Any],
        args: list[Any] | None = None,
        headers: list[str] | None = None,
        rows: list[list[Any]] | None = None,
        replacements: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        """한컴오피스 HWP 창에 표준 서식(표, 박스, 소제목) 삽입"""
        hwp = com_manager.get_hwp(create_if_missing=False)
        if not hwp:
            return {"success": False, "error": "현재 실행 중인 한글(HWP) 창을 찾을 수 없습니다."}

        cols = meta.get("cols", 3)
        default_rows = meta.get("rows", 3)

        final_headers = headers or meta.get("literals", [])[:cols]
        while len(final_headers) < cols:
            final_headers.append(f"항목 {len(final_headers)+1}")

        final_rows = rows or [["" for _ in range(cols)] for _ in range(max(1, default_rows - 1))]
        total_rows = 1 + len(final_rows) if final_headers else len(final_rows)

        try:
            # HWP 표 생성 액션
            hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
            hwp.HParameterSet.HTableCreation.Rows = total_rows
            hwp.HParameterSet.HTableCreation.Cols = cols
            hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

            # 헤더 입력
            if final_headers:
                for c_idx, h_text in enumerate(final_headers):
                    hwp.HParameterSet.HInsertText.Text = str(h_text)
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    if c_idx < cols - 1:
                        hwp.HAction.Run("TableRightCell")

                # 데이터 행으로 이동
                hwp.HAction.Run("TableRightCell")

            # 데이터 행 입력
            for r_idx, r_data in enumerate(final_rows):
                for c_idx in range(cols):
                    val = str(r_data[c_idx]) if c_idx < len(r_data) else ""
                    hwp.HParameterSet.HInsertText.Text = val
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    if not (r_idx == len(final_rows) - 1 and c_idx == cols - 1):
                        hwp.HAction.Run("TableRightCell")

            # 표 밖으로 커서 이동
            hwp.HAction.Run("CloseEx")
            hwp.HAction.Run("MoveDown")

            com_manager.activate_app("hwp")
            return {
                "success": True,
                "target_app": "hwp",
                "form_name": form_name,
                "rows": total_rows,
                "cols": cols,
                "message": f"한글(HWP) 문서에 표준 양식 '{form_name}'({total_rows}x{cols})이 삽입되었습니다."
            }
        except Exception as e:
            logger.exception("HWP 양식 삽입 실패")
            return {"success": False, "target_app": "hwp", "error": f"HWP 표 삽입 실패: {e}"}

    # =========================================================================
    # 2. Excel 렌더링
    # =========================================================================
    @classmethod
    def render_to_excel(
        cls,
        form_name: str,
        meta: dict[str, Any],
        args: list[Any] | None = None,
        headers: list[str] | None = None,
        rows: list[list[Any]] | None = None,
        replacements: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        """Excel 활성 시트에 표준 서식 삽입"""
        ensure_com()
        excel_app, wb, _ = com_manager.get_excel(create_if_missing=True)
        if not excel_app or not wb:
            return {"success": False, "error": "Excel 창을 열 수 없습니다."}

        ws = wb.ActiveSheet
        cols = meta.get("cols", 4)
        default_rows = meta.get("rows", 3)

        final_headers = headers or meta.get("literals", [])[:cols]
        while len(final_headers) < cols:
            final_headers.append(f"항목 {len(final_headers)+1}")

        final_rows = rows or [["" for _ in range(cols)] for _ in range(max(1, default_rows - 1))]

        try:
            sel = excel_app.Selection
            start_r = sel.Row if sel else 2
            start_c = sel.Column if sel else 2

            # 헤더 작성
            for c_idx, h in enumerate(final_headers):
                cell = ws.Cells(start_r, start_c + c_idx)
                cell.Value = str(h)
                cell.Font.Bold = True
                cell.Font.Color = 0xFFFFFF
                cell.Interior.Color = 0x002060  # Navy
                cell.HorizontalAlignment = -4108  # xlCenter

            # 데이터 행 작성
            for r_idx, r_data in enumerate(final_rows):
                cur_r = start_r + 1 + r_idx
                for c_idx in range(cols):
                    val = r_data[c_idx] if c_idx < len(r_data) else ""
                    cell = ws.Cells(cur_r, start_c + c_idx)
                    cell.Value = val
                    if r_idx % 2 == 1:
                        cell.Interior.Color = 0xF9F9F9

            # 테두리 설정
            total_rows = 1 + len(final_rows)
            tbl_rng = ws.Range(ws.Cells(start_r, start_c), ws.Cells(start_r + total_rows - 1, start_c + cols - 1))
            tbl_rng.Borders.Color = 0xCCCCCC

            # 열 너비 자동 맞춤
            for c_idx in range(cols):
                ws.Columns(start_c + c_idx).AutoFit()

            com_manager.activate_app("excel")
            return {
                "success": True,
                "target_app": "excel",
                "form_name": form_name,
                "range": f"R{start_r}C{start_c}:R{start_r + total_rows - 1}C{start_c + cols - 1}",
                "message": f"Excel 시트에 표준 양식 '{form_name}'이 삽입되었습니다."
            }
        except Exception as e:
            logger.exception("Excel 양식 삽입 실패")
            return {"success": False, "target_app": "excel", "error": f"Excel 표 삽입 실패: {e}"}

    # =========================================================================
    # 3. PowerPoint 렌더링
    # =========================================================================
    @classmethod
    def render_to_ppt(
        cls,
        form_name: str,
        meta: dict[str, Any],
        args: list[Any] | None = None,
        headers: list[str] | None = None,
        rows: list[list[Any]] | None = None,
        replacements: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        """PowerPoint 활성 슬라이드에 표준 서식 표 삽입"""
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=True)
        if not ppt_app or not prs:
            return {"success": False, "error": "PowerPoint 창을 열 수 없습니다."}

        cols = meta.get("cols", 3)
        default_rows = meta.get("rows", 3)

        final_headers = headers or meta.get("literals", [])[:cols]
        while len(final_headers) < cols:
            final_headers.append(f"항목 {len(final_headers)+1}")

        final_rows = rows or [["" for _ in range(cols)] for _ in range(max(1, default_rows - 1))]
        total_rows = 1 + len(final_rows)

        try:
            slide = ppt_app.ActiveWindow.View.Slide if prs.Slides.Count > 0 else prs.Slides.Add(1, 12)

            table_shape = slide.Shapes.AddTable(
                NumRows=total_rows,
                NumColumns=cols,
                Left=50,
                Top=120,
                Width=860,
                Height=max(100, total_rows * 35)
            )
            tbl = table_shape.Table

            # 헤더 입력
            for c_idx, h in enumerate(final_headers, start=1):
                cell = tbl.Cell(1, c_idx)
                cell.Shape.TextFrame.TextRange.Text = str(h)
                cell.Shape.TextFrame.TextRange.Font.Bold = True
                cell.Shape.TextFrame.TextRange.Font.Size = 12
                cell.Shape.Fill.Solid()
                cell.Shape.Fill.ForeColor.RGB = 0x002060
                cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

            # 데이터 행 입력
            for r_idx, r_data in enumerate(final_rows, start=2):
                for c_idx in range(1, cols + 1):
                    val = str(r_data[c_idx - 1]) if c_idx - 1 < len(r_data) else ""
                    cell = tbl.Cell(r_idx, c_idx)
                    cell.Shape.TextFrame.TextRange.Text = val
                    cell.Shape.TextFrame.TextRange.Font.Size = 11
                    cell.Shape.Fill.Solid()
                    cell.Shape.Fill.ForeColor.RGB = 0xFAFAFA if r_idx % 2 == 1 else 0xFFFFFF

            com_manager.activate_app("powerpoint")
            return {
                "success": True,
                "target_app": "ppt",
                "form_name": form_name,
                "rows": total_rows,
                "cols": cols,
                "message": f"PowerPoint 슬라이드에 표준 서식 표 '{form_name}'이 삽입되었습니다."
            }
        except Exception as e:
            logger.exception("PPT 양식 삽입 실패")
            return {"success": False, "target_app": "ppt", "error": f"PPT 표 삽입 실패: {e}"}

    # =========================================================================
    # 4. Word 렌더링 (표준 지식베이스 메타데이터 기반 폴백/직접 삽입)
    # =========================================================================
    @classmethod
    def render_to_word(
        cls,
        form_name: str,
        meta: dict[str, Any],
        args: list[Any] | None = None,
        headers: list[str] | None = None,
        rows: list[list[Any]] | None = None,
        replacements: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        """Word 활성 문서 커서 위치에 표준 서식 표 삽입"""
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if not word_app or not doc:
            return {"success": False, "error": "열려 있는 Word 문서를 찾을 수 없습니다."}

        cols = max(1, int(meta.get("cols", 3) or 3))
        default_rows = max(1, int(meta.get("rows", 3) or 3))

        raw_lits = [str(x) for x in meta.get("literals", [])]
        if args and not headers and len(args) >= cols:
            raw_lits = [str(x) for x in args]
        final_headers = list(headers) if headers else raw_lits[:cols]
        while len(final_headers) < cols:
            final_headers.append(f"항목 {len(final_headers) + 1}")

        final_rows = rows or [["" for _ in range(cols)] for _ in range(max(1, default_rows - 1))]
        total_rows = 1 + len(final_rows)

        def _apply_repl(text: Any) -> str:
            val = str(text)
            if replacements:
                for k, v in replacements.items():
                    if k in val:
                        val = val.replace(k, str(v))
            return val

        try:
            sel = word_app.Selection
            tbl = doc.Tables.Add(sel.Range, total_rows, cols)
            tbl.Borders.Enable = True
            tbl.Borders.OutsideColor = rgb_to_bgr(31, 41, 55)
            tbl.Borders.InsideColor = rgb_to_bgr(209, 213, 219)

            header_bgr = rgb_to_bgr(0, 32, 96)
            white_bgr = rgb_to_bgr(255, 255, 255)
            zebra_bgr = rgb_to_bgr(248, 250, 252)

            for c_idx, h in enumerate(final_headers, start=1):
                cell = tbl.Cell(1, c_idx)
                cell.Shading.BackgroundPatternColor = header_bgr
                rng = cell.Range
                rng.Text = _apply_repl(h)
                rng.Font.Bold = True
                rng.Font.Size = 10
                rng.Font.Color = white_bgr

            for r_idx, r_data in enumerate(final_rows, start=2):
                for c_idx in range(1, cols + 1):
                    val = r_data[c_idx - 1] if c_idx - 1 < len(r_data) else ""
                    cell = tbl.Cell(r_idx, c_idx)
                    if r_idx % 2 == 1:
                        cell.Shading.BackgroundPatternColor = zebra_bgr
                    rng = cell.Range
                    rng.Text = _apply_repl(val)
                    rng.Font.Size = 10

            end_pos = max(0, min(tbl.Range.End, doc.Content.End - 1))
            doc.Range(end_pos, end_pos).Select()
            com_manager.activate_app("word")
            return {
                "success": True,
                "target_app": "word",
                "document_name": doc.Name,
                "form_name": form_name,
                "rows": total_rows,
                "cols": cols,
                "message": f"Word 문서에 표준 양식 '{form_name}'({total_rows}x{cols})이 삽입되었습니다.",
            }
        except Exception as e:
            logger.exception("Word 양식 삽입 실패")
            return {"success": False, "target_app": "word", "error": f"Word 표 삽입 실패: {e}"}
