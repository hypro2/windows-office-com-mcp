# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_renderers/word_renderer.py
MS Word 청사진을 데스크톱 COM으로 100% 동일하게 실시간 재생(Playback)하는 렌더러.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com, word_undo_group

logger = logging.getLogger("mcp_server.word_renderer")


class WordBlueprintRenderer:
    """Word 사용자 청사진 실시간 COM 렌더러"""

    @classmethod
    def render(
        cls,
        blueprint: dict[str, Any],
        payload: dict[str, Any],
        output_path: str | Path | None = None,
    ) -> dict[str, Any]:
        ensure_com()
        word_app = com_manager.get_or_create_word_app(visible=True)
        doc = word_app.Documents.Add(Visible=True)

        canvas = blueprint.get("canvas", {})
        ps = doc.Sections(1).PageSetup
        if "top_margin" in canvas:
            ps.TopMargin = float(canvas["top_margin"])
        if "bottom_margin" in canvas:
            ps.BottomMargin = float(canvas["bottom_margin"])
        if "left_margin" in canvas:
            ps.LeftMargin = float(canvas["left_margin"])
        if "right_margin" in canvas:
            ps.RightMargin = float(canvas["right_margin"])

        total_elements = 0
        pages = blueprint.get("pages", [])

        with word_undo_group(doc, "Render User Template"):
            sel = word_app.Selection
            for page in pages:
                elements = page.get("elements", [])
                for el in elements:
                    el_type = el.get("type", "paragraph")
                    slot_id = el.get("slot_id")

                    if el_type == "paragraph":
                        text_val = el.get("text", "")
                        if slot_id and slot_id in payload:
                            text_val = str(payload[slot_id])

                        p_range = sel.Range
                        p_range.Text = text_val + "\r"
                        if el.get("font_name"):
                            p_range.Font.Name = el["font_name"]
                        if el.get("font_size"):
                            p_range.Font.Size = float(el["font_size"])
                        if el.get("bold") is not None:
                            p_range.Font.Bold = -1 if el["bold"] else 0
                        if el.get("alignment") is not None:
                            try:
                                p_range.ParagraphFormat.Alignment = int(el["alignment"])
                            except Exception:
                                pass

                        sel.Collapse(0)  # wdCollapseEnd = 0
                        total_elements += 1

                    elif el_type == "table":
                        t_data = el.get("table_data", [])
                        # 페이로드에서 새 표 데이터가 제공되었는지 확인
                        if slot_id and slot_id in payload and isinstance(payload[slot_id], list):
                            t_data = payload[slot_id]

                        rows = len(t_data)
                        cols = el.get("cols_count", len(t_data[0]) if t_data else 1)
                        if rows > 0 and cols > 0:
                            tbl = doc.Tables.Add(sel.Range, rows, cols)
                            # 셀 너비 복원
                            col_widths = el.get("col_widths", [])
                            for c_idx, w in enumerate(col_widths, start=1):
                                if c_idx <= tbl.Columns.Count:
                                    try:
                                        tbl.Columns(c_idx).Width = float(w)
                                    except Exception:
                                        pass

                            for r_idx, row in enumerate(t_data, start=1):
                                for c_idx, cell_item in enumerate(row, start=1):
                                    if c_idx <= cols:
                                        c_text = cell_item.get("text", "") if isinstance(cell_item, dict) else str(cell_item)
                                        cell_range = tbl.Cell(r_idx, c_idx).Range
                                        cell_range.Text = c_text
                                        if isinstance(cell_item, dict) and cell_item.get("bg_color"):
                                            tbl.Cell(r_idx, c_idx).Shading.BackgroundPatternColor = int(cell_item["bg_color"])

                            sel.SetRange(tbl.Range.End, tbl.Range.End)
                            sel.TypeParagraph()
                            total_elements += 1

        out_file = None
        if output_path:
            out_file = str(Path(output_path).resolve())
            doc.SaveAs2(out_file)
            logger.info(f"Word 청사진 문서 저장 완료: {out_file}")

        try:
            word_app.Visible = True
            doc.Activate()
        except Exception:
            pass

        return {
            "success": True,
            "doc_type": "word",
            "template_name": blueprint.get("template_name", ""),
            "elements_count": total_elements,
            "output_path": out_file,
            "message": f"Word 문서({total_elements}개 요소)가 성공적으로 복제 생성되었습니다.",
        }

