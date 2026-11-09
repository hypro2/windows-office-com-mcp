# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_analyzers/word_analyzer.py
MS Word (DOCX) 여백, 스타일 위계, 단락, 표 규격 역공학 분석기.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com
from .base_analyzer import BaseDocumentAnalyzer

logger = logging.getLogger("mcp_server.word_analyzer")


class WordDocumentAnalyzer(BaseDocumentAnalyzer):
    doc_type: str = "word"

    def analyze_file(self, file_path: Path) -> dict[str, Any]:
        """DOCX 파일을 열어 용지 여백, 스타일, 단락 및 표 구조 추출 (DRM 친화적 COM 개방)"""
        ensure_com()
        word_app = com_manager.get_or_create_word_app(visible=True)
        try:
            word_app.DisplayAlerts = 0  # wdAlertsNone = 0
        except Exception:
            pass

        abs_path = str(file_path.resolve())

        doc = None
        try:
            # DRM 복호화 통과를 위해 Visible=True 및 서식 변환 확인 팝업 차단(ConfirmConversions=False)
            doc = word_app.Documents.Open(
                abs_path,
                ReadOnly=True,
                Visible=True,
                ConfirmConversions=False,
            )
            return self._extract_from_document(doc, file_path.stem)
        finally:
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass

    def analyze_active_window(self) -> dict[str, Any]:
        """현재 모니터에 열려 있는 활성 Word 문서 분석"""
        ensure_com()
        word_app, doc = com_manager.get_word(create_if_missing=False)
        if not word_app or not doc:
            raise RuntimeError("현재 활성화된 Word 문서 창을 찾을 수 없습니다.")

        name = getattr(doc, "Name", "ActiveDocument")
        return self._extract_from_document(doc, Path(name).stem)

    def _extract_from_document(self, doc: Any, source_name: str) -> dict[str, Any]:
        sec = doc.Sections(1)
        ps = sec.PageSetup

        page_setup = {
            "page_width": round(float(ps.PageWidth), 1),
            "page_height": round(float(ps.PageHeight), 1),
            "top_margin": round(float(ps.TopMargin), 1),
            "bottom_margin": round(float(ps.BottomMargin), 1),
            "left_margin": round(float(ps.LeftMargin), 1),
            "right_margin": round(float(ps.RightMargin), 1),
        }

        elements: list[dict[str, Any]] = []
        fonts: set[str] = set()

        # 1. 단락 추출 (최대 100단락)
        p_count = min(100, doc.Paragraphs.Count)
        for p_idx in range(1, p_count + 1):
            p = doc.Paragraphs(p_idx)
            text = (p.Range.Text or "").strip()
            if not text:
                continue

            style_name = str(p.Style.NameLocal)
            f_name = str(p.Range.Font.Name)
            fonts.add(f_name)
            f_size = round(float(p.Range.Font.Size), 1)
            f_bold = bool(p.Range.Font.Bold)
            align = int(p.Alignment)

            is_heading = any(k in style_name.lower() for k in ["heading", "제목", "title"])

            elements.append({
                "element_id": f"para_{p_idx}",
                "type": "paragraph",
                "style_name": style_name,
                "text": text,
                "font_name": f_name,
                "font_size": f_size,
                "bold": f_bold,
                "alignment": align,
                "is_static": is_heading,
                "slot_id": f"slot_para_{p_idx}" if not is_heading else f"slot_title_{p_idx}",
                "label": f"[{style_name}] {text[:20]}",
            })

        # 2. 표 추출 (최대 10개 표)
        t_count = min(10, doc.Tables.Count)
        for t_idx in range(1, t_count + 1):
            tbl = doc.Tables(t_idx)
            rows = tbl.Rows.Count
            cols = tbl.Columns.Count

            col_widths = []
            for c in range(1, cols + 1):
                try:
                    col_widths.append(round(float(tbl.Columns(c).Width), 1))
                except Exception:
                    col_widths.append(100.0)

            table_rows_data = []
            for r in range(1, min(20, rows + 1)):
                row_cells = []
                for c in range(1, cols + 1):
                    try:
                        c_text = (tbl.Cell(r, c).Range.Text or "").replace("\r\x07", "").strip()
                        shd = tbl.Cell(r, c).Shading.BackgroundPatternColor
                        row_cells.append({
                            "text": c_text,
                            "bg_color": int(shd) if shd != -16777216 else None,
                        })
                    except Exception:
                        row_cells.append({"text": "", "bg_color": None})
                table_rows_data.append(row_cells)

            elements.append({
                "element_id": f"table_{t_idx}",
                "type": "table",
                "rows_count": rows,
                "cols_count": cols,
                "col_widths": col_widths,
                "table_data": table_rows_data,
                "is_static": False,
                "slot_id": f"slot_table_{t_idx}",
                "label": f"{rows}x{cols} 데이터 표 {t_idx}",
            })

        return {
            "doc_type": "word",
            "source_name": source_name,
            "canvas": page_setup,
            "design_tokens": {
                "fonts": list(fonts)[:5],
            },
            "pages": [{
                "page_index": 1,
                "title": "본문",
                "elements": elements,
            }],
        }

