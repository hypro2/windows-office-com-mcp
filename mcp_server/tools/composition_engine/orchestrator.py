# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/orchestrator.py
문서 구성 엔진 통합 오케스트레이터 및 실시간 데스크톱 오피스 렌더링 브릿지.
"""

from __future__ import annotations

import logging
from typing import Any

from ...com_manager import com_manager, ensure_com
from ..ppt_tools import ppt_freeze_display
from ..excel_tools import excel_freeze_display

logger = logging.getLogger("mcp_server.composition_engine")


class CompositionOrchestrator:
    """문서 구성 기획부터 실시간 COM 렌더링까지 총괄하는 통합 오케스트레이터"""

    # -------------------------------------------------------------------------
    # 1. PPT Real-time Composition Rendering
    # -------------------------------------------------------------------------
    @classmethod
    def render_ppt_composition(cls, plan: dict[str, Any], new_presentation: bool = True) -> dict[str, Any]:
        """기획된 PPT 슬라이드 구성안을 데스크톱 PowerPoint 창에 시각적 아키타입으로 렌더링"""
        ensure_com()
        ppt_app = com_manager.get_or_create_ppt_app(visible=True)

        if new_presentation or ppt_app.Presentations.Count == 0:
            prs = ppt_app.Presentations.Add(WithWindow=True)
            prs.PageSetup.SlideWidth = 960.0
            prs.PageSetup.SlideHeight = 540.0
        else:
            prs = ppt_app.ActivePresentation

        slides_data = plan.get("slides", [])
        created_count = 0

        with ppt_freeze_display(ppt_app):
            for s_idx, s_data in enumerate(slides_data, start=1):
                slide = prs.Slides.Add(prs.Slides.Count + 1, 12)  # ppLayoutBlank = 12
                created_count += 1

                # 1. 카테고리 뱃지 & Action Title
                badge_text = s_data.get("category_badge", "임원 보고")
                action_title = s_data.get("action_title", "")

                # 상단 뱃지
                badge = slide.Shapes.AddShape(1, 40, 30, 80, 22)  # msoShapeRectangle = 1
                badge.Fill.Solid()
                badge.Fill.ForeColor.RGB = 0x002060  # Navy
                badge.Line.Visible = False
                badge.TextFrame.TextRange.Text = badge_text
                badge.TextFrame.TextRange.Font.Size = 10
                badge.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
                badge.TextFrame.TextRange.Font.Bold = True

                # Action Title (결론형 리드문)
                title_box = slide.Shapes.AddTextbox(1, 40, 58, 880, 45)
                title_box.TextFrame.WordWrap = True
                title_box.TextFrame.TextRange.Text = action_title
                title_box.TextFrame.TextRange.Font.Size = 18
                title_box.TextFrame.TextRange.Font.Bold = True
                title_box.TextFrame.TextRange.Font.Color.RGB = 0x222222

                # 2. 아키타입별 컴포넌트 렌더링
                arch = s_data.get("slide_type", "kpi_grid_3")
                items = s_data.get("items", [])

                if arch == "kpi_grid_3":
                    # 3개 균등 카드 배치 (Left: 40, Width: 275, Gap: 27)
                    card_w = 275
                    card_h = 320
                    gap = 27
                    top = 120
                    for i, it in enumerate(items[:3]):
                        left = 40 + i * (card_w + gap)
                        # 카드 배경
                        card = slide.Shapes.AddShape(1, left, top, card_w, card_h)
                        card.Fill.Solid()
                        card.Fill.ForeColor.RGB = 0xF8F9FA
                        card.Line.ForeColor.RGB = 0xE0E0E0

                        # 카드 내부 상단 라벨
                        lbl = slide.Shapes.AddTextbox(1, left + 15, top + 15, card_w - 30, 25)
                        lbl.TextFrame.TextRange.Text = str(it.get("label", f"지표 {i+1}"))
                        lbl.TextFrame.TextRange.Font.Size = 13
                        lbl.TextFrame.TextRange.Font.Color.RGB = 0x777777

                        # 카드 대형 숫자 (KPI)
                        val = slide.Shapes.AddTextbox(1, left + 15, top + 45, card_w - 30, 50)
                        val.TextFrame.TextRange.Text = str(it.get("value", "0"))
                        val.TextFrame.TextRange.Font.Size = 36
                        val.TextFrame.TextRange.Font.Bold = True
                        val.TextFrame.TextRange.Font.Color.RGB = 0x002060

                        # 카드 부연 설명
                        sub = slide.Shapes.AddTextbox(1, left + 15, top + 110, card_w - 30, 180)
                        sub.TextFrame.WordWrap = True
                        sub.TextFrame.TextRange.Text = str(it.get("desc") or it.get("subtext") or "")
                        sub.TextFrame.TextRange.Font.Size = 12
                        sub.TextFrame.TextRange.Font.Color.RGB = 0x444444

                elif arch == "process_horizontal_4":
                    # 4단계 가로형 프로세스
                    card_w = 205
                    card_h = 320
                    gap = 20
                    top = 120
                    for i, it in enumerate(items[:4]):
                        left = 40 + i * (card_w + gap)
                        card = slide.Shapes.AddShape(1, left, top, card_w, card_h)
                        card.Fill.Solid()
                        card.Fill.ForeColor.RGB = 0xF5F7FA
                        card.Line.ForeColor.RGB = 0x002060 if i == 0 else 0xCCCCCC

                        # Step 라벨
                        step_lbl = slide.Shapes.AddTextbox(1, left + 12, top + 15, card_w - 24, 25)
                        step_lbl.TextFrame.TextRange.Text = f"STEP 0{i+1}"
                        step_lbl.TextFrame.TextRange.Font.Size = 11
                        step_lbl.TextFrame.TextRange.Font.Bold = True
                        step_lbl.TextFrame.TextRange.Font.Color.RGB = 0x002060

                        # 단계 제목
                        stitle = slide.Shapes.AddTextbox(1, left + 12, top + 42, card_w - 24, 40)
                        stitle.TextFrame.WordWrap = True
                        stitle.TextFrame.TextRange.Text = str(it.get("title", f"단계 {i+1}"))
                        stitle.TextFrame.TextRange.Font.Size = 14
                        stitle.TextFrame.TextRange.Font.Bold = True

                        # 세부 내용
                        sdesc = slide.Shapes.AddTextbox(1, left + 12, top + 90, card_w - 24, 200)
                        sdesc.TextFrame.WordWrap = True
                        sdesc.TextFrame.TextRange.Text = str(it.get("desc") or "")
                        sdesc.TextFrame.TextRange.Font.Size = 11
                        sdesc.TextFrame.TextRange.Font.Color.RGB = 0x555555

                # 3. 하단 각주 (Footnote)
                fn = s_data.get("footnote")
                if fn:
                    fn_box = slide.Shapes.AddTextbox(1, 40, 490, 880, 25)
                    fn_box.TextFrame.TextRange.Text = f"※ {fn}"
                    fn_box.TextFrame.TextRange.Font.Size = 9
                    fn_box.TextFrame.TextRange.Font.Color.RGB = 0x888888

        return {
            "status": "success",
            "message": f"PPT 구성안 총 {created_count}개 슬라이드가 실시간으로 데스크톱에 생성되었습니다.",
            "created_slides": created_count,
        }

    # -------------------------------------------------------------------------
    # 2. Excel Real-time Composition Rendering
    # -------------------------------------------------------------------------
    @classmethod
    def render_excel_composition(cls, plan: dict[str, Any], new_workbook: bool = True) -> dict[str, Any]:
        """기획된 Excel 토폴로지(대시보드 KPI 타일 + 본문 그리드 + 수식)를 데스크톱 Excel 창에 실시간 생성"""
        ensure_com()
        excel_app = com_manager.get_or_create_excel_app(visible=True)

        if new_workbook or excel_app.Workbooks.Count == 0:
            wb = excel_app.Workbooks.Add()
        else:
            wb = excel_app.ActiveWorkbook

        ws = wb.ActiveSheet
        spec = plan.get("spec", {})

        with excel_freeze_display(excel_app, pause_calculation=True):
            # A열 좌측 마진 축소 (너비 3)
            ws.Columns(1).ColumnWidth = 3

            # Row 1: 타이틀
            ws.Cells(2, 2).Value = spec.get("title", "실적 보고서")
            ws.Cells(2, 2).Font.Size = 16
            ws.Cells(2, 2).Font.Bold = True

            # Row 2~4: 상단 KPI 요약 타일 배치
            kpis = spec.get("kpi_tiles", [])
            for k_idx, kpi in enumerate(kpis[:4]):
                start_c = 2 + (k_idx * 2)
                # 타일 제목
                lbl_cell = ws.Cells(4, start_c)
                lbl_cell.Value = kpi.get("label", "KPI")
                lbl_cell.Font.Size = 9
                lbl_cell.Font.Color = 0x777777

                # 타일 수치
                val_cell = ws.Cells(5, start_c)
                val_cell.Value = kpi.get("value", "0")
                val_cell.Font.Size = 16
                val_cell.Font.Bold = True
                val_cell.Font.Color = 0x002060  # Navy

                # 배경 음영
                rng = ws.Range(ws.Cells(4, start_c), ws.Cells(5, start_c + 1))
                rng.Interior.Color = 0xF5F7FA
                rng.Borders.Color = 0xDDDDDD

            # Row 7: 단위 표기
            ws.Cells(7, 2).Value = spec.get("unit", "단위: 백만원, %")
            ws.Cells(7, 2).Font.Size = 9
            ws.Cells(7, 2).Font.Color = 0x777777

            # Row 8: 메인 테이블 헤더
            grid = spec.get("grid", {})
            headers = grid.get("headers", [])
            start_row = 8
            for h_idx, h in enumerate(headers):
                cell = ws.Cells(start_row, 2 + h_idx)
                cell.Value = h
                cell.Font.Bold = True
                cell.Font.Color = 0xFFFFFF
                cell.Interior.Color = 0x002060  # Dark Navy
                cell.HorizontalAlignment = -4108  # xlCenter

            # Row 9 ~ N: 데이터 행 입력
            data_rows = grid.get("data_rows", [])
            for r_idx, row in enumerate(data_rows):
                cur_r = start_row + 1 + r_idx
                for c_idx, val in enumerate(row):
                    cell = ws.Cells(cur_r, 2 + c_idx)
                    cell.Value = val
                    if isinstance(val, (int, float)):
                        cell.NumberFormat = "#,##0" if val >= 1000 else "0.0%" if isinstance(val, float) and val <= 1.0 else "0"
                    if r_idx % 2 == 1:
                        cell.Interior.Color = 0xFAFAFA

            # 마지막 행: 집계 수식 (SUM)
            summary_r = start_row + 1 + len(data_rows)
            ws.Cells(summary_r, 2).Value = "합계"
            ws.Cells(summary_r, 2).Font.Bold = True
            ws.Cells(summary_r, 2).HorizontalAlignment = -4108

            for c_idx in range(1, len(headers)):
                target_col_letter = chr(65 + 1 + c_idx)  # B열 = 66
                cell = ws.Cells(summary_r, 2 + c_idx)
                cell.Formula = f"=SUM({target_col_letter}{start_row + 1}:{target_col_letter}{summary_r - 1})"
                cell.Font.Bold = True
                cell.NumberFormat = "#,##0"

            # 테두리 설정
            total_range = ws.Range(ws.Cells(start_row, 2), ws.Cells(summary_r, 2 + len(headers) - 1))
            total_range.Borders.Color = 0xCCCCCC

            # 열 너비 자동 맞춤
            for c_idx in range(len(headers)):
                ws.Columns(2 + c_idx).AutoFit()
                if ws.Columns(2 + c_idx).ColumnWidth < 12:
                    ws.Columns(2 + c_idx).ColumnWidth = 14

            # 틀 고정 & 인쇄 설정
            try:
                ws.Cells(start_row + 1, 2).Select()
                excel_app.ActiveWindow.FreezePanes = True
                ws.PageSetup.Orientation = 2  # xlLandscape
                ws.PageSetup.FitToPagesWide = 1
                ws.PageSetup.FitToPagesTall = 1
            except Exception:
                pass

        return {
            "status": "success",
            "message": "Excel 대시보드 토폴로지가 성공적으로 데스크톱에 생성되었습니다.",
            "data_rows_inserted": len(data_rows),
            "summary_row": summary_r,
        }

    # -------------------------------------------------------------------------
    # 3. HWP Real-time Composition Rendering
    # -------------------------------------------------------------------------
    @classmethod
    def render_hwp_composition(cls, plan: dict[str, Any]) -> dict[str, Any]:
        """기획된 한글 A4 1장 보고서(요약 박스 + 내어쓰기 개조식)를 HWP ROT에 실시간 입력"""
        hwp = com_manager.get_hwp_app()
        if not hwp:
            return {"status": "error", "message": "실행 중인 한컴오피스(HWP) 창을 찾을 수 없습니다."}

        spec = plan.get("spec", {})

        # 1. 문서 제목
        title = spec.get("title", "기획 보고서")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = f"\n  {title}\n\n"
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 2. 상단 요약 박스 (1행 1열 표 생성)
        sum_items = spec.get("executive_summary_box", [])
        if sum_items:
            hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
            hwp.HParameterSet.HTableCreation.Rows = 1
            hwp.HParameterSet.HTableCreation.Cols = 1
            hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

            # 요약 내용 입력
            summary_text = " [핵심 요약]\n" + "\n".join([f" • {it}" for it in sum_items])
            hwp.HParameterSet.HInsertText.Text = summary_text
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            # 표 밖으로 빠져나오기
            hwp.HAction.Run("CloseEx")
            hwp.HAction.Run("MoveDown")

        # 3. 본문 개조식 섹션 입력
        sections = spec.get("sections", [])
        for sec in sections:
            sec_title = sec.get("title", "")
            hwp.HParameterSet.HInsertText.Text = f"\n\n□ {sec_title}\n"
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

            for item in sec.get("items", []):
                sym = item.get("symbol", "○")
                txt = item.get("text", "")
                hwp.HParameterSet.HInsertText.Text = f"  {sym} {txt}\n"
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        return {
            "status": "success",
            "message": "한글 A4 1장 표준 개조식 보고서가 데스크톱 HWP에 실시간 생성되었습니다.",
        }

    # -------------------------------------------------------------------------
    # 4. DOCX Real-time Composition Rendering
    # -------------------------------------------------------------------------
    @classmethod
    def render_docx_composition(cls, plan: dict[str, Any], new_document: bool = True) -> dict[str, Any]:
        """기획된 워드 4대 표준 블록 품의서를 Word 창에 실시간 생성"""
        ensure_com()
        word_app = com_manager.get_or_create_word_app(visible=True)

        if new_document or word_app.Documents.Count == 0:
            doc = word_app.Documents.Add()
        else:
            doc = word_app.ActiveDocument

        spec = plan.get("spec", {})
        blocks = spec.get("blocks", {})

        sel = word_app.Selection

        # 1. 제목
        title = spec.get("title", "사업 기획 및 품의서")
        sel.Font.Size = 20
        sel.Font.Bold = True
        sel.Font.Color = 0x002060
        sel.TypeText(f"{title}\n\n")

        # 2. Block 1: Executive Summary Box
        sel.Font.Size = 11
        sel.Font.Bold = True
        sel.TypeText("■ Executive Summary (핵심 결론 요약)\n")
        sel.Font.Bold = False
        for s in blocks.get("block1_summary", []):
            sel.TypeText(f"  • {s}\n")
        sel.TypeText("\n")

        # 3. Block 2: 현행 대비 개선안 대조표
        sel.Font.Size = 14
        sel.Font.Bold = True
        sel.TypeText("1. 현행 문제점 및 개선 방향\n")
        comp = blocks.get("block2_comparison", {})
        headers = comp.get("headers", ["구분", "현행(As-Is)", "개선안(To-Be)"])
        rows = comp.get("rows", [["프로세스", "수작업 수기 집계", "AI MCP 자동화 파이프라인"]])

        table = doc.Tables.Add(sel.Range, NumRows=len(rows) + 1, NumColumns=len(headers))
        table.Borders.Enable = True
        table.Rows(1).Range.Font.Bold = True
        table.Rows(1).Range.Shading.BackgroundPatternColor = 0x002060
        table.Rows(1).Range.Font.Color = 0xFFFFFF

        for c_idx, h in enumerate(headers, start=1):
            table.Cell(1, c_idx).Range.Text = h

        for r_idx, r in enumerate(rows, start=2):
            for c_idx, val in enumerate(r, start=1):
                table.Cell(r_idx, c_idx).Range.Text = str(val)

        sel.EndKey(6)  # wdStory = 6 (문서 끝으로 이동)
        sel.TypeText("\n\n")

        # 4. Block 4: 기대효과
        sel.Font.Size = 14
        sel.Font.Bold = True
        sel.TypeText("2. 기대효과 및 정량 지표\n")
        sel.Font.Bold = False
        sel.Font.Size = 11
        for eff in blocks.get("block4_effects", []):
            sel.TypeText(f"  - {eff}\n")

        return {
            "status": "success",
            "message": "워드 4대 블록 품의서가 성공적으로 데스크톱 Word 창에 실시간 생성되었습니다.",
        }
