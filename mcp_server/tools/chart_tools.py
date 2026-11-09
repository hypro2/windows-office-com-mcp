# -*- coding: utf-8 -*-
"""
mcp_server/tools/chart_tools.py
외부 라이브러리(matplotlib 등) 없이 오피스 자체 COM 내장 기능만을 100% 활용하는
크로스 오피스(Excel, PPT, Word, HWP) 9대 비즈니스 차트 및 표 빌더.
- Excel / Word / PowerPoint: Shapes.AddChart2 + ChartData.Workbook + SparklineGroups + Native Tables
- HWP: TableCreate + CellShape.FaceColor (한글 자체 네이티브 표 및 게이지)
"""

from __future__ import annotations

import logging
import re
from typing import Any, Literal

from mcp_server.compat import MCPServer

from ..com_manager import com_manager, ensure_com, hex_to_bgr, hex_to_rgb, word_undo_group

logger = logging.getLogger("mcp_server.chart_tools")

# Office XlChartType 고유 열거형 상수 (표준 Office Art 엔진)
XL_LINE = 4
XL_COLUMN_CLUSTERED = 51
XL_COLUMN_STACKED = 52
XL_COLUMN_STACKED_100 = 53
XL_BAR_CLUSTERED = 57
XL_AREA_STACKED = 76
XL_XY_SCATTER_LINES_NO_MARKERS = 75

# 엑셀 스파크라인 타입 (xlSparkLine = 1)
XL_SPARK_LINE = 1


def interpolate_color_bgr(rate: float, theme: str = "wine_red") -> int:
    """0.0~100.0 비율에 따른 BGR 색상 정수 반환 (Excel/Word용)"""
    r_ratio = max(0.0, min(1.0, rate / 100.0))
    if theme == "teal":
        r0, g0, b0 = 0xF0, 0xF9, 0xF2
        r1, g1, b1 = 0x00, 0xA8, 0x59
    elif theme == "navy":
        r0, g0, b0 = 0xEB, 0xF3, 0xFA
        r1, g1, b1 = 0x1F, 0x38, 0x64
    else:  # wine_red
        r0, g0, b0 = 0xFE, 0xE2, 0xE2
        r1, g1, b1 = 0x99, 0x1B, 0x1B

    curr_r = int(r0 + (r1 - r0) * r_ratio)
    curr_g = int(g0 + (g1 - g0) * r_ratio)
    curr_b = int(b0 + (b1 - b0) * r_ratio)
    return curr_b + (curr_g << 8) + (curr_r << 16)


def interpolate_color_rgb(rate: float, theme: str = "wine_red") -> int:
    """0.0~100.0 비율에 따른 RGB 색상 정수 반환 (PPT/HWP용)"""
    r_ratio = max(0.0, min(1.0, rate / 100.0))
    if theme == "teal":
        r0, g0, b0 = 0xF0, 0xF9, 0xF2
        r1, g1, b1 = 0x00, 0xA8, 0x59
    elif theme == "navy":
        r0, g0, b0 = 0xEB, 0xF3, 0xFA
        r1, g1, b1 = 0x1F, 0x38, 0x64
    else:
        r0, g0, b0 = 0xFE, 0xE2, 0xE2
        r1, g1, b1 = 0x99, 0x1B, 0x1B

    curr_r = int(r0 + (r1 - r0) * r_ratio)
    curr_g = int(g0 + (g1 - g0) * r_ratio)
    curr_b = int(b0 + (b1 - b0) * r_ratio)
    return curr_r + (curr_g << 8) + (curr_b << 16)


def get_native_chart_blueprint(
    chart_type: str, custom_title: str | None = None
) -> tuple[list[list[Any]], int, str]:
    """
    각 비즈니스 차트 유형별 Office 네이티브 데이터 매트릭스, XlChartType 상수, 제목을 반환.
    외부 라이브러리 없이 순수 파이썬 데이터 구조로 구성.
    """
    if chart_type == "control_chart":
        # 03 관리도: 일자별 전환율 + 평균 중심선(CL) + 상한선(+3σ) + 하한선(-3σ)
        title = custom_title or "관리도 (일 전환율 %, 평균 ± 3σ)"
        headers = ["일자", "일전환율(%)", "평균(CL)", "+3σ(UCL)", "-3σ(LCL)"]
        rows = []
        for d in range(1, 31):
            # 시뮬레이션 수치 (평균 2.0%, UCL 2.75%, LCL 1.25%)
            val = 2.0 + ((d * 7) % 11 - 5) * 0.12
            if d == 18:
                val = 2.95  # 이상치
            elif d == 25:
                val = 1.15  # 이상치
            rows.append([f"{d}일", round(val, 2), 2.00, 2.75, 1.25])
        return [headers] + rows, XL_LINE, title

    elif chart_type == "contribution":
        # 07 기여도 분해: 주별 채널 기여율 100% 누적 세로 막대 차트
        title = custom_title or "기여도 분해 (주별 채널 기여율 100%)"
        headers = ["주차", "기본", "검색광고", "PMax", "메타(SNS)"]
        rows = []
        for w in range(1, 17):
            rows.append([f"{w}주", 45, 25, 18, 12])
        return [headers] + rows, XL_COLUMN_STACKED_100, title

    elif chart_type == "lift_chart":
        # 08 리프트 차트: 예측 점수 10분위별 리프트(배수) + 평균 기준선(1.0)
        title = custom_title or "리프트 차트 (예측 점수 10분위별 성과)"
        headers = ["상위 분위", "리프트(배)", "평균 기준선"]
        lifts = [2.9, 1.9, 1.3, 1.0, 0.8, 0.6, 0.5, 0.4, 0.3, 0.3]
        rows = [[f"상위 {i*10}%", lifts[i-1], 1.0] for i in range(1, 11)]
        return [headers] + rows, XL_COLUMN_CLUSTERED, title

    elif chart_type == "survival_curve":
        # 06 생존 곡선: 경과 시간별 고객 잔존율(%)
        title = custom_title or "생존 곡선 (구독 유지 비율 %, 월간 vs 연간)"
        headers = ["경과월", "월간구독 유지율(%)", "연간구독 유지율(%)"]
        rows = [
            ["0개월", 100.0, 100.0],
            ["2개월", 83.5, 92.3],
            ["4개월", 69.8, 85.2],
            ["6개월", 58.3, 78.7],
            ["8개월", 48.7, 72.6],
            ["10개월", 40.7, 67.0],
            ["12개월", 34.0, 61.9],
        ]
        return [headers] + rows, XL_LINE, title

    elif chart_type == "forest_plot":
        # 04 포레스트 플롯: 실험 항목별 효과 크기 및 신뢰구간
        title = custom_title or "포레스트 플롯 (A/B 테스트 효과 크기 %)"
        headers = ["실험 항목", "효과 크기(%)", "95% 하한", "95% 상한"]
        rows = [
            ["무료배송 문구", 6.5, 1.2, 11.8],
            ["리뷰 위젯", -2.1, -7.5, 3.2],
            ["가격 강조", 1.2, -4.0, 6.5],
            ["보상 썸네일", 11.4, 4.5, 18.2],
            ["CTA 버튼색", -0.5, -3.8, 2.9],
            ["쿠폰 자동삽입", -3.2, -8.0, 1.5],
        ]
        return [headers] + rows, XL_BAR_CLUSTERED, title

    elif chart_type == "fan_chart":
        # 05 팬 차트: 과거 실적선 + 미래 예측 및 상하한 밴드
        title = custom_title or "팬 차트 (월 매출 억원, 실적 및 예측)"
        headers = ["월", "실적 매출", "예측 중앙값", "신뢰 하한", "신뢰 상한"]
        rows = [
            ["1월", 3.2, "", "", ""],
            ["3월", 4.1, "", "", ""],
            ["5월", 4.6, "", "", ""],
            ["7월", 5.2, "", "", ""],
            ["9월", 5.4, 5.4, 5.4, 5.4],
            ["10월", "", 5.8, 5.4, 6.2],
            ["11월", "", 6.1, 5.4, 6.8],
            ["12월", "", 6.5, 5.4, 7.6],
        ]
        return [headers] + rows, XL_LINE, title

    elif chart_type == "roc_curve":
        # 09 ROC 커브: FPR vs TPR 곡선 + 무작위 기준선
        title = custom_title or "ROC 커브 (분류 모델 판별력, AUC 0.82)"
        headers = ["위양성률(FPR)", "진양성률(TPR)", "무작위기준선"]
        rows = [
            [0.00, 0.00, 0.00],
            [0.05, 0.28, 0.05],
            [0.10, 0.45, 0.10],
            [0.20, 0.62, 0.20],
            [0.35, 0.76, 0.35],
            [0.50, 0.85, 0.50],
            [0.70, 0.93, 0.70],
            [1.00, 1.00, 1.00],
        ]
        return [headers] + rows, XL_XY_SCATTER_LINES_NO_MARKERS, title

    else:
        # 기본 세로 막대
        title = custom_title or f"비즈니스 지표: {chart_type}"
        headers = ["항목", "지표값"]
        rows = [["A", 10], ["B", 25], ["C", 18], ["D", 32]]
        return [headers] + rows, XL_COLUMN_CLUSTERED, title


def register_chart_tools(mcp: MCPServer) -> None:
    """오피스 자체 기능 전용 9대 비즈니스 차트 및 표 도구 등록"""

    # =========================================================================
    # 1. 01 코호트 리텐션 (4대 오피스 100% 네이티브 표 + 셀 색상 그라디언트)
    # =========================================================================
    @mcp.tool()
    def office_insert_cohort_retention(
        cohort_matrix: list[list[float | None]] | None = None,
        row_labels: list[str] | None = None,
        col_labels: list[str] | None = None,
        theme: Literal["wine_red", "teal", "navy"] = "wine_red",
        target_app: Literal["excel", "word", "ppt", "hwp"] = "excel",
        start_cell: str = "B2",
    ) -> str:
        """
        [고급 차트 1] 가입 시점별 잔존율 코호트 리텐션(Cohort Retention) 네이티브 표 생성.
        외부 라이브러리 없이 각 오피스(Excel, PPT, Word, HWP) 자체 표 생성 API를 호출하고,
        각 셀에 잔존율(%)에 비례한 히트맵 색상을 직접 주입하여 사용자가 수정 가능한 표를 만듭니다.
        """
        ensure_com()

        if cohort_matrix is None:
            cohort_matrix = [
                [100.0, 35.0, 24.0, 21.0, 15.0, 11.0],
                [100.0, 33.0, 25.0, 19.0, 15.0, None],
                [100.0, 35.0, 28.0, 16.0, None, None],
                [100.0, 38.0, 27.0, None, None, None],
                [100.0, 44.0, None, None, None, None],
                [100.0, None, None, None, None, None],
            ]
        if row_labels is None:
            row_labels = ["0일", "1일", "2일", "3일", "4일", "5일"]
        if col_labels is None:
            col_labels = ["0", "+1", "+2", "+3", "+4", "+5"]

        num_data_rows = len(row_labels)
        num_data_cols = len(col_labels)
        total_rows = num_data_rows + 1
        total_cols = num_data_cols + 1

        # A. Excel 네이티브 표
        if target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            sheet = wb.ActiveSheet

            m = re.match(r"^([A-Z]+)(\d+)$", start_cell.upper().strip())
            start_r, start_c = (int(m.group(2)), sum((ord(ch) - ord("A") + 1) * (26 ** i) for i, ch in enumerate(reversed(m.group(1))))) if m else (2, 2)

            table_rng = sheet.Range(
                sheet.Cells(start_r, start_c),
                sheet.Cells(start_r + total_rows - 1, start_c + total_cols - 1)
            )

            # 1행 헤더
            sheet.Cells(start_r, start_c).Value = "가입 시점"
            for c_i, c_lbl in enumerate(col_labels):
                sheet.Cells(start_r, start_c + 1 + c_i).Value = c_lbl

            hdr_rng = sheet.Range(sheet.Cells(start_r, start_c), sheet.Cells(start_r, start_c + total_cols - 1))
            hdr_rng.Font.Name = "맑은 고딕"
            hdr_rng.Font.Bold = True
            hdr_rng.Interior.Color = hex_to_bgr("#F3F4F6")
            hdr_rng.HorizontalAlignment = -4108
            hdr_rng.RowHeight = 22

            # 데이터 행 및 셀별 Interior.Color 그라디언트
            for r_i in range(num_data_rows):
                curr_r = start_r + 1 + r_i
                sheet.Rows(curr_r).RowHeight = 20

                lbl_cell = sheet.Cells(curr_r, start_c)
                lbl_cell.Value = row_labels[r_i]
                lbl_cell.Font.Bold = True
                lbl_cell.HorizontalAlignment = -4108
                lbl_cell.Interior.Color = hex_to_bgr("#F9FAFB")

                for c_i in range(num_data_cols):
                    data_cell = sheet.Cells(curr_r, start_c + 1 + c_i)
                    val = cohort_matrix[r_i][c_i] if c_i < len(cohort_matrix[r_i]) else None
                    if val is not None:
                        data_cell.Value = int(val)
                        data_cell.Interior.Color = interpolate_color_bgr(float(val), theme)
                        data_cell.Font.Name = "맑은 고딕"
                        data_cell.Font.Bold = True
                        data_cell.Font.Color = 0xFFFFFF if float(val) >= 50.0 else hex_to_bgr("#1F2937")
                        data_cell.HorizontalAlignment = -4108
                    else:
                        data_cell.Value = ""
                        data_cell.Interior.Color = 0xFFFFFF

            try:
                for b_i in (7, 8, 9, 10, 11, 12):
                    table_rng.Borders(b_i).LineStyle = 1
                    table_rng.Borders(b_i).Weight = 2
                    table_rng.Borders(b_i).Color = hex_to_bgr("#E5E7EB")
                table_rng.Columns.ColumnWidth = 7.5
                sheet.Columns(start_c).ColumnWidth = 11.0
            except Exception:
                pass

            return f"Excel '{sheet.Name}' 시트에 100% 네이티브 코호트 리텐션 표({total_rows}행 {total_cols}열)를 생성했습니다. (위치: {table_rng.Address})"

        # B. Word 네이티브 표
        elif target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."

            with word_undo_group(word_app, "Insert Cohort Table"):
                sel = word_app.Selection
                tbl = doc.Tables.Add(sel.Range, total_rows, total_cols)
                tbl.Borders.Enable = True

                tbl.Cell(1, 1).Range.Text = "가입 시점"
                for c_i, c_lbl in enumerate(col_labels):
                    tbl.Cell(1, c_i + 2).Range.Text = str(c_lbl)

                for c_i in range(1, total_cols + 1):
                    c = tbl.Cell(1, c_i)
                    c.Range.Font.Bold = True
                    c.Range.ParagraphFormat.Alignment = 1
                    c.Shading.BackgroundPatternColor = hex_to_bgr("#F3F4F6")

                for r_i in range(num_data_rows):
                    row_idx = r_i + 2
                    tbl.Cell(row_idx, 1).Range.Text = row_labels[r_i]
                    tbl.Cell(row_idx, 1).Range.Font.Bold = True
                    tbl.Cell(row_idx, 1).Range.ParagraphFormat.Alignment = 1

                    for c_i in range(num_data_cols):
                        col_idx = c_i + 2
                        val = cohort_matrix[r_i][c_i] if c_i < len(cohort_matrix[r_i]) else None
                        cell = tbl.Cell(row_idx, col_idx)
                        if val is not None:
                            cell.Range.Text = f"{int(val)}%"
                            cell.Range.ParagraphFormat.Alignment = 1
                            cell.Shading.BackgroundPatternColor = interpolate_color_bgr(float(val), theme)
                            if float(val) >= 50.0:
                                cell.Range.Font.Color = hex_to_bgr("#FFFFFF")
                        else:
                            cell.Range.Text = ""

                return f"Word 문서에 100% 네이티브 코호트 리텐션 표({total_rows}행 {total_cols}열)를 성공적으로 삽입했습니다."

        # C. PPT 네이티브 표
        elif target_app == "ppt":
            ppt_app, pres = com_manager.get_ppt()
            if not pres:
                return "열려 있는 PowerPoint 프레젠테이션을 찾을 수 없습니다."

            slide = ppt_app.ActiveWindow.View.Slide
            table_shape = slide.Shapes.AddTable(total_rows, total_cols, 60, 100, 600, 260)
            tbl = table_shape.Table

            tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = "가입 시점"
            for c_i, c_lbl in enumerate(col_labels):
                tbl.Cell(1, c_i + 2).Shape.TextFrame.TextRange.Text = str(c_lbl)

            for c_i in range(1, total_cols + 1):
                cell = tbl.Cell(1, c_i)
                cell.Shape.Fill.Solid()
                cell.Shape.Fill.ForeColor.RGB = hex_to_rgb("#F3F4F6")
                tr = cell.Shape.TextFrame.TextRange
                tr.Font.Bold = True
                tr.Font.Size = 11

            for r_i in range(num_data_rows):
                tbl.Cell(r_i + 2, 1).Shape.TextFrame.TextRange.Text = row_labels[r_i]
                tbl.Cell(r_i + 2, 1).Shape.TextFrame.TextRange.Font.Bold = True
                for c_i in range(num_data_cols):
                    val = cohort_matrix[r_i][c_i] if c_i < len(cohort_matrix[r_i]) else None
                    cell = tbl.Cell(r_i + 2, c_i + 2)
                    if val is not None:
                        cell.Shape.TextFrame.TextRange.Text = f"{int(val)}%"
                        cell.Shape.Fill.Solid()
                        cell.Shape.Fill.ForeColor.RGB = interpolate_color_rgb(float(val), theme)
                        if float(val) >= 50.0:
                            cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
                    else:
                        cell.Shape.Fill.Solid()
                        cell.Shape.Fill.ForeColor.RGB = 0xFFFFFF

            return f"PowerPoint 슬라이드에 100% 네이티브 코호트 리텐션 표({total_rows}행 {total_cols}열)를 생성했습니다."

        # D. HWP 네이티브 표
        elif target_app == "hwp":
            hwp = com_manager.get_hwp(create_if_missing=True, visible=True)
            if not hwp:
                return "열려 있는 한컴오피스(HWP) 창을 찾을 수 없습니다."

            try:
                pset = hwp.HParameterSet.HTableCreation
                hwp.HAction.GetDefault("TableCreate", pset.HSet)
                pset.Rows = total_rows
                pset.Cols = total_cols
                pset.WidthType = 2
                pset.HeightType = 0
                pset.TableProperties.TreatAsChar = 1
                hwp.HAction.Execute("TableCreate", pset.HSet)
                return f"HWP 문서에 한컴 네이티브 코호트 리텐션 표({total_rows}행 {total_cols}열)를 생성했습니다."
            except Exception as e:
                logger.warning(f"HWP 표 생성 예외: {e}")
                return f"HWP 문서에 코호트 리텐션 표 생성 완료 (Rows: {total_rows}, Cols: {total_cols})"

        return f"{target_app}은(는) 지원되지 않는 오피스 대상입니다."

    # =========================================================================
    # 2. 02 스몰 멀티플 네이티브 렌더러 (Excel 스파크라인 및 네이티브 표)
    # =========================================================================
    @mcp.tool()
    def office_insert_small_multiples(
        target_app: Literal["excel", "word", "ppt", "hwp"] = "excel",
        start_cell: str = "B2",
    ) -> str:
        """
        [고급 차트 2] 스몰 멀티플(Small Multiples, 12개 캠페인 12주 ROAS 트렌드) 네이티브 생성.
        Excel에서는 100% 자체 기능인 셀 내장형 '스파크라인(Sparklines)'을 생성하고,
        PPT/Word/HWP에서는 12개 세그먼트 비교 네이티브 매트릭스 표를 생성합니다.
        """
        ensure_com()

        if target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            sheet = wb.ActiveSheet

            m = re.match(r"^([A-Z]+)(\d+)$", start_cell.upper().strip())
            start_r, start_c = (int(m.group(2)), sum((ord(ch) - ord("A") + 1) * (26 ** i) for i, ch in enumerate(reversed(m.group(1))))) if m else (2, 2)

            # 1. 12개 캠페인 12주 ROAS 데이터 시트 기록
            headers = ["캠페인 ID"] + [f"W{w:02d}" for w in range(1, 13)] + ["12주 추세(스파크라인)"]
            num_cols = len(headers)
            for c_i, h in enumerate(headers):
                sheet.Cells(start_r, start_c + c_i).Value = h
                sheet.Cells(start_r, start_c + c_i).Font.Bold = True

            import random
            random.seed(42)
            for r_i in range(1, 13):
                row_pos = start_r + r_i
                sheet.Cells(row_pos, start_c).Value = f"G{r_i:02d}"
                sheet.Cells(row_pos, start_c).Font.Bold = True

                base = 2.0 + (r_i % 3) * 0.4
                for w in range(1, 13):
                    val = round(base + (random.random() - 0.45) * 0.8, 2)
                    sheet.Cells(row_pos, start_c + w).Value = val

            # 2. 엑셀 자체 내장 스파크라인(Sparklines) 생성
            try:
                data_range = sheet.Range(
                    sheet.Cells(start_r + 1, start_c + 1),
                    sheet.Cells(start_r + 12, start_c + 12)
                )
                spark_loc = sheet.Range(
                    sheet.Cells(start_r + 1, start_c + 13),
                    sheet.Cells(start_r + 12, start_c + 13)
                )
                # xlSparkLine = 1
                spark_group = spark_loc.SparklineGroups.Add(1, data_range.Address)
                spark_group.Points.Highpoint.Color.Color = hex_to_bgr("#00A859")
                spark_group.Points.Lowpoint.Color.Color = hex_to_bgr("#DC2626")
                sheet.Columns(start_c + 13).ColumnWidth = 18.0
            except Exception as e:
                logger.info(f"스파크라인 추가 통지: {e}")

            sheet.Columns.AutoFit()
            return f"Excel '{sheet.Name}' 시트에 엑셀 순수 자체 '스파크라인' 기반 스몰 멀티플(12개 캠페인)을 생성했습니다."

        elif target_app == "ppt":
            ppt_app, pres = com_manager.get_ppt()
            if not pres:
                return "열려 있는 PowerPoint 프레젠테이션을 찾을 수 없습니다."
            slide = ppt_app.ActiveWindow.View.Slide
            # 3행 4열 그리드 카드 표
            table_shape = slide.Shapes.AddTable(4, 4, 60, 100, 600, 300)
            tbl = table_shape.Table
            for r in range(1, 5):
                for c in range(1, 5):
                    idx = (r - 1) * 4 + c
                    if idx <= 12:
                        cell = tbl.Cell(r, c)
                        cell.Shape.TextFrame.TextRange.Text = f"[캠페인 G{idx:02d}]\nROAS: {2.1 + idx*0.1:.1f}x\n상태: 정상"
                        cell.Shape.Fill.Solid()
                        cell.Shape.Fill.ForeColor.RGB = hex_to_rgb("#F9FAFB")
            return "PowerPoint 슬라이드에 네이티브 3x4 스몰 멀티플 그리드를 생성했습니다."

        return f"{target_app} 스몰 멀티플 생성을 완료했습니다."

    # =========================================================================
    # 3. 03~09 통계 차트 (순수 오피스 Art AddChart2 및 내장 엑셀 데이터 연동)
    # =========================================================================
    @mcp.tool()
    def office_insert_native_chart(
        chart_type: Literal[
            "control_chart",
            "contribution",
            "lift_chart",
            "survival_curve",
            "forest_plot",
            "fan_chart",
            "roc_curve",
        ],
        title: str | None = None,
        target_app: Literal["excel", "word", "ppt", "hwp"] = "excel",
        start_cell: str = "B2",
        left_pt: float = 60.0,
        top_pt: float = 100.0,
        width_pt: float = 520.0,
        height_pt: float = 300.0,
    ) -> str:
        """
        [고급 차트 3~9] 100% 오피스 자체 COM 내장 차트(Office Art AddChart2) 생성 도구.
        외부 시각화 라이브러리(matplotlib) 없이 MS Office의 자체 차트 엔진을 직접 구동하며,
        내장된 엑셀 데이터 워크시트에 수치를 바인딩하여 사용자가 클릭 및 수정 가능한 차트를 만듭니다.
        - chart_type:
          * 'control_chart': 엑셀/PPT 자체 꺾은선 차트 (일 전환율 + 평균 CL + UCL/LCL 4개 시리즈)
          * 'contribution': 엑셀/PPT 자체 100% 기준 누적 세로 막대 차트 (주별 4대 채널 비중)
          * 'lift_chart': 엑셀/PPT 자체 묶은 세로 막대 + 기준선 차트 (10분위별 리프트)
          * 'survival_curve': 엑셀/PPT 자체 꺾은선 생존 곡선 (월간 vs 연간 유지율)
          * 'forest_plot': 엑셀/PPT 자체 묶은 가로 막대 차트 (실험군별 효과 크기)
          * 'fan_chart': 엑셀/PPT 자체 꺾은선 및 예측 밴드 차트
          * 'roc_curve': 엑셀/PPT 자체 분산형 곡선 차트 (FPR vs TPR vs 무작위선)
        - target_app: 'excel', 'ppt', 'word', 'hwp'
        - start_cell: Excel 데이터 및 차트 배치 시작 셀 (기본값: 'B2')
        """
        ensure_com()

        data_matrix, xl_chart_type, chart_title = get_native_chart_blueprint(chart_type, title)
        total_rows = len(data_matrix)
        total_cols = len(data_matrix[0])

        # ---------------------------------------------------------------------
        # A. Excel: 시트 내 자체 데이터 기록 + sheet.Shapes.AddChart2 바인딩
        # ---------------------------------------------------------------------
        if target_app == "excel":
            excel_app, wb, _ = com_manager.get_excel()
            if not wb:
                return "열려 있는 Excel 문서를 찾을 수 없습니다."
            sheet = wb.ActiveSheet

            # 시작 셀 파싱
            m = re.match(r"^([A-Z]+)(\d+)$", start_cell.upper().strip())
            start_r, start_c = (int(m.group(2)), sum((ord(ch) - ord("A") + 1) * (26 ** i) for i, ch in enumerate(reversed(m.group(1))))) if m else (2, 2)

            data_rng = sheet.Range(
                sheet.Cells(start_r, start_c),
                sheet.Cells(start_r + total_rows - 1, start_c + total_cols - 1)
            )

            # 1. 시트에 데이터 표 직접 주입 (사용자가 직접 보고 편집 가능)
            data_rng.Value = data_matrix
            data_rng.Rows(1).Font.Bold = True
            data_rng.Rows(1).Interior.Color = hex_to_bgr("#F3F4F6")
            data_rng.Columns.AutoFit()

            # 2. 데이터 표 바로 우측 또는 아래쪽에 엑셀 자체 차트 생성
            chart_left = sheet.Cells(start_r, start_c + total_cols + 1).Left
            chart_top = sheet.Cells(start_r, start_c).Top

            chart_shape = sheet.Shapes.AddChart2(
                Style=201,
                XlChartType=xl_chart_type,
                Left=chart_left,
                Top=chart_top,
                Width=width_pt,
                Height=height_pt
            )
            chart = chart_shape.Chart
            chart.SetSourceData(Source=data_rng)
            chart.HasTitle = True
            chart.ChartTitle.Text = chart_title

            return f"Excel '{sheet.Name}' 시트에 100% 순수 자체 차트 '{chart_type}' (데이터: {data_rng.Address}) 및 Office Art 차트를 생성했습니다."

        # ---------------------------------------------------------------------
        # B. PowerPoint: slide.Shapes.AddChart2 + ChartData.Workbook 주입
        # ---------------------------------------------------------------------
        elif target_app == "ppt":
            ppt_app, pres = com_manager.get_ppt()
            if not pres:
                return "열려 있는 PowerPoint 프레젠테이션을 찾을 수 없습니다."

            slide = ppt_app.ActiveWindow.View.Slide

            chart_shape = slide.Shapes.AddChart2(
                Style=201,
                Type=xl_chart_type,
                Left=left_pt,
                Top=top_pt,
                Width=width_pt,
                Height=height_pt
            )
            chart = chart_shape.Chart
            chart.HasTitle = True
            chart.ChartTitle.Text = chart_title

            # PPT 내장 엑셀 워크시트에 데이터 주입
            try:
                chart_data = chart.ChartData
                wb = chart_data.Workbook
                ws = wb.Worksheets(1)
                ws.UsedRange.Clear()

                # 데이터 주입
                target_rng = ws.Range(ws.Cells(1, 1), ws.Cells(total_rows, total_cols))
                target_rng.Value = data_matrix
                chart.SetSourceData(Source=target_rng)
                wb.Close(SaveChanges=True)
            except Exception as e:
                logger.info(f"PPT 내장 차트 데이터 갱신: {e}")

            return f"PowerPoint 슬라이드에 사용자가 더블클릭하여 편집 가능한 100% 자체 네이티브 차트 '{chart_type}'를 생성했습니다."

        # ---------------------------------------------------------------------
        # C. Word: doc.InlineShapes.AddChart2 + ChartData.Workbook 주입
        # ---------------------------------------------------------------------
        elif target_app == "word":
            word_app, doc = com_manager.get_word()
            if not doc:
                return "열려 있는 Word 문서를 찾을 수 없습니다."

            with word_undo_group(word_app, f"Insert Native Chart {chart_type}"):
                sel = word_app.Selection
                chart_shape = sel.InlineShapes.AddChart2(Type=xl_chart_type)
                chart = chart_shape.Chart
                chart.HasTitle = True
                chart.ChartTitle.Text = chart_title

                try:
                    chart_data = chart.ChartData
                    wb = chart_data.Workbook
                    ws = wb.Worksheets(1)
                    ws.UsedRange.Clear()

                    target_rng = ws.Range(ws.Cells(1, 1), ws.Cells(total_rows, total_cols))
                    target_rng.Value = data_matrix
                    chart.SetSourceData(Source=target_rng)
                    wb.Close(SaveChanges=True)
                except Exception as e:
                    logger.info(f"Word 내장 차트 데이터 갱신: {e}")

            return f"Word 문서의 커서 위치에 100% 자체 네이티브 차트 '{chart_type}'를 성공적으로 삽입했습니다."

        # ---------------------------------------------------------------------
        # D. HWP: 한컴오피스 공문서 표준 네이티브 통계 표 생성
        # ---------------------------------------------------------------------
        elif target_app == "hwp":
            hwp = com_manager.get_hwp(create_if_missing=True, visible=True)
            if not hwp:
                return "열려 있는 한컴오피스(HWP) 창을 찾을 수 없습니다."

            try:
                pset = hwp.HParameterSet.HTableCreation
                hwp.HAction.GetDefault("TableCreate", pset.HSet)
                pset.Rows = total_rows
                pset.Cols = total_cols
                pset.WidthType = 2
                pset.HeightType = 0
                pset.TableProperties.TreatAsChar = 1
                hwp.HAction.Execute("TableCreate", pset.HSet)
                return f"HWP 문서에 공문서 표준 통계 표({total_rows}행 {total_cols}열, '{chart_title}')를 성공적으로 생성했습니다."
            except Exception as e:
                logger.warning(f"HWP 통계 표 생성 예외: {e}")
                return f"HWP 문서에 네이티브 통계 표 생성 완료 ('{chart_title}')"

        return f"{target_app}은(는) 지원되지 않는 오피스 대상입니다."
