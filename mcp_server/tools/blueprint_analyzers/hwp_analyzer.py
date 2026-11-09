# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_analyzers/hwp_analyzer.py
한컴오피스 (HWP) 용지 여백, 표 컨트롤, 문단/글자 모양 역공학 분석기.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com
from .base_analyzer import BaseDocumentAnalyzer

logger = logging.getLogger("mcp_server.hwp_analyzer")


class HwpDocumentAnalyzer(BaseDocumentAnalyzer):
    doc_type: str = "hwp"

    def analyze_file(self, file_path: Path) -> dict[str, Any]:
        """HWP 파일을 열어 페이지 여백 및 표, 문단 추출 (DRM 친화적 COM 개방)"""
        ensure_com()
        hwp_app = com_manager.get_or_create_hwp_app(visible=True)
        try:
            hwp_app.SetMessageBoxMode(0x00020000)  # 보안 승인 및 경고 대화상자 자동 수락
        except Exception:
            pass

        abs_path = str(file_path.resolve())

        try:
            hwp_app.Open(abs_path, "HWP", "versionwarning:false;lock:false")
            return self._extract_from_hwp(hwp_app, file_path.stem)
        finally:
            try:
                hwp_app.Clear(1)  # 1 = 저장 여부 묻지 않고 닫기
            except Exception:
                pass

    def analyze_active_window(self) -> dict[str, Any]:
        """현재 모니터에 열려 있는 활성 한컴오피스(HWP) 창 분석"""
        ensure_com()
        hwp_app = com_manager.get_hwp()
        if not hwp_app:
            raise RuntimeError("현재 활성화된 한컴오피스(HWP) 창을 찾을 수 없습니다.")

        name = "ActiveHWP"
        try:
            p = hwp_app.Path
            if p:
                name = Path(p).stem
        except Exception:
            pass

        return self._extract_from_hwp(hwp_app, name)

    def _extract_from_hwp(self, hwp: Any, source_name: str) -> dict[str, Any]:
        page_setup = {
            "top_margin_mm": 20.0,
            "bottom_margin_mm": 15.0,
            "left_margin_mm": 20.0,
            "right_margin_mm": 20.0,
            "header_mm": 10.0,
            "footer_mm": 10.0,
        }

        # PageDef 시도
        try:
            act = hwp.CreateAction("PageSetup")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pdef = pset.Item("PageDef")
            if pdef:
                # 1 mm = 약 283.465 hwpunit
                page_setup["top_margin_mm"] = round(pdef.Item("TopMargin") / 283.465, 1)
                page_setup["bottom_margin_mm"] = round(pdef.Item("BottomMargin") / 283.465, 1)
                page_setup["left_margin_mm"] = round(pdef.Item("LeftMargin") / 283.465, 1)
                page_setup["right_margin_mm"] = round(pdef.Item("RightMargin") / 283.465, 1)
        except Exception as e:
            logger.debug(f"HWP PageDef 추출 실패, 기본값 사용: {e}")

        elements: list[dict[str, Any]] = []

        # 1. 텍스트 스캔
        try:
            hwp.InitScan()
            line_idx = 1
            while line_idx <= 80:
                res, text = hwp.GetText()
                if res in (0, 1):  # 0: 텍스트 없음/끝, 1: 텍스트 있음
                    if not text:
                        if res == 0:
                            break
                        continue
                    text_str = text.strip()
                    if text_str:
                        is_heading = any(text_str.startswith(p) for p in ["□", "○", "※", "1.", "I.", "가."])
                        elements.append({
                            "element_id": f"hwp_line_{line_idx}",
                            "type": "paragraph",
                            "text": text_str,
                            "is_static": is_heading,
                            "slot_id": f"slot_hwp_heading_{line_idx}" if is_heading else f"slot_hwp_text_{line_idx}",
                            "label": f"[{'항목제목' if is_heading else '본문'}] {text_str[:20]}",
                        })
                        line_idx += 1
                else:
                    break
            hwp.ReleaseScan()
        except Exception as e:
            logger.debug(f"HWP 텍스트 스캔 오류: {e}")

        # 2. 표 컨트롤 탐색
        tbl_idx = 1
        try:
            ctrl = hwp.HeadCtrl
            while ctrl:
                if getattr(ctrl, "CtrlID", "") == "tbl":
                    elements.append({
                        "element_id": f"hwp_table_{tbl_idx}",
                        "type": "table",
                        "is_static": False,
                        "slot_id": f"slot_hwp_table_{tbl_idx}",
                        "label": f"HWP 표준 규격 표 {tbl_idx}",
                    })
                    tbl_idx += 1
                ctrl = ctrl.Next
        except Exception as e:
            logger.debug(f"HWP 표 컨트롤 탐색 오류: {e}")

        return {
            "doc_type": "hwp",
            "source_name": source_name,
            "canvas": page_setup,
            "design_tokens": {
                "font_family": "함초롬바탕",
            },
            "pages": [{
                "page_index": 1,
                "title": "공문서 본문",
                "elements": elements,
            }],
        }

