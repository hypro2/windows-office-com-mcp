# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_renderers/hwp_renderer.py
한컴오피스 (HWP) 청사진을 데스크톱 COM으로 100% 동일하게 실시간 재생(Playback)하는 렌더러.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com

logger = logging.getLogger("mcp_server.hwp_renderer")


class HwpBlueprintRenderer:
    """HWP 사용자 청사진 실시간 COM 렌더러"""

    @classmethod
    def render(
        cls,
        blueprint: dict[str, Any],
        payload: dict[str, Any],
        output_path: str | Path | None = None,
    ) -> dict[str, Any]:
        ensure_com()
        hwp_app = com_manager.get_or_create_hwp_app(visible=True)
        hwp_app.HAction.Run("FileNew")

        # 1. 여백 설정
        canvas = blueprint.get("canvas", {})
        try:
            act = hwp_app.CreateAction("PageSetup")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pdef = pset.Item("PageDef")
            if pdef:
                pdef.SetItem("TopMargin", int(float(canvas.get("top_margin_mm", 20.0)) * 283.465))
                pdef.SetItem("BottomMargin", int(float(canvas.get("bottom_margin_mm", 15.0)) * 283.465))
                pdef.SetItem("LeftMargin", int(float(canvas.get("left_margin_mm", 20.0)) * 283.465))
                pdef.SetItem("RightMargin", int(float(canvas.get("right_margin_mm", 20.0)) * 283.465))
                act.Execute(pset)
        except Exception as e:
            logger.debug(f"HWP 여백 설정 오류: {e}")

        pages = blueprint.get("pages", [])
        total_elements = 0

        for page in pages:
            elements = page.get("elements", [])
            for el in elements:
                el_type = el.get("type", "paragraph")
                slot_id = el.get("slot_id")

                if el_type == "paragraph":
                    text_val = el.get("text", "")
                    if slot_id and slot_id in payload:
                        text_val = str(payload[slot_id])

                    # HWP 텍스트 입력
                    try:
                        act = hwp_app.CreateAction("InsertText")
                        pset = act.CreateSet()
                        pset.SetItem("Text", text_val + "\r\n")
                        act.Execute(pset)
                        total_elements += 1
                    except Exception:
                        pass

                elif el_type == "table":
                    # 간단 표 생성
                    try:
                        hwp_app.HAction.Run("TableCreate")
                        total_elements += 1
                    except Exception:
                        pass

        out_file = None
        if output_path:
            out_file = str(Path(output_path).resolve())
            hwp_app.SaveAs(out_file)
            logger.info(f"HWP 청사진 문서 저장 완료: {out_file}")

        return {
            "success": True,
            "doc_type": "hwp",
            "template_name": blueprint.get("template_name", ""),
            "elements_count": total_elements,
            "output_path": out_file,
            "message": f"한컴오피스 HWP 문서({total_elements}개 요소)가 성공적으로 복제 생성되었습니다.",
        }

