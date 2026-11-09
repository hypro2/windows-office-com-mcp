# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_renderers/ppt_renderer.py
PowerPoint 청사진을 데스크톱 COM으로 100% 동일하게 실시간 재생(Playback)하는 렌더러.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com
from ..ppt_recipe_engine import PptRecipeEngine
from ..ppt_tools import ppt_freeze_display

logger = logging.getLogger("mcp_server.ppt_renderer")


class PptBlueprintRenderer:
    """PPT 사용자 청사진 실시간 COM 렌더러"""

    @classmethod
    def render(
        cls,
        blueprint: dict[str, Any],
        payload: dict[str, Any],
        output_path: str | Path | None = None,
    ) -> dict[str, Any]:
        ensure_com()
        ppt_app = com_manager.get_or_create_ppt_app(visible=True)
        prs = ppt_app.Presentations.Add(WithWindow=True)

        canvas = blueprint.get("canvas", {})
        if "width" in canvas and "height" in canvas:
            try:
                prs.PageSetup.SlideWidth = float(canvas["width"])
                prs.PageSetup.SlideHeight = float(canvas["height"])
            except Exception:
                pass

        pages = blueprint.get("pages", [])
        total_created_shapes = 0

        with ppt_freeze_display(ppt_app):
            for p_idx, page in enumerate(pages, start=1):
                slide = prs.Slides.Add(p_idx, 12)  # ppLayoutBlank = 12

                # 1. 배경색
                bg = page.get("background", {})
                if bg.get("visible") and "rgb" in bg:
                    try:
                        slide.Background.Fill.Solid()
                        slide.Background.Fill.ForeColor.RGB = int(bg["rgb"])
                    except Exception:
                        pass

                # 2. 요소 렌더링
                elements = page.get("elements", [])
                shapes_to_render = []

                for el in elements:
                    shape_spec = dict(el.get("shape_spec", {}))
                    slot_id = el.get("slot_id")

                    # 페이로드에서 새 데이터 치환
                    if slot_id and slot_id in payload:
                        new_val = str(payload[slot_id])
                        if "text" in shape_spec:
                            shape_spec["text"]["content"] = new_val
                        else:
                            shape_spec["text"] = {"content": new_val}
                    elif el.get("text") and "text" in shape_spec:
                        # 기본 텍스트 유지
                        shape_spec["text"]["content"] = el["text"]

                    shapes_to_render.append(shape_spec)

                recipe = {
                    "slide_width": canvas.get("width", 960.0),
                    "slide_height": canvas.get("height", 540.0),
                    "background": bg,
                    "shapes": shapes_to_render,
                }

                created_ids = PptRecipeEngine.render_recipe(slide, recipe)
                total_created_shapes += len(created_ids)

        out_file = None
        if output_path:
            out_file = str(Path(output_path).resolve())
            prs.SaveAs(out_file)
            logger.info(f"PowerPoint 청사진 문서 저장 완료: {out_file}")

        try:
            ppt_app.Visible = True
            ppt_app.ActiveWindow.Activate()
        except Exception:
            pass

        return {
            "success": True,
            "doc_type": "ppt",
            "template_name": blueprint.get("template_name", ""),
            "slides_count": len(pages),
            "shapes_count": total_created_shapes,
            "output_path": out_file,
            "message": f"PowerPoint 프레젠테이션({len(pages)}개 슬라이드, {total_created_shapes}개 도형)이 성공적으로 복제 생성되었습니다.",
        }

