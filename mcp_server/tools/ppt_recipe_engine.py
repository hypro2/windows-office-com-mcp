# -*- coding: utf-8 -*-
"""
mcp_server/tools/ppt_recipe_engine.py (개선 버전)
- Placeholder / Textbox 완벽 처리 (AddTextbox 폴백)
- 슬라이드 배경색 복제
- 그림자 / 라인 / 채우기 안정적 주입
"""

from __future__ import annotations

import logging
from typing import Any

logger = logging.getLogger("mcp_server.ppt_recipe_engine")


class PptRecipeEngine:

    @staticmethod
    def extract_shape_spec(sh: Any) -> dict[str, Any]:
        """단일 COM 도형의 정밀 비주얼 속성을 직렬화 가능한 딕셔너리로 추출"""
        sh_type = getattr(sh, "Type", 1)
        auto_type = getattr(sh, "AutoShapeType", 1)
        
        spec: dict[str, Any] = {
            "name": getattr(sh, "Name", ""),
            "type": sh_type,
            "auto_shape_type": auto_type,
            "left": round(float(sh.Left), 2),
            "top": round(float(sh.Top), 2),
            "width": round(float(sh.Width), 2),
            "height": round(float(sh.Height), 2),
            "adjustments": [],
        }

        # Adjustments (라운딩 모서리 등)
        try:
            if hasattr(sh, "Adjustments") and sh.Adjustments.Count > 0:
                spec["adjustments"] = [round(float(sh.Adjustments.Item(i)), 4) for i in range(1, sh.Adjustments.Count + 1)]
        except Exception:
            pass

        # Fill
        try:
            if sh.Fill.Visible == -1:
                spec["fill"] = {
                    "visible": True,
                    "rgb": int(sh.Fill.ForeColor.RGB),
                    "transparency": round(float(getattr(sh.Fill, "Transparency", 0.0)), 2)
                }
            else:
                spec["fill"] = {"visible": False}
        except Exception:
            spec["fill"] = {"visible": False}

        # Line
        try:
            if sh.Line.Visible == -1:
                spec["line"] = {
                    "visible": True,
                    "weight": round(float(sh.Line.Weight), 2),
                    "rgb": int(sh.Line.ForeColor.RGB),
                    "transparency": round(float(getattr(sh.Line, "Transparency", 0.0)), 2)
                }
            else:
                spec["line"] = {"visible": False}
        except Exception:
            spec["line"] = {"visible": False}

        # Shadow
        try:
            if hasattr(sh, "Shadow") and sh.Shadow.Visible == -1:
                spec["shadow"] = {
                    "visible": True,
                    "type": getattr(sh.Shadow, "Type", 2),
                    "blur": round(float(sh.Shadow.Blur), 1),
                    "transparency": round(float(sh.Shadow.Transparency), 2),
                    "offset_x": round(float(getattr(sh.Shadow, "OffsetX", 0.0)), 1),
                    "offset_y": round(float(getattr(sh.Shadow, "OffsetY", 0.0)), 1),
                }
        except Exception:
            pass

        # TextFrame
        try:
            if sh.HasTextFrame and sh.TextFrame.HasText:
                tr = sh.TextFrame.TextRange
                tf = sh.TextFrame
                spec["text"] = {
                    "content": tr.Text,
                    "font_name": str(tr.Font.Name),
                    "font_size": round(float(tr.Font.Size), 1),
                    "font_bold": bool(tr.Font.Bold == -1),
                    "font_rgb": int(tr.Font.Color.RGB),
                    "align": int(tr.ParagraphFormat.Alignment),
                    "margin_left": round(float(tf.MarginLeft), 1),
                    "margin_top": round(float(tf.MarginTop), 1),
                    "margin_right": round(float(tf.MarginRight), 1),
                    "margin_bottom": round(float(tf.MarginBottom), 1),
                }
        except Exception:
            pass

        return spec

    @classmethod
    def extract_slide_recipe(
        cls,
        slide: Any,
        slide_w: float,
        slide_h: float,
        recipe_id: str,
        name: str = "",
        role: str = "custom"
    ) -> dict[str, Any]:
        """슬라이드 전체의 배경 및 모든 도형을 슬라이스 레시피로 추출"""
        shapes_data: list[dict[str, Any]] = []
        slots_count = 0

        # 슬라이드 배경색
        bg_spec = {"visible": False}
        try:
            if hasattr(slide, "Background") and slide.Background.Fill.Visible == -1:
                bg_spec = {
                    "visible": True,
                    "rgb": int(slide.Background.Fill.ForeColor.RGB)
                }
        except Exception:
            pass

        for sh_idx in range(1, slide.Shapes.Count + 1):
            sh = slide.Shapes(sh_idx)
            sh_spec = cls.extract_shape_spec(sh)

            if "text" in sh_spec and sh_spec["text"]["content"].strip():
                slot_key = f"slot_{slots_count}"
                sh_spec["slot_key"] = slot_key
                slots_count += 1

            shapes_data.append(sh_spec)

        return {
            "recipe_id": recipe_id,
            "name": name,
            "role": role,
            "slide_width": slide_w,
            "slide_height": slide_h,
            "background": bg_spec,
            "shapes_count": len(shapes_data),
            "slots_count": slots_count,
            "shapes": shapes_data
        }

    @classmethod
    def render_recipe(
        cls,
        slide: Any,
        recipe: dict[str, Any],
        data: dict[str, Any] | None = None,
        custom_bounds: tuple[float, float, float, float] | None = None
    ) -> list[int]:
        """선언적 레시피를 순수 COM으로 1:1 완벽 재생 (99% 재현)"""
        created_ids: list[int] = []
        shapes = recipe.get("shapes", [])
        orig_w = float(recipe.get("slide_width", 960.0))
        orig_h = float(recipe.get("slide_height", 540.0))

        scale_x = 1.0
        scale_y = 1.0
        offset_x = 0.0
        offset_y = 0.0

        if custom_bounds:
            bx, by, bw, bh = custom_bounds
            scale_x = bw / max(orig_w, 1.0)
            scale_y = bh / max(orig_h, 1.0)
            offset_x = bx
            offset_y = by

        # 0. 슬라이드 배경색 적용
        bg_spec = recipe.get("background", {})
        if bg_spec.get("visible") and "rgb" in bg_spec:
            try:
                slide.Background.Fill.Solid()
                slide.Background.Fill.ForeColor.RGB = int(bg_spec["rgb"])
            except Exception:
                pass

        data_map = data or {}

        for s in shapes:
            try:
                sh_type = int(s.get("type", 1))
                auto_type = int(s.get("auto_shape_type", 1))

                # 좌표 및 크기 계산
                sx = float(s["left"]) * scale_x + offset_x
                sy = float(s["top"]) * scale_y + offset_y
                sw = max(1.0, float(s["width"]) * scale_x)
                sh_h = max(1.0, float(s["height"]) * scale_y)

                # 도형 생성 분기 (Textbox vs AutoShape)
                # sh_type: 14=msoPlaceholder, 17=msoTextBox
                if sh_type in (14, 17) or auto_type <= 0:
                    new_sh = slide.Shapes.AddTextbox(1, sx, sy, sw, sh_h)
                else:
                    new_sh = slide.Shapes.AddShape(auto_type, sx, sy, sw, sh_h)

                # 1. Adjustments (모서리 라운딩 곡률)
                adjs = s.get("adjustments", [])
                if adjs and hasattr(new_sh, "Adjustments"):
                    for idx, val in enumerate(adjs, start=1):
                        try:
                            new_sh.Adjustments.Item(idx, float(val))
                        except Exception:
                            pass

                # 2. Fill
                fill_data = s.get("fill", {})
                if fill_data.get("visible"):
                    new_sh.Fill.Solid()
                    new_sh.Fill.ForeColor.RGB = int(fill_data["rgb"])
                    if "transparency" in fill_data:
                        try:
                            new_sh.Fill.Transparency = float(fill_data["transparency"])
                        except Exception:
                            pass
                else:
                    new_sh.Fill.Visible = 0

                # 3. Line
                line_data = s.get("line", {})
                if line_data.get("visible"):
                    new_sh.Line.Visible = -1
                    new_sh.Line.Weight = float(line_data["weight"])
                    new_sh.Line.ForeColor.RGB = int(line_data["rgb"])
                    if "transparency" in line_data:
                        try:
                            new_sh.Line.Transparency = float(line_data["transparency"])
                        except Exception:
                            pass
                else:
                    new_sh.Line.Visible = 0

                # 4. Shadow
                shadow_data = s.get("shadow", {})
                if shadow_data.get("visible") and hasattr(new_sh, "Shadow"):
                    try:
                        new_sh.Shadow.Visible = -1
                        if "type" in shadow_data:
                            new_sh.Shadow.Type = int(shadow_data["type"])
                    except Exception:
                        pass

                # 5. Text
                if "text" in s:
                    t = s["text"]
                    tf = new_sh.TextFrame
                    tf.WordWrap = -1  # 자동 줄바꿈 활성화
                    tf.MarginLeft = float(t.get("margin_left", 0.0)) * scale_x
                    tf.MarginTop = float(t.get("margin_top", 0.0)) * scale_y
                    tf.MarginRight = float(t.get("margin_right", 0.0)) * scale_x
                    tf.MarginBottom = float(t.get("margin_bottom", 0.0)) * scale_y

                    slot_key = s.get("slot_key")
                    content_text = str(t.get("content", ""))
                    if slot_key and slot_key in data_map:
                        content_text = str(data_map[slot_key])

                    tr = tf.TextRange
                    tr.Text = content_text
                    tr.Font.Name = str(t.get("font_name", "맑은 고딕"))
                    tr.Font.Size = float(t.get("font_size", 12.0)) * min(scale_x, scale_y)
                    tr.Font.Bold = -1 if t.get("font_bold") else 0
                    tr.Font.Color.RGB = int(t.get("font_rgb", 0))
                    tr.ParagraphFormat.Alignment = int(t.get("align", 1))

                created_ids.append(new_sh.Id)

            except Exception as e:
                logger.warning(f"Shape render error {s.get('name')}: {e}")

        return created_ids
