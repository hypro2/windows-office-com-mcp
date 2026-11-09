"""
SVG to PowerPoint Native Object Compiler (Inspired by Hugo He ppt-master)

Compiles an SVG Intermediate Representation (IR) containing <rect>, <circle>,
<path>, and <text> into native PowerPoint DrawingML / Slide Shapes with EMU precision.
"""

import re
from typing import Dict, Any, List, Optional
from xml.etree import ElementTree
from mcp_server.response_envelope import ok_response, error_response


def pt_to_emu(pt_val: float) -> int:
    """Converts points to English Metric Units (1 pt = 12,700 EMU)."""
    return int(round(pt_val * 12700))


def px_to_emu(px_val: float) -> int:
    """Converts 96 DPI pixels to EMU (1 px = 9,525 EMU)."""
    return int(round(px_val * 9525))


class SvgPptCompiler:
    """Parses SVG layout blocks and maps them to PowerPoint native shapes with EMU precision."""

    @staticmethod
    def _parse_attr_float(elem: ElementTree.Element, attr_name: str, default: float = 0.0) -> float:
        val = elem.attrib.get(attr_name, "")
        val = re.sub(r"[^\d.-]", "", val)
        try:
            return float(val) if val else default
        except ValueError:
            return default

    @classmethod
    def compile_svg_to_shape_ops(cls, svg_xml: str) -> Dict[str, Any]:
        """
        Parses SVG XML and emits a list of native PPT shape creation operations with EMU values.
        """
        try:
            clean_xml = re.sub(r'\sxmlns="[^"]+"', '', svg_xml)
            root = ElementTree.fromstring(clean_xml)
        except ElementTree.ParseError as e:
            return error_response(
                command="compile_svg_to_shapes",
                summary=f"SVG XML 구문 분석 실패: {e}",
                suggestion="올바른 형식의 SVG 태그('<svg>...</svg>')를 제공하세요.",
                code="ERR_SVG_PARSE_FAILED"
            )

        width_pt = cls._parse_attr_float(root, "width", 960.0)
        height_pt = cls._parse_attr_float(root, "height", 540.0)

        shape_ops = []

        for elem in root.iter():
            tag = elem.tag.rsplit("}", 1)[-1].lower()

            if tag == "rect":
                x = cls._parse_attr_float(elem, "x", 0.0)
                y = cls._parse_attr_float(elem, "y", 0.0)
                w = cls._parse_attr_float(elem, "width", 100.0)
                h = cls._parse_attr_float(elem, "height", 50.0)
                fill = elem.attrib.get("fill", "#E0E0E0")
                rx = cls._parse_attr_float(elem, "rx", 0.0)

                shape_ops.append({
                    "shape_type": "round_rect" if rx > 0 else "rectangle",
                    "left_pt": x,
                    "top_pt": y,
                    "width_pt": w,
                    "height_pt": h,
                    "left_emu": pt_to_emu(x),
                    "top_emu": pt_to_emu(y),
                    "width_emu": pt_to_emu(w),
                    "height_emu": pt_to_emu(h),
                    "fill_color": fill,
                    "corner_radius": rx
                })

            elif tag == "text":
                x = cls._parse_attr_float(elem, "x", 0.0)
                y = cls._parse_attr_float(elem, "y", 0.0)
                font_size = cls._parse_attr_float(elem, "font-size", 14.0)
                fill = elem.attrib.get("fill", "#000000")
                text_content = (elem.text or "").strip()

                if text_content:
                    top_adjusted = max(0.0, y - font_size)
                    w = max(100.0, len(text_content) * font_size * 0.7)
                    h = font_size * 1.5
                    shape_ops.append({
                        "shape_type": "text_box",
                        "left_pt": x,
                        "top_pt": top_adjusted,
                        "width_pt": w,
                        "height_pt": h,
                        "left_emu": pt_to_emu(x),
                        "top_emu": pt_to_emu(top_adjusted),
                        "width_emu": pt_to_emu(w),
                        "height_emu": pt_to_emu(h),
                        "text": text_content,
                        "font_size": font_size,
                        "font_color": fill
                    })

            elif tag == "circle":
                cx = cls._parse_attr_float(elem, "cx", 50.0)
                cy = cls._parse_attr_float(elem, "cy", 50.0)
                r = cls._parse_attr_float(elem, "r", 25.0)
                fill = elem.attrib.get("fill", "#1E88E5")
                x, y = cx - r, cy - r
                w, h = r * 2, r * 2

                shape_ops.append({
                    "shape_type": "oval",
                    "left_pt": x,
                    "top_pt": y,
                    "width_pt": w,
                    "height_pt": h,
                    "left_emu": pt_to_emu(x),
                    "top_emu": pt_to_emu(y),
                    "width_emu": pt_to_emu(w),
                    "height_emu": pt_to_emu(h),
                    "fill_color": fill
                })

            elif tag == "path":
                d = elem.attrib.get("d", "").strip()
                fill = elem.attrib.get("fill", "none")
                stroke = elem.attrib.get("stroke", "#000000")
                if d:
                    shape_ops.append({
                        "shape_type": "custom_path",
                        "path_data": d,
                        "fill_color": fill,
                        "stroke_color": stroke
                    })

        return ok_response(
            command="compile_svg_to_shapes",
            summary=f"SVG로부터 {len(shape_ops)}개의 파워포인트 DrawingML 네이티브 작업이 컴파일되었습니다.",
            detail={
                "canvas": {"width_pt": width_pt, "height_pt": height_pt, "width_emu": pt_to_emu(width_pt)},
                "shape_count": len(shape_ops),
                "operations": shape_ops
            }
        )
