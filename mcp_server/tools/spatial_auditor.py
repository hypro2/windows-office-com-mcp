"""
Spatial Layout & Geometry Auditor (Inspired by GenOffice slides audit & ppt-master font metrics)

Performs mathematical boundary, text overflow, and overlap audits on presentation slides,
producing machine-readable error codes and actionable `suggestion` transforms for AI agents.
"""

from typing import Dict, Any, List, Optional
from dataclasses import dataclass


@dataclass
class BoundingBox:
    id: str
    left: float
    top: float
    width: float
    height: float
    text: str = ""
    font_size: float = 14.0
    aspect_ratio: Optional[float] = None


class SpatialAuditor:
    """Audits slide layouts for boundary violations, overlaps, and text overflow."""

    CANVAS_WIDTH = 960.0   # 16:9 widescreen standard (pt)
    CANVAS_HEIGHT = 540.0

    @classmethod
    def calculate_text_box_height(cls, text: str, font_size: float, box_width: float) -> float:
        """
        Calculates exact text height using CJK (1.0em) and Latin (0.55em) advance width metrics
        derived from ppt-master's font_advances.json.
        """
        if not text:
            return 0.0

        lines = 1
        current_line_w = 0.0
        line_height = font_size * 1.35  # standard line spacing

        for char in text:
            if char == '\n':
                lines += 1
                current_line_w = 0.0
                continue

            # CJK Unicode range: Hangul syllables, Hanja, Katakana/Hiragana -> 1.0em
            char_w = font_size if ord(char) >= 0x1100 else font_size * 0.55
            if current_line_w + char_w > box_width:
                lines += 1
                current_line_w = char_w
            else:
                current_line_w += char_w

        return lines * line_height

    @classmethod
    def audit_slide_elements(
        cls,
        elements: List[BoundingBox],
        slide_index: int = 1
    ) -> Dict[str, Any]:
        """
        Audits a collection of slide bounding boxes against spatial rules.
        """
        issues = []

        # 1. Canvas Boundary Check (out_of_bounds & off_slide)
        for el in elements:
            if el.left < 0 or el.top < 0:
                issues.append({
                    "code": "off_slide",
                    "element_id": el.id,
                    "level": "error",
                    "message": f"요소 '{el.id}'가 슬라이드 상단/좌측 경계를 벗어났습니다.",
                    "suggestion": {"op": "setTransform", "left": max(20.0, el.left), "top": max(20.0, el.top)}
                })
            if el.left + el.width > cls.CANVAS_WIDTH or el.top + el.height > cls.CANVAS_HEIGHT:
                safe_width = min(el.width, cls.CANVAS_WIDTH - el.left - 20.0)
                issues.append({
                    "code": "out_of_bounds",
                    "element_id": el.id,
                    "level": "error",
                    "message": f"요소 '{el.id}'가 슬라이드 우측/하단 경계를 초과했습니다.",
                    "suggestion": {"op": "setTransform", "width": max(50.0, safe_width)}
                })

        # 2. Text Overflow Check (Using exact advance width calculations)
        for el in elements:
            if el.text:
                required_height = cls.calculate_text_box_height(el.text, el.font_size, el.width)
                if required_height > el.height * 1.10:  # 10% tolerance buffer
                    issues.append({
                        "code": "text_overflow",
                        "element_id": el.id,
                        "level": "warning",
                        "message": f"요소 '{el.id}'의 텍스트가 박스 높이({el.height:.1f}pt)를 초과합니다 (소요 높이: {required_height:.1f}pt).",
                        "suggestion": {"op": "setTransform", "height": round(required_height + 15.0, 1)}
                    })

        # 3. Overlap Check (Intersection over Minimum Area >= 30%)
        for i in range(len(elements)):
            for j in range(i + 1, len(elements)):
                b1, b2 = elements[i], elements[j]
                intersect_x = max(0.0, min(b1.left + b1.width, b2.left + b2.width) - max(b1.left, b2.left))
                intersect_y = max(0.0, min(b1.top + b1.height, b2.top + b2.height) - max(b1.top, b2.top))
                overlap_area = intersect_x * intersect_y

                min_area = min(b1.width * b1.height, b2.width * b2.height)
                if min_area > 0 and (overlap_area / min_area) > 0.3:
                    issues.append({
                        "code": "overlap",
                        "element_ids": [b1.id, b2.id],
                        "level": "warning",
                        "message": f"요소 '{b1.id}'와 '{b2.id}'가 30% 이상 겹쳐 가독성을 해칩니다.",
                        "suggestion": {"op": "reposition", "target": b2.id, "suggested_top": b1.top + b1.height + 15.0}
                    })

        status = "ok" if not issues else "partial" if all(i["level"] == "warning" for i in issues) else "error"
        summary = "슬라이드 공간 기하학 검사 통과 (결함 없음)" if not issues else f"{len(issues)}개의 공간 레이아웃 결함이 발견되었습니다."

        return {
            "status": status,
            "command": "spatial_layout_audit",
            "slide_index": slide_index,
            "summary": summary,
            "issues": issues,
            "passed": len(issues) == 0
        }
