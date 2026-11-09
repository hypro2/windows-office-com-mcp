# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_analyzers/ppt_analyzer.py
PowerPoint (PPTX) 슬라이드 도형, 텍스트 슬롯, 서식 역공학 분석기.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ...com_manager import com_manager, ensure_com
from ..ppt_recipe_engine import PptRecipeEngine
from .base_analyzer import BaseDocumentAnalyzer

logger = logging.getLogger("mcp_server.ppt_analyzer")


class PptDocumentAnalyzer(BaseDocumentAnalyzer):
    doc_type: str = "ppt"

    def analyze_file(self, file_path: Path) -> dict[str, Any]:
        """PPTX 파일을 열어 슬라이드 전체의 레이아웃과 요소를 추출 (DRM 친화적 COM 개방)"""
        ensure_com()
        # 사내 DRM 복호화 통과를 위해 가시 모드로 프로세스 바인딩
        ppt_app = com_manager.get_or_create_ppt_app(visible=True)
        abs_path = str(file_path.resolve())

        prs = None
        try:
            # WithWindow=True 로 열어야 사내 DRM 필터 모듈이 HWND를 정상 감지하고 메모리 복호화 수행
            prs = ppt_app.Presentations.Open(abs_path, ReadOnly=True, Untitled=False, WithWindow=True)
            return self._extract_from_presentation(prs, file_path.stem)
        finally:
            if prs:
                try:
                    prs.Close()
                except Exception:
                    pass

    def analyze_active_window(self) -> dict[str, Any]:
        """현재 모니터에 열려 있는 활성 PowerPoint 창의 프레젠테이션 분석"""
        ensure_com()
        ppt_app, prs = com_manager.get_ppt(create_if_missing=False)
        if not ppt_app or not prs:
            raise RuntimeError("현재 활성화된 PowerPoint 프레젠테이션 창을 찾을 수 없습니다.")

        name = getattr(prs, "Name", "ActivePresentation")
        return self._extract_from_presentation(prs, Path(name).stem)

    def _extract_from_presentation(self, prs: Any, source_name: str) -> dict[str, Any]:
        w = float(prs.PageSetup.SlideWidth)
        h = float(prs.PageSetup.SlideHeight)

        pages: list[dict[str, Any]] = []
        design_colors: list[int] = []
        font_names: set[str] = set()

        for s_idx in range(1, prs.Slides.Count + 1):
            s = prs.Slides(s_idx)
            recipe = PptRecipeEngine.extract_slide_recipe(
                slide=s,
                slide_w=w,
                slide_h=h,
                recipe_id=f"{source_name}_s{s_idx}",
                name=f"슬라이드 {s_idx}",
            )

            elements = []
            for sh_idx, sh in enumerate(recipe.get("shapes", [])):
                t_content = ""
                if "text" in sh:
                    t_content = sh["text"].get("content", "").strip()
                    font_names.add(sh["text"].get("font_name", "맑은 고딕"))

                fill_rgb = sh.get("fill", {}).get("rgb")
                if fill_rgb:
                    design_colors.append(fill_rgb)

                elements.append({
                    "element_id": f"s{s_idx}_sh{sh_idx + 1}",
                    "name": sh.get("name", f"Shape_{sh_idx + 1}"),
                    "type": "shape",
                    "shape_spec": sh,
                    "text": t_content,
                    "is_static": bool(not t_content),
                    "slot_id": f"slide_{s_idx}_slot_{sh_idx + 1}" if t_content else None,
                    "label": f"슬라이드 {s_idx} - {t_content[:15]}" if t_content else "",
                })

            pages.append({
                "page_index": s_idx,
                "title": f"{s_idx}번째 슬라이드",
                "background": recipe.get("background", {}),
                "elements": elements,
            })

        return {
            "doc_type": "ppt",
            "source_name": source_name,
            "canvas": {"width": w, "height": h},
            "design_tokens": {
                "fonts": list(font_names)[:5],
                "sample_colors_rgb": design_colors[:8],
            },
            "pages": pages,
        }

