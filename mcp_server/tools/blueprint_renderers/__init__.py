# -*- coding: utf-8 -*-
from .excel_renderer import ExcelBlueprintRenderer
from .hwp_renderer import HwpBlueprintRenderer
from .ppt_renderer import PptBlueprintRenderer
from .word_renderer import WordBlueprintRenderer

RENDERERS = {
    "ppt": PptBlueprintRenderer,
    "excel": ExcelBlueprintRenderer,
    "word": WordBlueprintRenderer,
    "hwp": HwpBlueprintRenderer,
}


def get_renderer_for_type(doc_type: str):
    """문서 종류에 맞는 전용 COM 렌더러 반환"""
    dt = doc_type.lower()
    if dt not in RENDERERS:
        raise ValueError(f"지원하지 않는 문서 종류입니다: {doc_type} (지원: ppt, excel, word, hwp)")
    return RENDERERS[dt]

