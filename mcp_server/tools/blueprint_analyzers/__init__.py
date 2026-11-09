# -*- coding: utf-8 -*-
from .base_analyzer import (
    BaseDocumentAnalyzer,
    detect_doc_type,
    get_blueprint_dir,
    iterate_files_with_com,
    list_all_blueprints,
    load_blueprint,
    save_blueprint,
)
from .excel_analyzer import ExcelDocumentAnalyzer
from .hwp_analyzer import HwpDocumentAnalyzer
from .ppt_analyzer import PptDocumentAnalyzer
from .word_analyzer import WordDocumentAnalyzer

ANALYZERS: dict[str, type[BaseDocumentAnalyzer]] = {
    "ppt": PptDocumentAnalyzer,
    "excel": ExcelDocumentAnalyzer,
    "word": WordDocumentAnalyzer,
    "hwp": HwpDocumentAnalyzer,
}


def get_analyzer_for_type(doc_type: str) -> BaseDocumentAnalyzer:
    """문서 종류에 맞는 전용 분석기 인스턴스 반환"""
    dt = doc_type.lower()
    if dt not in ANALYZERS:
        raise ValueError(f"지원하지 않는 문서 종류입니다: {doc_type} (지원: ppt, excel, word, hwp)")
    return ANALYZERS[dt]()

