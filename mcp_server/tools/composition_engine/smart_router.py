# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/smart_router.py
비정형 텍스트 성격 분석 기반 최적 오피스 포맷(PPT, Excel, HWP, DOCX) 자동 추천 라우터.
"""

from __future__ import annotations

from typing import Any


class SmartRouter:
    """오피스 문서 포맷 자동 판별 엔진"""

    EXCEL_KEYWORDS = ["데이터", "실적", "목표액", "집행액", "달성률", "합계", "평균", "단가", "수량", "금액", "수식", "시트", "표"]
    HWP_KEYWORDS = ["공문", "기안", "품의", "안건", "개조식", "A4", "1장", "제출", "행안부", "부서장", "지침", "규정"]
    PPT_KEYWORDS = ["발표", "슬라이드", "장표", "임원", "보고", "전략", "로드맵", "비전", "제안서", "요약"]
    DOCX_KEYWORDS = ["계약", "품의서", "기획서", "보고서", "WBS", "업무계획", "사내규정"]

    @classmethod
    def route_format(cls, raw_text: str, decomposed: dict[str, Any]) -> str:
        """
        텍스트 성분 및 키워드 가중치를 분석하여 최적의 문서 포맷을 결정
        Returns: 'ppt' | 'excel' | 'hwp' | 'docx'
        """
        text = raw_text.lower()

        excel_score = 0
        hwp_score = 0
        ppt_score = 0
        docx_score = 0

        # 1. 키워드 점수 계산
        for kw in cls.EXCEL_KEYWORDS:
            if kw in text:
                excel_score += 2
        for kw in cls.HWP_KEYWORDS:
            if kw in text:
                hwp_score += 2
        for kw in cls.PPT_KEYWORDS:
            if kw in text:
                ppt_score += 2
        for kw in cls.DOCX_KEYWORDS:
            if kw in text:
                docx_score += 2

        # 2. 성분 구조 점수 가산
        metrics_count = len(decomposed.get("metrics", []))
        chrono_count = len(decomposed.get("chronology", []))
        duality_count = len(decomposed.get("duality", []))

        # 수치가 압도적으로 많고 문장 수가 적으면 Excel 적합
        if metrics_count >= 4 and len(decomposed.get("raw_sentences", [])) <= 6:
            excel_score += 5

        # 일정/단계와 대조가 풍부하면 PPT 적합
        if chrono_count >= 2 or duality_count >= 1:
            ppt_score += 4

        # 개조식 또는 A4 1장 키워드가 있거나 문장이 길면 HWP/DOCX 가산
        if "1장" in text or "개조식" in text:
            hwp_score += 6
        if "품의" in text or "기획서" in text:
            docx_score += 4

        # 점수 집계 및 결정
        scores = {
            "ppt": ppt_score,
            "excel": excel_score,
            "hwp": hwp_score,
            "docx": docx_score,
        }

        best_format = max(scores, key=scores.get)
        # 모든 점수가 0이면 비주얼 임원 보고용 PPT 기본 추천
        if scores[best_format] == 0:
            return "ppt"

        return best_format
