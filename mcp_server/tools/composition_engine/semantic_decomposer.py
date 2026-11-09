# -*- coding: utf-8 -*-
"""
mcp_server/tools/composition_engine/semantic_decomposer.py
비정형 원시 텍스트 덤프(회의록, 메모, 기사 등)에서 4대 핵심 정보 성분을 자동 추출하는 시맨틱 분해기.
- Metrics: 수치, 금액, 비율, 달성률, KPI
- Chronology: 일정, 분기, 단계(Step), 마일스톤
- Duality: 현행 vs 개선, 기존 vs 신규, 문제점 vs 해결책
- Directives: 핵심 안건, 기대효과, 결론, 승인 요청
"""

from __future__ import annotations

import re
from typing import Any


class SemanticDecomposer:
    """원시 텍스트 4대 시맨틱 성분 분해 엔진"""

    # 수치 및 단위 정규식 (예: 120%, 3.5억, 45개사, 1.2만원, 32% 절감)
    METRIC_PATTERN = re.compile(
        r"([가-힣a-zA-Z\s]{1,15})?\s*([+-]?\d+(?:,\d+)*(?:\.\d+)?)\s*(%|억\s*원?|조\s*원?|만\s*원?|원|개사?|명|건|배|pt|점)"
    )

    # 시간 및 단계 정규식 (예: 1단계, STEP 01, 10월, 3분기, 상반기)
    CHRONO_PATTERN = re.compile(
        r"(\b\d{1,2}월|\b\d{1,2}분기|\b\d{4}년|상반기|하반기|\b\d+단계|\bSTEP\s*\d+|1차|2차|3차)"
    )

    # 대조 및 비교 키워드
    DUALITY_KEYWORDS = [
        "기존", "현행", "신규", "개선", "As-Is", "To-Be", "대비", "전환",
        "문제점", "해결책", "장점", "단점", "한계", "보완", "다변화", "비교"
    ]

    # 결론 및 지시 키워드
    DIRECTIVE_KEYWORDS = [
        "결론", "시사점", "제언", "요청", "승인", "기대효과", "핵심", "목적", "방침", "추진"
    ]

    @classmethod
    def decompose(cls, raw_text: str) -> dict[str, Any]:
        """비정형 텍스트를 문장 단위로 파싱하여 4대 성분으로 구조화"""
        # 줄바꿈 및 마침표 기준 문장 분리
        raw_lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
        sentences = []
        for line in raw_lines:
            # 문장 분리 (마침표, 줄바꿈)
            split_s = re.split(r"(?<=[.?!])\s+", line)
            for s in split_s:
                if s.strip():
                    sentences.append(s.strip())

        metrics: list[dict[str, str]] = []
        chronology: list[dict[str, str]] = []
        duality: list[dict[str, str]] = []
        directives: list[str] = []

        # 1. 수치(Metrics) 성분 추출
        seen_metrics = set()
        for s in sentences:
            matches = cls.METRIC_PATTERN.findall(s)
            for prefix, val_num, unit in matches:
                label = prefix.strip() if prefix else "지표"
                label = re.sub(r"^(은|는|이|가|을|를|의|에|에서)\s*", "", label).strip()
                if not label:
                    label = "핵심 지표"
                value = f"{val_num}{unit}"
                key = (label, value)
                if key not in seen_metrics:
                    seen_metrics.add(key)
                    metrics.append({
                        "label": label[:12],
                        "value": value,
                        "context": s[:50]
                    })

        # 2. 시간/단계(Chronology) 성분 추출
        step_idx = 1
        for s in sentences:
            m = cls.CHRONO_PATTERN.search(s)
            if m:
                chrono_tag = m.group(1)
                chronology.append({
                    "step": f"STEP 0{step_idx}",
                    "tag": chrono_tag,
                    "title": s[:30],
                    "desc": s[:60]
                })
                step_idx += 1

        # 3. 대조/비교(Duality) 성분 추출
        for s in sentences:
            if any(k in s for k in cls.DUALITY_KEYWORDS):
                duality.append({
                    "sentence": s,
                    "type": "contrast"
                })

        # 4. 결론/디렉티브(Directives) 성분 추출
        for s in sentences:
            if any(k in s for k in cls.DIRECTIVE_KEYWORDS):
                directives.append(s)

        # 수집 결과가 비어 있을 경우 폴백
        if not directives and sentences:
            directives.append(sentences[0])  # 첫 문장을 결론/주제로 활용

        return {
            "total_sentences": len(sentences),
            "metrics": metrics,
            "chronology": chronology,
            "duality": duality,
            "directives": directives,
            "raw_sentences": sentences,
        }
