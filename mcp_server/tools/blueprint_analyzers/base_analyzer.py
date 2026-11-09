# -*- coding: utf-8 -*-
"""
mcp_server/tools/blueprint_analyzers/base_analyzer.py
4대 오피스 문서(PPTX, XLSX, DOCX, HWP) 역공학 분석 공통 기반 및 다중 문서 교집합(Diff) 일반화 엔진.
"""

from __future__ import annotations

import json
import logging
import re
from abc import ABC, abstractmethod
from pathlib import Path
import time
from typing import Any, Callable

logger = logging.getLogger("mcp_server.blueprint_analyzers")

BLUEPRINT_ROOT = Path(__file__).resolve().parents[2] / "blueprints"


def get_blueprint_dir(doc_type: str) -> Path:
    """문서 종류별 청사진 저장 폴더 반환 및 자동 생성"""
    target = BLUEPRINT_ROOT / doc_type.lower()
    target.mkdir(parents=True, exist_ok=True)
    return target


def detect_doc_type(file_path: str | Path) -> str:
    """파일 확장자 기반 4대 오피스 포맷 자동 감지"""
    suffix = Path(file_path).suffix.lower()
    if suffix in (".pptx", ".ppt"):
        return "ppt"
    if suffix in (".xlsx", ".xls", ".xlsm"):
        return "excel"
    if suffix in (".docx", ".doc"):
        return "word"
    if suffix in (".hwp", ".hwpx"):
        return "hwp"
    raise ValueError(f"지원하지 않는 문서 확장자입니다: {suffix} (지원: .pptx, .xlsx, .docx, .hwp)")


class BaseDocumentAnalyzer(ABC):
    """4대 오피스 포맷별 역공학 분석기 추상 기반 클래스"""

    doc_type: str = "base"

    @abstractmethod
    def analyze_file(self, file_path: Path) -> dict[str, Any]:
        """단일 파일 역공학 분석"""
        pass

    @abstractmethod
    def analyze_active_window(self) -> dict[str, Any]:
        """데스크톱에 현재 열려 있는 활성 창 역공학 분석"""
        pass

    def analyze_files_batch(self, file_paths: list[Path], pause_sec: float = 0.3) -> list[dict[str, Any]]:
        """복수 파일을 순차적으로 열고 확인하고 닫는 배치 분석 (DRM 안전 순회)"""
        return iterate_files_with_com(file_paths, self.analyze_file, pause_sec=pause_sec)

    @classmethod
    def generalize_multiple(cls, docs: list[dict[str, Any]], template_name: str, description: str = "") -> dict[str, Any]:
        """
        1건 이상의 분석 결과로부터 공통 불변 요소(정적 프레임)와
        가변 데이터 영역(동적 슬롯)을 자동 분리 판별하여 표준 청사진을 생성합니다.
        """
        if not docs:
            raise ValueError("분석할 문서 결과가 비어 있습니다.")

        doc_type = docs[0].get("doc_type", "unknown")

        if len(docs) == 1:
            blueprint = docs[0]
            blueprint["template_name"] = template_name
            blueprint["description"] = description
            blueprint["samples_count"] = 1
            return blueprint

        # 2건 이상 제공 시 교집합 diff 분석 수행
        first = docs[0]
        generalized: dict[str, Any] = {
            "template_name": template_name,
            "description": description,
            "doc_type": doc_type,
            "samples_count": len(docs),
            "canvas": first.get("canvas", {}),
            "design_tokens": first.get("design_tokens", {}),
            "pages": [],
        }

        # 페이지/슬라이드/시트별 비교
        max_pages = min(len(d.get("pages", [])) for d in docs)
        for p_idx in range(max_pages):
            base_page = docs[0]["pages"][p_idx]
            page_elements = base_page.get("elements", [])
            gen_elements = []

            for e_idx, elem in enumerate(page_elements):
                # 다른 문서들의 동일 위치 요소 텍스트 수집
                texts = []
                for d in docs:
                    d_page = d["pages"][p_idx]
                    if e_idx < len(d_page.get("elements", [])):
                        t = d_page["elements"][e_idx].get("text", "")
                        texts.append(str(t).strip())

                elem_copy = dict(elem)
                # 텍스트가 모든 문서에서 100% 동일하면 -> 정적 프레임(Static)
                # 텍스트가 문서마다 다르면 -> 가변 슬롯(Dynamic Slot)
                unique_texts = set(texts)
                if len(unique_texts) > 1:
                    elem_copy["is_static"] = False
                    elem_copy["slot_id"] = elem.get("slot_id") or f"slot_{p_idx + 1}_{e_idx + 1}"
                    elem_copy["label"] = f"가변 텍스트 ({texts[0][:20]})"
                else:
                    elem_copy["is_static"] = True

                gen_elements.append(elem_copy)

            generalized["pages"].append({
                "page_index": p_idx + 1,
                "title": base_page.get("title", f"페이지 {p_idx + 1}"),
                "elements": gen_elements,
            })

        return generalized


def save_blueprint(blueprint: dict[str, Any]) -> Path:
    """생성된 청사진을 JSON 파일로 영구 적재"""
    doc_type = blueprint.get("doc_type", "misc")
    t_name = blueprint.get("template_name", "unnamed")
    safe_name = re.sub(r'[\\/*?:"<>| ]', "_", t_name)

    folder = get_blueprint_dir(doc_type)
    file_path = folder / f"{safe_name}.json"
    file_path.write_text(json.dumps(blueprint, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info(f"사용자 템플릿 청사진 저장 완료: {file_path}")
    return file_path


def load_blueprint(template_name: str, doc_type: str | None = None) -> dict[str, Any]:
    """템플릿 이름으로 청사진 JSON 로드"""
    safe_name = re.sub(r'[\\/*?:"<>| ]', "_", template_name)

    search_types = [doc_type] if doc_type and doc_type != "auto" else ["ppt", "excel", "word", "hwp"]
    for dt in search_types:
        folder = get_blueprint_dir(dt)
        f = folder / f"{safe_name}.json"
        if f.exists():
            return json.loads(f.read_text(encoding="utf-8"))

    raise FileNotFoundError(f"등록된 템플릿 '{template_name}'을 찾을 수 없습니다.")


def list_all_blueprints(doc_type: str | None = None) -> list[dict[str, Any]]:
    """등록된 템플릿 목록 반환"""
    results: list[dict[str, Any]] = []
    search_types = [doc_type] if doc_type and doc_type not in ("all", "auto") else ["ppt", "excel", "word", "hwp"]

    for dt in search_types:
        folder = get_blueprint_dir(dt)
        for json_file in folder.glob("*.json"):
            try:
                data = json.loads(json_file.read_text(encoding="utf-8"))
                slots = []
                for p in data.get("pages", []):
                    for el in p.get("elements", []):
                        if not el.get("is_static", True):
                            slots.append({
                                "slot_id": el.get("slot_id"),
                                "label": el.get("label", ""),
                                "default_text": el.get("text", ""),
                            })

                results.append({
                    "template_name": data.get("template_name", json_file.stem),
                    "doc_type": dt,
                    "description": data.get("description", ""),
                    "pages_count": len(data.get("pages", [])),
                    "slots_count": len(slots),
                    "slots": slots[:15],
                    "file_path": str(json_file),
                })
            except Exception as e:
                logger.warning(f"청사진 파일 읽기 실패 ({json_file}): {e}")

    return results


def iterate_files_with_com(
    file_paths: list[Path],
    open_and_analyze_fn: Callable[[Path], dict[str, Any]],
    pause_sec: float = 0.3,
) -> list[dict[str, Any]]:
    """
    [DRM 안전 순차 배치 순회 엔진]
    복수 문서 파일을 COM을 통해 하나씩 순차적으로 열어 분석하고 즉시 닫는 과정을 반복합니다.
    - 무조건적 try...finally로 프로세스 누수 및 좀비 문서 방지
    - 파일 간 0.3초 딜레이로 사내 DRM 후킹 모듈의 메모리 락 해제 대기
    - 개별 파일 실패 시에도 나머지 정상 파일 순회 보장 (Fault-Tolerance)
    """
    results: list[dict[str, Any]] = []
    errors: list[str] = []

    total = len(file_paths)
    for idx, fpath in enumerate(file_paths, start=1):
        logger.info(f"[{idx}/{total}] COM 문서 자동 개방 및 확인 중: {fpath.name}")
        try:
            doc_data = open_and_analyze_fn(fpath)
            results.append(doc_data)
            logger.info(f"[{idx}/{total}] {fpath.name} 확인 완료 및 정상 닫기 성공")
        except Exception as e:
            err_msg = f"{fpath.name} 확인 실패: {e}"
            logger.error(f"[{idx}/{total}] {err_msg}", exc_info=True)
            errors.append(err_msg)

        if idx < total and pause_sec > 0:
            time.sleep(pause_sec)

    if not results and errors:
        raise RuntimeError(f"모든 문서의 COM 확인이 실패했습니다: {'; '.join(errors)}")

    return results


