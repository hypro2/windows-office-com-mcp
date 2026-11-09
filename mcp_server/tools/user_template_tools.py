# -*- coding: utf-8 -*-
"""
mcp_server/tools/user_template_tools.py
사용자 제공 문서(PPTX/XLSX/DOCX/HWP 1건 이상) 역공학 분석 및 데스크톱 COM 실시간 복제 생성 MCP 도구 세트.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Literal

from mcp_server.compat import MCPServer

from .blueprint_analyzers import (
    detect_doc_type,
    get_analyzer_for_type,
    list_all_blueprints,
    load_blueprint,
    save_blueprint,
)
from .blueprint_renderers import get_renderer_for_type

logger = logging.getLogger("mcp_server.user_template_tools")


def register_user_template_tools(mcp: MCPServer) -> None:
    @mcp.tool()
    def office_learn_user_template(
        file_paths: list[str],
        template_name: str,
        doc_type: Literal["ppt", "excel", "word", "hwp", "auto"] = "auto",
        description: str = "",
    ) -> dict[str, Any]:
        """
        [사용자 특화 문서 역공학 학습]
        사용자가 제공한 1건 이상의 특정 문서(PPTX, XLSX, DOCX, HWP)를 분석하여
        고정된 외피 서식(프레임/여백/스타일)과 가변 데이터 영역(동적 슬롯)을 자동 분리 판별하고,
        이후 100% 동일하게 재생 가능한 선언적 청사진(Blueprint)으로 영구 적재합니다.

        - file_paths: 분석할 문서 경로 목록 (1건 이상)
        - template_name: 저장할 사용자 템플릿 이름 (예: '경영전략_주간보고', '여신_심사의견서')
        - doc_type: 'ppt' | 'excel' | 'word' | 'hwp' | 'auto' (기본값 'auto': 확장자 자동 감지)
        - description: 템플릿 설명
        """
        if not file_paths:
            return {"success": False, "error": "분석할 파일 경로가 최소 1건 이상 필요합니다."}

        try:
            # 1. 파일 검증 및 포맷 감지
            resolved_paths = [Path(p).resolve() for p in file_paths]
            for p in resolved_paths:
                if not p.exists():
                    return {"success": False, "error": f"파일을 찾을 수 없습니다: {p}"}

            eff_doc_type = doc_type
            if eff_doc_type == "auto":
                eff_doc_type = detect_doc_type(resolved_paths[0])

            analyzer = get_analyzer_for_type(eff_doc_type)

            # 2. DRM 친화적 자동 순차 개방-확인-종료 배치 루프 실행
            extracted_docs = analyzer.analyze_files_batch(resolved_paths, pause_sec=0.3)
            if not extracted_docs:
                return {"success": False, "error": "제공된 문서에서 유효한 COM 서식 데이터를 추출하지 못했습니다."}

            # 3. 다중 문서 교집합 diff 분석 및 청사진 일반화
            blueprint = analyzer.generalize_multiple(
                docs=extracted_docs,
                template_name=template_name,
                description=description,
            )

            # 4. 청사진 영구 저장
            saved_file = save_blueprint(blueprint)

            # 슬롯 목록 집계
            slots = []
            for page in blueprint.get("pages", []):
                for el in page.get("elements", []):
                    if not el.get("is_static", True):
                        slots.append({
                            "slot_id": el.get("slot_id"),
                            "label": el.get("label", ""),
                            "sample_text": el.get("text", "")[:30],
                        })

            return {
                "success": True,
                "template_name": template_name,
                "doc_type": eff_doc_type,
                "source_files_count": len(resolved_paths),
                "pages_count": len(blueprint.get("pages", [])),
                "dynamic_slots_count": len(slots),
                "dynamic_slots": slots,
                "blueprint_path": str(saved_file),
                "message": (
                    f"사용자 특화 템플릿 '{template_name}'({eff_doc_type.upper()})이 성공적으로 학습되었습니다. "
                    f"총 {len(blueprint.get('pages', []))}개 페이지/시트에서 {len(slots)}개의 가변 데이터 슬롯이 자동 추출되었습니다."
                ),
            }
        except Exception as e:
            logger.error(f"사용자 템플릿 학습 실패: {e}", exc_info=True)
            return {"success": False, "error": f"사용자 템플릿 학습 실패: {e}"}

    @mcp.tool()
    def office_capture_active_template(
        template_name: str,
        doc_type: Literal["ppt", "excel", "word", "hwp", "auto"] = "auto",
        description: str = "",
    ) -> dict[str, Any]:
        """
        [데스크톱 활성 창 즉시 캡처 학습]
        현재 모니터에 열려 있는 활성화된 오피스 프로그램(PowerPoint, Excel, Word, HWP) 창의 문서를
        별도 파일 저장 없이 그 자리에서 즉시 역공학 분석하여 사용자 템플릿으로 적재합니다.

        - template_name: 저장할 사용자 템플릿 식별자
        - doc_type: 'ppt' | 'excel' | 'word' | 'hwp' | 'auto' (기본값 'auto': 활성 창 자동 탐지)
        - description: 템플릿 설명
        """
        try:
            eff_doc_type = doc_type
            if eff_doc_type == "auto":
                # 활성 창 탐색 순서: PPT -> Excel -> Word -> HWP
                from ..com_manager import com_manager, ensure_com
                ensure_com()
                if com_manager.get_ppt(create_if_missing=False)[0] is not None:
                    eff_doc_type = "ppt"
                elif com_manager.get_excel(create_if_missing=False)[0] is not None:
                    eff_doc_type = "excel"
                elif com_manager.get_word(create_if_missing=False)[0] is not None:
                    eff_doc_type = "word"
                elif com_manager.get_hwp() is not None:
                    eff_doc_type = "hwp"
                else:
                    return {"success": False, "error": "현재 열려 있는 PowerPoint, Excel, Word, HWP 창을 찾을 수 없습니다."}

            analyzer = get_analyzer_for_type(eff_doc_type)
            doc_info = analyzer.analyze_active_window()

            blueprint = analyzer.generalize_multiple(
                docs=[doc_info],
                template_name=template_name,
                description=description,
            )

            saved_file = save_blueprint(blueprint)

            slots_count = sum(
                1 for p in blueprint.get("pages", [])
                for el in p.get("elements", [])
                if not el.get("is_static", True)
            )

            return {
                "success": True,
                "template_name": template_name,
                "doc_type": eff_doc_type,
                "dynamic_slots_count": slots_count,
                "blueprint_path": str(saved_file),
                "message": f"활성 창({eff_doc_type.upper()})으로부터 템플릿 '{template_name}'이 즉시 캡처 등록되었습니다.",
            }
        except Exception as e:
            return {"success": False, "error": f"활성 창 캡처 실패: {e}"}

    @mcp.tool()
    def office_list_user_templates(
        doc_type: Literal["ppt", "excel", "word", "hwp", "all"] = "all",
    ) -> dict[str, Any]:
        """
        [등록된 사용자 특화 템플릿 목록 조회]
        학습되어 영구 저장된 사용자 정의 서식 청사진 목록과 포맷, 슬롯 스키마를 반환합니다.
        """
        try:
            templates = list_all_blueprints(doc_type=doc_type)
            return {
                "success": True,
                "total_templates": len(templates),
                "templates": templates,
            }
        except Exception as e:
            return {"success": False, "error": f"템플릿 목록 조회 실패: {e}"}

    @mcp.tool()
    def office_inspect_user_template(
        template_name: str,
        doc_type: Literal["ppt", "excel", "word", "hwp", "auto"] = "auto",
    ) -> dict[str, Any]:
        """
        [사용자 특화 템플릿 상세 명세 조회]
        특정 템플릿의 전체 구조, 페이지/시트별 요소, 필수 입력 슬롯 목록 및 샘플 페이로드 형식을 반환합니다.
        """
        try:
            blueprint = load_blueprint(template_name, doc_type=doc_type)
            slots_map: dict[str, str] = {}
            for p in blueprint.get("pages", []):
                for el in p.get("elements", []):
                    if not el.get("is_static", True) and el.get("slot_id"):
                        slots_map[el["slot_id"]] = el.get("text", "") or el.get("label", "")

            return {
                "success": True,
                "template_name": blueprint.get("template_name", template_name),
                "doc_type": blueprint.get("doc_type"),
                "description": blueprint.get("description", ""),
                "canvas": blueprint.get("canvas", {}),
                "pages_count": len(blueprint.get("pages", [])),
                "required_slots": slots_map,
                "sample_payload": slots_map,
                "blueprint": blueprint,
            }
        except Exception as e:
            return {"success": False, "error": f"템플릿 상세 조회 실패: {e}"}

    @mcp.tool()
    def office_generate_from_user_template(
        template_name: str,
        payload: dict[str, Any],
        doc_type: Literal["ppt", "excel", "word", "hwp", "auto"] = "auto",
        output_path: str | None = None,
    ) -> dict[str, Any]:
        """
        [사용자 특화 템플릿 기반 실시간 COM 클론 생성]
        학습된 사용자 템플릿 청사진에 새로운 데이터를 주입하여 원본과 100% 동일한 비주얼/서식으로
        데스크톱 COM 창에 실시간 복제 렌더링(Playback)하고 창을 화면에 활성화합니다.

        - template_name: 사용할 등록된 템플릿 이름
        - payload: 슬롯별 입력 데이터 딕셔너리 (키: slot_id 또는 필드명, 값: 신규 텍스트/표/수치)
        - doc_type: 'ppt' | 'excel' | 'word' | 'hwp' | 'auto'
        - output_path: 저장할 파일 경로 (생략 시 임시 문서로 열려 데스크톱 창 활성화)
        """
        try:
            blueprint = load_blueprint(template_name, doc_type=doc_type)
            eff_doc_type = blueprint.get("doc_type", "ppt")

            renderer = get_renderer_for_type(eff_doc_type)
            result = renderer.render(
                blueprint=blueprint,
                payload=payload,
                output_path=output_path,
            )
            return result
        except Exception as e:
            logger.error(f"템플릿 기반 문서 생성 실패: {e}", exc_info=True)
            return {"success": False, "error": f"템플릿 기반 문서 생성 실패: {e}"}

