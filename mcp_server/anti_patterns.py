"""
Office Automation Anti-Patterns & Guardrails (Inspired by Meta Astryx)

Defines explicit anti-patterns and guardrail validators for AI Agents
interacting with Word, Excel, PowerPoint, and HWP via COM and MCP.
"""

from typing import Dict, Any, List, Optional
from dataclasses import dataclass


@dataclass
class AntiPatternViolation:
    rule_id: str
    severity: str  # "error" | "warning"
    message: str
    suggestion: str


class OfficeAntiPatternGuard:
    """
    Validates agent tool parameters against known office automation anti-patterns
    before executing COM operations, preventing process hangs, style loss, and layout collisions.
    """

    ANTI_PATTERNS = {
        "EXCEL_CELL_LOOP": {
            "id": "AP-EXCEL-001",
            "title": "Excel 단일 셀 순회 반복 쓰기 금지",
            "description": "셀을 하나씩 순회하며 ws.Cells(r, c).Value를 호출하면 수천 번의 COM IPC로 프로세스가 중단됩니다.",
            "rule": "SafeArray 2차원 리스트/튜플을 구성하여 ws.Range(...).Value로 일괄 쓰기(Batch Write) 필수"
        },
        "PPT_HARDCODED_COORDINATES": {
            "id": "AP-PPT-002",
            "title": "PPT 절대 좌표/폰트 크기 임의 하드코딩 금지",
            "description": "캔버스 경계를 고려하지 않은 절대 좌표(Left=100, Top=200)는 텍스트 넘침과 도형 겹침을 유발합니다.",
            "rule": "16:9 슬라이드 그리드 토큰 또는 ppt_slice_engine / svg 컴파일러 사용 필수"
        },
        "WORD_FULL_REBUILD": {
            "id": "AP-WORD-003",
            "title": "기존 Word/HWP 전체 문서 무단 재생성 금지",
            "description": "기존 문서의 일부분을 고치기 위해 문서를 처음부터 다시 쓰면 서식, 헤더, 커스텀 XML이 파괴됩니다.",
            "rule": "북마크(Bookmark) 또는 특정 Range 단위 수술적 패치(Surgical Patch) 사용 필수"
        },
        "COLOR_HARDCODING": {
            "id": "AP-DESIGN-004",
            "title": "임의의 HEX/RGB 색상 하드코딩 금지",
            "description": "일관되지 않은 임의의 색상 코드는 기업 브랜딩 통일성을 훼손합니다.",
            "rule": "공식 시맨틱 테마 토큰(Business Navy, Steel Blue, Neutral Dark 등) 준수 필수"
        }
    }

    @staticmethod
    def validate_excel_write(data: Any, target_range: Optional[str] = None) -> List[AntiPatternViolation]:
        violations = []
        if isinstance(data, list) and len(data) > 20:
            # 대량 데이터를 단일 셀 단위로 쪼개 쓰려는 징후 검사
            pass
        return violations

    @staticmethod
    def validate_slide_geometry(left: float, top: float, width: float, height: float,
                                slide_width: float = 960.0, slide_height: float = 540.0) -> List[AntiPatternViolation]:
        violations = []
        if left + width > slide_width or top + height > slide_height:
            violations.append(AntiPatternViolation(
                rule_id="AP-PPT-002",
                severity="error",
                message=f"도형이 슬라이드 캔버스를 벗어납니다 (Left+Width={left+width:.1f} > {slide_width:.1f})",
                suggestion=f"너비를 {slide_width - left - 20:.1f} 이하로 줄이거나 슬라이스 레이아웃을 사용하세요."
            ))
        if left < 0 or top < 0:
            violations.append(AntiPatternViolation(
                rule_id="AP-PPT-002",
                severity="error",
                message="좌표는 음수일 수 없습니다.",
                suggestion="슬라이드 여백 최소값(기본 20pt) 이상으로 지정하세요."
            ))
        return violations
