"""
Surgical Document Patcher (Inspired by GenOffice text-patch.ts)

Applies pinpoint replacements to Word (.docx) and HWP documents by targeting
specific bookmarks, content controls, or paragraph fragments without
re-serializing the entire document, preserving styles and metadata byte-for-byte.
"""

import os
import re
from typing import Dict, Any, Optional, List, Tuple
from mcp_server.response_envelope import ok_response, error_response


class SurgicalPatcher:
    """Performs non-destructive, targeted surgical patches on office documents."""

    @staticmethod
    def extract_depth_aware_paragraphs(xml_content: str) -> List[Tuple[int, int, str]]:
        """
        Extracts outermost <w:p> slices accounting for nested <w:p> tags inside <w:txbxContent>,
        directly implementing GenOffice text-patch.ts paragraphSlices algorithm.
        Returns list of (start_offset, end_offset, inner_xml).
        """
        slices = []
        token_re = re.compile(r'<w:p(?:\s[^>]*)?\/>|<w:p(?:\s[^>]*)?>|<\/w:p>')
        depth = 0
        start = -1

        for m in token_re.finditer(xml_content):
            token = m.group(0)
            if token == '</w:p>':
                if depth == 0:
                    continue
                depth -= 1
                if depth == 0 and start != -1:
                    end = m.end()
                    slices.append((start, end, xml_content[start:end]))
                    start = -1
            elif token.endswith('/>'):
                if depth == 0:
                    slices.append((m.start(), m.end(), token))
            else:
                if depth == 0:
                    start = m.start()
                depth += 1

        return slices

    @staticmethod
    def patch_docx_bookmark_text(
        doc_object: Any,  # Word Document COM Object
        bookmark_name: str,
        new_text: str
    ) -> Dict[str, Any]:
        """
        Surgically patches the text of a specific bookmark in an active Word document
        while preserving font, style, and paragraph properties.
        """
        try:
            if not doc_object.Bookmarks.Exists(bookmark_name):
                return error_response(
                    command="surgical_patch_bookmark",
                    summary=f"북마크 '{bookmark_name}'를 찾을 수 없습니다.",
                    suggestion="문서 내 유효한 북마크 목록을 먼저 조회하세요.",
                    code="ERR_BOOKMARK_NOT_FOUND"
                )

            bm = doc_object.Bookmarks(bookmark_name)
            rng = bm.Range
            rng.Text = new_text
            doc_object.Bookmarks.Add(bookmark_name, rng)

            return ok_response(
                command="surgical_patch_bookmark",
                summary=f"북마크 '{bookmark_name}' 영역의 텍스트가 서식 손실 없이 수술적으로 교체되었습니다.",
                detail={"bookmark": bookmark_name, "new_length": len(new_text)}
            )
        except Exception as e:
            return error_response(
                command="surgical_patch_bookmark",
                summary=f"수술적 패치 중 오류 발생: {e}",
                suggestion="COM 프로세스 상태 및 문서 잠금 여부를 확인하세요.",
                code="ERR_COM_MUTATION"
            )

    @staticmethod
    def patch_string_common_affix(old_text: str, new_text: str) -> Dict[str, Any]:
        """
        Calculates common prefix and suffix to minimize byte modifications,
        mimicking GenOffice's minimal run replacement algorithm.
        """
        prefix_len = 0
        min_len = min(len(old_text), len(new_text))
        while prefix_len < min_len and old_text[prefix_len] == new_text[prefix_len]:
            prefix_len += 1

        suffix_len = 0
        while (suffix_len < min_len - prefix_len and 
               old_text[-(suffix_len + 1)] == new_text[-(suffix_len + 1)]):
            suffix_len += 1

        replaced_old = old_text[prefix_len:len(old_text) - suffix_len if suffix_len > 0 else None]
        replacing_new = new_text[prefix_len:len(new_text) - suffix_len if suffix_len > 0 else None]

        return {
            "prefix_preserved": old_text[:prefix_len],
            "suffix_preserved": old_text[len(old_text) - suffix_len:] if suffix_len > 0 else "",
            "delta_old": replaced_old,
            "delta_new": replacing_new,
            "unchanged": old_text == new_text
        }
