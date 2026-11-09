"""
Lightweight Zero-Dependency Office Document Reader (Inspired by Presenton)

Reads .docx, .pptx, and .xlsx files directly using Python's standard `zipfile`
and `xml.etree.ElementTree` in milliseconds without launching Microsoft Office COM.
"""

import os
import re
import zipfile
from pathlib import Path
from typing import Dict, Any, List, Optional
from xml.etree import ElementTree


class LightOfficeReader:
    """Reads office file contents headlessly using zero external dependencies."""

    @staticmethod
    def _local_name(tag: str) -> str:
        return tag.rsplit("}", 1)[-1]

    @staticmethod
    def _text_nodes(root: ElementTree.Element) -> List[str]:
        return [
            value
            for elem in root.iter()
            if LightOfficeReader._local_name(elem.tag) == "t"
            and (value := (elem.text or "").strip())
        ]

    @classmethod
    def read_docx_text(cls, file_path: str) -> str:
        """Extracts text paragraphs from a Word (.docx) file."""
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        paragraphs = []
        with zipfile.ZipFile(file_path, "r") as archive:
            try:
                xml_data = archive.read("word/document.xml")
                root = ElementTree.fromstring(xml_data)
                for p in root.iter():
                    if cls._local_name(p.tag) == "p":
                        text = "".join(cls._text_nodes(p))
                        if text:
                            paragraphs.append(text)
            except KeyError:
                pass
        return "\n".join(paragraphs)

    @classmethod
    def read_pptx_slides(cls, file_path: str) -> List[Dict[str, Any]]:
        """Extracts slides and their inner text boxes from a PowerPoint (.pptx) file."""
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        slides = []
        with zipfile.ZipFile(file_path, "r") as archive:
            slide_files = [
                name for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ]
            # Natural sort: slide1.xml, slide2.xml, ... slide10.xml
            slide_files.sort(key=lambda s: int(re.search(r"\d+", s).group(0)) if re.search(r"\d+", s) else 0)

            for idx, slide_path in enumerate(slide_files, 1):
                try:
                    xml_data = archive.read(slide_path)
                    root = ElementTree.fromstring(xml_data)
                    slide_texts = cls._text_nodes(root)
                    slides.append({
                        "slide_index": idx,
                        "file_member": slide_path,
                        "texts": slide_texts,
                        "combined_text": " ".join(slide_texts)
                    })
                except Exception:
                    continue
        return slides

    @classmethod
    def read_xlsx_sheets(cls, file_path: str) -> List[str]:
        """Extracts sheet names from an Excel (.xlsx) workbook without COM."""
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        sheet_names = []
        with zipfile.ZipFile(file_path, "r") as archive:
            try:
                wb_xml = archive.read("xl/workbook.xml")
                root = ElementTree.fromstring(wb_xml)
                for sheet in root.iter():
                    if cls._local_name(sheet.tag) == "sheet":
                        name = sheet.attrib.get("name")
                        if name:
                            sheet_names.append(name)
            except KeyError:
                pass
        return sheet_names
