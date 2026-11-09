# -*- coding: utf-8 -*-
"""
mcp_server/com_manager.py
MS Word, Excel, PowerPoint 및 한컴오피스(HWP) 실시간 데스크톱 COM/ROT 바인딩 매니저.
- stdio JSON-RPC 스트림 오염 방지를 위해 모든 로그는 stderr로만 출력합니다.
- HWP는 Running Object Table (ROT)의 !HwpObject 모니커를 탐색하여 이미 열려 있는 한글 창에 즉시 연결합니다.
"""

from __future__ import annotations

import ctypes
import logging
import os
import sys
import winreg
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
import win32gui

# 모든 로깅을 stderr로 강제하여 FastMCP stdio 파이프 오염 방지
logger = logging.getLogger("mcp_server.com_manager")
if not logger.handlers:
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("[%(levelname)s] %(name)s: %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)


# 주요 비즈니스 및 오피스 프리셋 색상 (HEX)
NAMED_COLORS_HEX: dict[str, str] = {
    "business_navy": "#1F3864",
    "teal": "#0F766E",
    "emerald": "#16A34A",
    "ink": "#1E293B",
    "muted": "#64748B",
    "slate": "#475569",
    "ib_steel_blue": "#2E608A",
    "navy": "#1F3864",
    "deep_blue": "#002060",
    "blue": "#2563EB",
    "dark_gray": "#333333",
    "light_gray": "#F2F2F2",
    "border_gray": "#D9D9D9",
    "black": "#000000",
    "white": "#FFFFFF",
    "red": "#C00000",
    "green": "#27A066",
    "yellow": "#FFD966",
}


def rgb_to_bgr(r: int, g: int, b: int) -> int:
    """MS Office / HWP COM 표준 COLORREF BGR 정수 변환 (r + g*256 + b*65536)"""
    return (int(r) & 0xFF) | ((int(g) & 0xFF) << 8) | ((int(b) & 0xFF) << 16)


def hex_to_rgb(color_str: str | None, default: tuple[int, int, int] = (0, 0, 0)) -> tuple[int, int, int]:
    """'#1F3864', '#FFF', 'navy' 같은 문자열을 (R, G, B) 튜플로 변환. 해석할 수 없으면 ValueError"""
    if not color_str:
        return default
    cleaned = color_str.strip().lower()
    if cleaned in NAMED_COLORS_HEX:
        cleaned = NAMED_COLORS_HEX[cleaned]
    cleaned = cleaned.lstrip("#")
    if len(cleaned) == 3:  # 축약형 (예: FFF -> FFFFFF)
        cleaned = "".join(ch * 2 for ch in cleaned)
    if len(cleaned) == 8:  # ARGB (예: FF2E608A) -> 뒤 6자리 RGB
        cleaned = cleaned[2:]
    if len(cleaned) == 6:
        try:
            return (int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16))
        except ValueError:
            pass
    raise ValueError(f"알 수 없는 색상입니다: {color_str} (#RRGGBB 또는 등록된 색상 이름을 쓰세요)")


def hex_to_bgr(color_str: str | None, default_rgb: tuple[int, int, int] = (0, 0, 0)) -> int:
    """'#1F3864' 또는 색상 이름을 COM BGR 정수로 변환"""
    r, g, b = hex_to_rgb(color_str, default=default_rgb)
    return rgb_to_bgr(r, g, b)


def ensure_com() -> None:
    """현재 스레드에 대해 COM 아파트먼트 초기화를 보장"""
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass


@contextmanager
def word_undo_group(word_app: Any, record_name: str = "MCP AI 작업"):
    """Word COM 다중 작업을 단일 Ctrl+Z 실행취소 그룹으로 묶음"""
    started = False
    try:
        word_app.UndoRecord.StartCustomRecord(record_name[:60])
        started = True
    except Exception:
        pass
    try:
        yield
    finally:
        if started:
            try:
                word_app.UndoRecord.EndCustomRecord()
            except Exception:
                pass


class OfficeComManager:
    """
    MS Word, Excel, PowerPoint, 한컴 HWP의 실시간 데스크톱 COM 연결 및 상태 관리자.
    사용자가 이미 열어둔 창(GetActiveObject / ROT)에 최우선 바인딩합니다.
    """

    PROG_IDS = {
        "word": ["Word.Application"],
        "excel": ["Excel.Application"],
        "ppt": ["PowerPoint.Application"],
        "hwp": ["HWPFrame.HwpObject"],
    }

    def __init__(self) -> None:
        self._word_cached: Any = None
        self._excel_cached: Any = None
        self._ppt_cached: Any = None
        self._hwp_cached: Any = None

    def _find_office_app_from_hwnd(self, top_class: str, child_classes: tuple[str, ...]) -> Any:
        """
        Windows oleacc.dll AccessibleObjectFromWindow (OBJID_NATIVEOM = -16)를 이용해
        포커스를 잃기 전이라 ROT(GetActiveObject)에 아직 등록되지 않은 실시간 데스크톱
        Word(OpusApp->_WwG), Excel(XLMAIN->EXCEL7), PowerPoint(PPTFrameClass->mdiClass) 창에서도
        100% 결정적으로 Application COM 객체를 추출합니다.
        """
        ensure_com()
        found_app: list[Any] = []
        iid_bytes = bytes(pythoncom.IID_IDispatch)

        def _get_native_om(hwnd_child: int) -> Any:
            try:
                buf = (ctypes.c_char * 16).from_buffer_copy(iid_bytes)
                ptr = ctypes.c_void_p()
                res = ctypes.windll.oleacc.AccessibleObjectFromWindow(
                    int(hwnd_child), -16, ctypes.byref(buf), ctypes.byref(ptr)
                )
                if res == 0 and ptr.value:
                    disp = pythoncom.ObjectFromAddress(ptr.value, pythoncom.IID_IDispatch)
                    return win32.Dispatch(disp)
            except Exception:
                pass
            return None

        def _enum_top(hwnd: int, _: Any) -> bool:
            if found_app:
                return False
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if win32gui.GetClassName(hwnd) != top_class:
                    return True

                def _enum_child(ch: int, __: Any) -> bool:
                    if found_app:
                        return False
                    try:
                        if win32gui.GetClassName(ch) in child_classes:
                            om = _get_native_om(ch)
                            if om is not None:
                                app = om.Application
                                found_app.append(app)
                                return False
                    except Exception:
                        pass
                    return True

                win32gui.EnumChildWindows(hwnd, _enum_child, None)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(_enum_top, None)
        except Exception:
            pass
        return found_app[0] if found_app else None

    def is_app_installed(self, app_name: str) -> bool:
        """윈도우 레지스트리(HKCR)를 통해 해당 오피스 앱 COM ProgID가 등록되어 있는지 확인"""
        prog_ids = self.PROG_IDS.get(app_name.lower(), [])
        for prog_id in prog_ids:
            try:
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID", 0, winreg.KEY_READ)
                winreg.CloseKey(key)
                return True
            except OSError:
                continue
        return False

    # =========================================================================
    # 1. MS Word 연결
    # =========================================================================
    def get_word(
        self, create_if_missing: bool = True, visible: bool = True, add_blank_doc: bool = True
    ) -> tuple[Any, Any]:
        """(word_app, active_doc) 튜플 반환. 실행 중인 Word 창을 1순위로 바인딩."""
        ensure_com()
        word_app = None

        # 1. 캐시된 Word 인스턴스 유효성 확인
        if self._word_cached is not None:
            try:
                _ = self._word_cached.Documents.Count
                word_app = self._word_cached
            except Exception:
                self._word_cached = None

        # 2. GetActiveObject 시도
        if word_app is None:
            try:
                word_app = win32.GetActiveObject("Word.Application")
            except Exception:
                word_app = None

        # 3. Win32 창 핸들(OpusApp -> _WwG)에서 Native OM 직접 추출
        if word_app is None:
            word_app = self._find_office_app_from_hwnd("OpusApp", ("_WwG",))

        # 4. 신규 실행
        if word_app is None:
            if not create_if_missing:
                return None, None
            try:
                word_app = win32.Dispatch("Word.Application")
            except Exception as e:
                logger.error(f"Word COM 연결 실패: {e}")
                return None, None

        try:
            if visible:
                word_app.Visible = True
                try:
                    word_app.UserControl = True
                except Exception:
                    pass
            if word_app.Documents.Count == 0 and create_if_missing and add_blank_doc:
                word_app.Documents.Add()
            self._word_cached = word_app
            doc = word_app.ActiveDocument if word_app.Documents.Count > 0 else None
            return word_app, doc
        except Exception as e:
            logger.error(f"Word 문서 활성화 오류: {e}")
            return word_app, None

    # =========================================================================
    # 2. MS Excel 연결
    # =========================================================================
    def get_excel(
        self, create_if_missing: bool = True, visible: bool = True, add_blank_doc: bool = True
    ) -> tuple[Any, Any, Any]:
        """(excel_app, active_wb, active_sheet) 튜플 반환. 실행 중인 Excel 창을 1순위로 바인딩."""
        ensure_com()
        excel_app = None

        # 1. 캐시된 Excel 인스턴스 유효성 확인
        if self._excel_cached is not None:
            try:
                _ = self._excel_cached.Workbooks.Count
                excel_app = self._excel_cached
            except Exception:
                self._excel_cached = None

        # 2. GetActiveObject 시도
        if excel_app is None:
            try:
                excel_app = win32.GetActiveObject("Excel.Application")
            except Exception:
                excel_app = None

        # 3. Win32 창 핸들(XLMAIN -> EXCEL7)에서 Native OM 직접 추출
        if excel_app is None:
            excel_app = self._find_office_app_from_hwnd("XLMAIN", ("EXCEL7",))

        # 4. 신규 실행
        if excel_app is None:
            if not create_if_missing:
                return None, None, None
            try:
                excel_app = win32.Dispatch("Excel.Application")
            except Exception as e:
                logger.error(f"Excel COM 연결 실패: {e}")
                return None, None, None

        try:
            if visible:
                excel_app.Visible = True
                try:
                    excel_app.UserControl = True
                except Exception:
                    pass
            if excel_app.Workbooks.Count == 0 and create_if_missing and add_blank_doc:
                excel_app.Workbooks.Add()
            self._excel_cached = excel_app
            wb = excel_app.ActiveWorkbook if excel_app.Workbooks.Count > 0 else None
            sheet = wb.ActiveSheet if wb is not None else None
            return excel_app, wb, sheet
        except Exception as e:
            logger.error(f"Excel 워크북 활성화 오류: {e}")
            return excel_app, None, None

    # =========================================================================
    # 3. MS PowerPoint 연결
    # =========================================================================
    def get_ppt(
        self, create_if_missing: bool = True, visible: bool = True, add_blank_doc: bool = True
    ) -> tuple[Any, Any]:
        """(ppt_app, active_prs) 튜플 반환. 실행 중인 PowerPoint 창을 1순위로 바인딩."""
        ensure_com()
        ppt_app = None

        # 1. 캐시된 PowerPoint 인스턴스 유효성 확인
        if self._ppt_cached is not None:
            try:
                _ = self._ppt_cached.Presentations.Count
                ppt_app = self._ppt_cached
            except Exception:
                self._ppt_cached = None

        # 2. GetActiveObject 시도
        if ppt_app is None:
            try:
                ppt_app = win32.GetActiveObject("PowerPoint.Application")
            except Exception:
                ppt_app = None

        # 3. Win32 창 핸들(PPTFrameClass -> mdiClass / paneClassDC)에서 Native OM 직접 추출
        if ppt_app is None:
            ppt_app = self._find_office_app_from_hwnd("PPTFrameClass", ("mdiClass", "paneClassDC"))

        # 4. 신규 실행
        if ppt_app is None:
            if not create_if_missing:
                return None, None
            try:
                ppt_app = win32.Dispatch("PowerPoint.Application")
            except Exception as e:
                logger.error(f"PowerPoint COM 연결 실패: {e}")
                return None, None

        try:
            if visible:
                ppt_app.Visible = True
                try:
                    ppt_app.UserControl = True
                except Exception:
                    pass
            if ppt_app.Presentations.Count == 0 and create_if_missing and add_blank_doc:
                ppt_app.Presentations.Add()
            self._ppt_cached = ppt_app
            prs = ppt_app.ActivePresentation if ppt_app.Presentations.Count > 0 else None
            return ppt_app, prs
        except Exception as e:
            logger.error(f"PowerPoint 프레젠테이션 활성화 오류: {e}")
            return ppt_app, None

    def get_or_create_ppt_app(self, visible: bool = True) -> Any:
        app, _ = self.get_ppt(create_if_missing=True, visible=visible, add_blank_doc=False)
        return app

    def get_or_create_word_app(self, visible: bool = True) -> Any:
        app, _ = self.get_word(create_if_missing=True, visible=visible, add_blank_doc=False)
        return app

    def get_or_create_excel_app(self, visible: bool = True) -> Any:
        app, _ = self.get_excel(create_if_missing=True, visible=visible, add_blank_doc=False)
        return app

    def get_or_create_hwp_app(self, visible: bool = True) -> Any:
        app, _ = self.get_hwp(create_if_missing=True, visible=visible, add_blank_doc=False)
        return app

    # =========================================================================
    # 4. 한컴오피스(HWP) ROT (!HwpObject) 실시간 연결
    # =========================================================================
    def _find_running_hwp_from_rot(self) -> Any:
        """
        Windows Running Object Table (ROT)을 순회하여 이미 켜져 있는 한글(!HwpObject.*) 창에 바인딩.
        KISA 범피스 1310의 '한컴열기' 방식과 동일한 원리입니다.
        """
        ensure_com()
        try:
            context = pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning():
                try:
                    display_name = moniker.GetDisplayName(context, moniker)
                except Exception:
                    continue
                if display_name and ("!HwpObject" in display_name or "HWPFrame.HwpObject" in display_name):
                    try:
                        obj = rot.GetObject(moniker)
                        dispatch_iface = obj.QueryInterface(pythoncom.IID_IDispatch)
                        hwp = win32.Dispatch(dispatch_iface)
                        return hwp
                    except Exception:
                        continue
        except Exception as e:
            logger.debug(f"ROT HWP 탐색 중 예외: {e}")
        return None

    def get_hwp(self, create_if_missing: bool = True, visible: bool = True) -> Any:
        """
        실행 중인 한컴오피스(HWP) 객체를 반환.
        1순위: 캐시된 유효 객체 -> 2순위: ROT(!HwpObject) 활성 창 -> 3순위: GetActiveObject -> 4순위: Dispatch
        """
        ensure_com()

        # 1. 기존 캐시된 HWP 객체 유효성 확인
        if self._hwp_cached is not None:
            try:
                _ = self._hwp_cached.PageCount
                if visible:
                    try:
                        self._hwp_cached.XHwpWindows.Item(0).Visible = True
                    except Exception:
                        pass
                return self._hwp_cached
            except Exception:
                self._hwp_cached = None

        # 2. ROT에서 실행 중인 한글 창 탐색
        hwp = self._find_running_hwp_from_rot()

        # 3. GetActiveObject 시도
        if hwp is None:
            try:
                hwp = win32.GetActiveObject("HWPFrame.HwpObject")
            except Exception:
                hwp = None

        # 4. 신규 생성 (create_if_missing=True 인 경우)
        if hwp is None and create_if_missing:
            try:
                try:
                    from win32com.client.gencache import EnsureDispatch
                    hwp = EnsureDispatch("HWPFrame.HwpObject")
                except Exception:
                    hwp = win32.Dispatch("HWPFrame.HwpObject")
            except Exception as e:
                logger.error(f"HWP COM 객체 생성 실패: {e}")
                return None

        if hwp is not None:
            # 보안 팝업 방지 모듈 등록
            try:
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckDLLModule")
            except Exception:
                pass
            if visible:
                try:
                    hwp.XHwpWindows.Item(0).Visible = True
                except Exception:
                    pass
            self._hwp_cached = hwp

        return hwp

    # =========================================================================
    # 5. 파일 데스크톱 즉시 열기 & 전체 상태 조회
    # =========================================================================
    def prepare_output_path(self, file_path: str | Path) -> str:
        """
        출력 파일을 저장하기 전에:
        1) 이미 Word/Excel/PowerPoint 또는 ROT에 동일 파일이 열려 있으면 닫기를 시도합니다.
        2) 만약 여전히 OS 파일 잠금이 걸려 있다면, 자동으로 _2, _3 접미사가 붙은 경로를 반환하여
           PermissionError 없이 100% 안전하게 저장되도록 보장합니다.
        """
        ensure_com()
        p = Path(file_path).resolve()
        p.parent.mkdir(parents=True, exist_ok=True)
        abs_target = os.path.normcase(str(p))
        target_name = os.path.normcase(p.name)
        ext = p.suffix.lower()

        # 1. ROT 및 활성 앱에서 동일 파일 닫기 시도
        try:
            ctx = pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning():
                try:
                    disp_name = moniker.GetDisplayName(ctx, moniker)
                    if disp_name and os.path.normcase(disp_name) == abs_target:
                        obj = win32.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
                        try:
                            obj.Close(False)
                        except Exception:
                            try:
                                obj.Close()
                            except Exception:
                                pass
                except Exception:
                    continue
        except Exception:
            pass

        try:
            if ext in (".docx", ".doc"):
                word_app, _ = self.get_word(create_if_missing=False)
                if word_app is not None:
                    for i in range(word_app.Documents.Count, 0, -1):
                        d = word_app.Documents(i)
                        if os.path.normcase(d.FullName) == abs_target or os.path.normcase(d.Name) == target_name:
                            d.Close(0)
            elif ext in (".xlsx", ".xls"):
                excel_app, _, _ = self.get_excel(create_if_missing=False)
                if excel_app is not None:
                    for i in range(excel_app.Workbooks.Count, 0, -1):
                        w = excel_app.Workbooks(i)
                        if os.path.normcase(w.FullName) == abs_target or os.path.normcase(w.Name) == target_name:
                            w.Close(False)
            elif ext in (".pptx", ".ppt"):
                ppt_app, _ = self.get_ppt(create_if_missing=False)
                if ppt_app is not None:
                    for i in range(ppt_app.Presentations.Count, 0, -1):
                        pr = ppt_app.Presentations(i)
                        if os.path.normcase(pr.FullName) == abs_target or os.path.normcase(pr.Name) == target_name:
                            pr.Close()
        except Exception as e:
            logger.debug(f"열려 있는 파일 닫기 중 무시 가능한 예외: {e}")

        # 2. 쓰기 가능 여부 확인 후 잠겨 있으면 대체 파일명 생성
        def _is_writable(candidate: Path) -> bool:
            if not candidate.exists():
                return True
            try:
                with open(candidate, "ab"):
                    return True
            except PermissionError:
                return False

        if _is_writable(p):
            return str(p)

        for idx in range(2, 100):
            alt = p.with_name(f"{p.stem}_{idx}{p.suffix}")
            if _is_writable(alt):
                return str(alt)
        return str(p)

    def open_file_in_desktop(self, app_name: str, file_path: str | Path, visible: bool = True) -> dict[str, Any]:
        """
        생성된 문서 파일(.docx, .xlsx, .pptx, .hwp)을 데스크톱 앱 COM으로 즉시 화면에 띄웁니다.
        문서(Documents.Open / Workbooks.Open / Presentations.Open)를 먼저 연 직후
        Visible = True 및 UserControl = True를 설정하여 프로세스 종료 후에도 데스크톱 창이 영구 유지되도록 합니다.
        """
        ensure_com()
        abs_path = str(Path(file_path).resolve())
        if not os.path.exists(abs_path):
            return {"success": False, "error": f"파일이 존재하지 않습니다: {abs_path}"}

        app_key = app_name.lower().strip()
        try:
            if app_key in ("word", "docx"):
                word_app, _ = self.get_word(create_if_missing=True, visible=False, add_blank_doc=False)
                if word_app is None:
                    return {"success": False, "error": "MS Word를 실행할 수 없습니다."}
                doc = word_app.Documents.Open(abs_path)
                if visible:
                    word_app.Visible = True
                    try:
                        word_app.Activate()
                    except Exception:
                        pass
                return {"success": True, "app": "Word", "document_name": doc.Name, "path": abs_path}

            elif app_key in ("excel", "xlsx"):
                excel_app, _, _ = self.get_excel(create_if_missing=True, visible=False, add_blank_doc=False)
                if excel_app is None:
                    return {"success": False, "error": "MS Excel을 실행할 수 없습니다."}
                wb = excel_app.Workbooks.Open(abs_path)
                if visible:
                    excel_app.Visible = True
                    try:
                        excel_app.UserControl = True
                    except Exception:
                        pass
                return {"success": True, "app": "Excel", "workbook_name": wb.Name, "path": abs_path}

            elif app_key in ("ppt", "powerpoint", "pptx"):
                ppt_app, _ = self.get_ppt(create_if_missing=True, visible=True, add_blank_doc=False)
                if ppt_app is None:
                    return {"success": False, "error": "MS PowerPoint를 실행할 수 없습니다."}
                prs = ppt_app.Presentations.Open(abs_path)
                if visible:
                    ppt_app.Visible = True
                return {
                    "success": True,
                    "app": "PowerPoint",
                    "presentation_name": prs.Name,
                    "slide_count": prs.Slides.Count,
                    "path": abs_path,
                }

            elif app_key in ("hwp", "hwpx", "hancom"):
                hwp = self.get_hwp(create_if_missing=True, visible=visible)
                if hwp is None:
                    return {"success": False, "error": "한컴오피스(HWP)를 실행할 수 없습니다."}
                hwp.Open(abs_path)
                return {"success": True, "app": "HWP", "page_count": hwp.PageCount, "path": abs_path}

            return {"success": False, "error": f"지원하지 않는 앱 종류입니다: {app_name}"}
        except Exception as e:
            return {"success": False, "error": f"{app_name} 파일 열기 실패: {e}"}

    def activate_app(self, app_name: str) -> bool:
        """지정한 오피스 앱(word, excel, ppt, hwp) 창을 사용자의 화면 최상단으로 포커스/활성화"""
        ensure_com()
        key = app_name.lower().strip()
        try:
            if key in ("word", "docx"):
                app, doc = self.get_word(create_if_missing=False)
                if app is not None:
                    app.Visible = True
                    try:
                        app.Activate()
                    except Exception:
                        pass
                    if doc is not None:
                        try:
                            doc.Activate()
                        except Exception:
                            pass
                    return True
            elif key in ("excel", "xlsx"):
                app, wb, _ = self.get_excel(create_if_missing=False)
                if app is not None:
                    app.Visible = True
                    try:
                        app.UserControl = True
                        app.ScreenUpdating = True
                    except Exception:
                        pass
                    try:
                        if hasattr(app, "ActiveWindow") and app.ActiveWindow:
                            app.ActiveWindow.Activate()
                    except Exception:
                        pass
                    return True
            elif key in ("ppt", "powerpoint", "pptx"):
                app, prs = self.get_ppt(create_if_missing=False)
                if app is not None:
                    app.Visible = True
                    try:
                        app.Activate()
                    except Exception:
                        pass
                    return True
            elif key in ("hwp", "hancom"):
                hwp = self.get_hwp(create_if_missing=False)
                if hwp is not None:
                    try:
                        hwp.XHwpWindows.Item(0).Visible = True
                    except Exception:
                        pass
                    return True
        except Exception as e:
            logger.debug(f"{app_name} 창 활성화 실패: {e}")
        return False

    def get_status(self) -> dict[str, Any]:
        """현재 PC의 4대 오피스(Word, Excel, PowerPoint, HWP) 설치 및 실시간 실행 상태 요약 반환"""
        ensure_com()
        status: dict[str, Any] = {}

        # 1. Word 상태
        word_app, doc = self.get_word(create_if_missing=False)
        word_info: dict[str, Any] = {
            "installed": self.is_app_installed("word"),
            "running": word_app is not None,
        }
        if word_app is not None and doc is not None:
            try:
                sel_text = ""
                try:
                    sel_text = (word_app.Selection.Text or "").strip()[:200]
                except Exception:
                    pass
                word_info.update({
                    "active_document": doc.Name,
                    "paragraphs_count": doc.Paragraphs.Count,
                    "tables_count": doc.Tables.Count,
                    "selection_preview": sel_text,
                })
            except Exception:
                pass
        status["word"] = word_info

        # 2. Excel 상태
        excel_app, wb, sheet = self.get_excel(create_if_missing=False)
        excel_info: dict[str, Any] = {
            "installed": self.is_app_installed("excel"),
            "running": excel_app is not None,
        }
        if excel_app is not None and wb is not None:
            try:
                sel_addr = ""
                try:
                    sel_addr = excel_app.Selection.Address
                except Exception:
                    pass
                excel_info.update({
                    "active_workbook": wb.Name,
                    "active_sheet": sheet.Name if sheet else None,
                    "sheets": [ wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1) ],
                    "selection_address": sel_addr,
                })
            except Exception:
                pass
        status["excel"] = excel_info

        # 3. PowerPoint 상태
        ppt_app, prs = self.get_ppt(create_if_missing=False)
        ppt_info: dict[str, Any] = {
            "installed": self.is_app_installed("ppt"),
            "running": ppt_app is not None,
        }
        if ppt_app is not None and prs is not None:
            try:
                ppt_info.update({
                    "active_presentation": prs.Name,
                    "slides_count": prs.Slides.Count,
                })
            except Exception:
                pass
        status["powerpoint"] = ppt_info

        # 4. HWP 상태
        hwp = self.get_hwp(create_if_missing=False)
        hwp_info: dict[str, Any] = {
            "installed": self.is_app_installed("hwp"),
            "running": hwp is not None,
        }
        if hwp is not None:
            try:
                path_str = ""
                try:
                    path_str = hwp.Path
                except Exception:
                    pass
                hwp_info.update({
                    "page_count": hwp.PageCount,
                    "path": path_str,
                })
            except Exception:
                pass
        status["hwp"] = hwp_info

        return status


# 전역 싱글톤 인스턴스
com_manager = OfficeComManager()
