# -*- coding: utf-8 -*-
"""WPS PDF 当前页截图测试脚本。

用法：
  默认打开小窗口（置顶按钮）：
       python _wps_pdf_capture_test.py
       python _wps_pdf_capture_test.py --gui

  命令行单次截图：
       python _wps_pdf_capture_test.py --once
       python _wps_pdf_capture_test.py --once --pdf "D:\\path\\file.pdf" --page 12

  热键模式（需 pip install keyboard）：
       python _wps_pdf_capture_test.py --hotkey

  诊断：
       python _wps_pdf_capture_test.py --probe

输出目录：{程序目录}\\PDFPageCaptures\\{文档名}\\{文档名}_第009页.png
支持 PDF（PyMuPDF）与 Word（WPS COM 导出指定页）。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import threading
import time
import urllib.parse
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    import fitz
except ImportError:
    print("缺少 PyMuPDF，请先运行: pip install pymupdf")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    Image = None  # type: ignore[misc, assignment]

try:
    import pythoncom
    import win32com.client
    import win32gui
except ImportError:
    print("缺少 pywin32，请先运行: pip install pywin32")
    sys.exit(1)

WORD_SUFFIXES = (".docx", ".doc")
PDF_SUFFIXES = (".pdf",)
DOCUMENT_SUFFIXES = PDF_SUFFIXES + WORD_SUFFIXES


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = app_dir()
DEFAULT_OUT = APP_DIR / "PDFPageCaptures"
DEFAULT_SEARCH_ROOT = Path(r"d:\工作")
CONFIG_PATH = DEFAULT_OUT / "_capture_config.json"
WPS_PROG_IDS = (
    "Kwps.Application",
    "KWPS.Application",
    "WPS.Application",
    "wps.Application",
)
ZOOM = 2.0
_FILE_SEARCH_CACHE: dict[str, Path] = {}


@dataclass
class AppConfig:
    search_root: Path = DEFAULT_SEARCH_ROOT
    last_pdf: Path | None = None
    last_page: int = 1

    @classmethod
    def load(cls) -> "AppConfig":
        if not CONFIG_PATH.exists():
            return cls()
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            last_pdf_raw = data.get("last_pdf")
            last_page = int(data.get("last_page") or 1)
            search_root = Path(data.get("search_root") or DEFAULT_SEARCH_ROOT)
            last_pdf = Path(last_pdf_raw) if last_pdf_raw else None
            if last_pdf and not last_pdf.exists():
                last_pdf = None
            if last_page < 1:
                last_page = 1
            return cls(search_root=search_root, last_pdf=last_pdf, last_page=last_page)
        except Exception:
            return cls()

    def save(self) -> None:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "search_root": str(self.search_root),
            "last_pdf": str(self.last_pdf) if self.last_pdf else "",
            "last_page": self.last_page,
        }
        CONFIG_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


@dataclass
class CaptureResult:
    ok: bool
    message: str
    out_png: Path | None = None
    elapsed: float = 0.0
    page: int | None = None
    pdf_path: Path | None = None


def ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def log(msg: str) -> None:
    print(msg, flush=True)


def connect_wps(quiet: bool = False):
    """Attach to running WPS, else start a new instance."""
    last_err = None
    for prog_id in WPS_PROG_IDS:
        try:
            app = win32com.client.GetActiveObject(prog_id)
            if not quiet:
                log(f"[COM] 已连接运行中的 WPS: {prog_id}")
            return app, prog_id
        except Exception as exc:
            last_err = exc
    for prog_id in WPS_PROG_IDS:
        try:
            app = win32com.client.Dispatch(prog_id)
            app.Visible = True
            if not quiet:
                log(f"[COM] 已启动 WPS: {prog_id}")
            return app, prog_id
        except Exception as exc:
            last_err = exc
    raise RuntimeError(f"无法连接 WPS COM: {last_err}")


def com_get_active_pdf(app):
    for name in ("ActivePDF", "ActivePdf", "activePDF"):
        try:
            pdf = getattr(app, name)
            if pdf is not None:
                return pdf
        except Exception:
            pass
    return None


def com_get_int(obj, *names: str) -> int | None:
    for name in names:
        try:
            val = getattr(obj, name)
            if val is not None:
                return int(val)
        except Exception:
            pass
    return None


def com_get_str(obj, *names: str) -> str | None:
    for names_chain in (names,):
        for name in names_chain:
            try:
                val = getattr(obj, name)
                if val:
                    return str(val).strip()
            except Exception:
                pass
    return None


def normalize_path(raw: str | None) -> Path | None:
    return normalize_document_path(raw, allowed_suffixes=PDF_SUFFIXES)


def normalize_document_path(
    raw: str | None,
    allowed_suffixes: tuple[str, ...] = DOCUMENT_SUFFIXES,
) -> Path | None:
    if not raw:
        return None
    text = raw.strip().strip('"')
    if text.startswith("file:///"):
        text = urllib.parse.unquote(text[8:])
    elif text.startswith("file://"):
        text = urllib.parse.unquote(text[7:])
    path = Path(text)
    if path.suffix.lower() in allowed_suffixes and path.exists():
        return path.resolve()
    return None


def document_kind(path: Path | None) -> str:
    if path is None:
        return "unknown"
    ext = path.suffix.lower()
    if ext in PDF_SUFFIXES:
        return "pdf"
    if ext in WORD_SUFFIXES:
        return "word"
    return "unknown"


def build_output_png(out_dir: Path, document_path: Path | None, page: int) -> Path:
    stem = document_path.stem if document_path else "wps_capture"
    return out_dir / stem / f"{stem}_第{page:03d}页.png"


def get_foreground_window_title() -> str:
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            return str(win32gui.GetWindowText(hwnd)).strip()
    except Exception:
        pass
    return ""


def collect_child_window_texts(hwnd, limit: int = 200) -> list[str]:
    texts: list[str] = []

    def _enum(child_hwnd, _):
        if len(texts) >= limit:
            return False
        try:
            text = win32gui.GetWindowText(child_hwnd).strip()
            if text:
                texts.append(text)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _enum, None)
    except Exception:
        pass
    return texts


def parse_page_from_ui_texts(texts: list[str]) -> int | None:
    patterns = (
        r"(?:^|\s)(\d+)\s*/\s*(\d+)(?:\s|$)",
        r"第\s*(\d+)\s*页",
        r"页码[:：]?\s*(\d+)",
        r"page\s*(\d+)",
    )
    joined = " | ".join(texts)
    for pat in patterns:
        m = re.search(pat, joined, flags=re.IGNORECASE)
        if m:
            page = int(m.group(1))
            if page > 0:
                return page
    return None


def parse_page_from_title(title: str) -> int | None:
    for pat in (
        r"第\s*(\d+)\s*页",
        r"[\s\-—](\d+)\s*/\s*\d+",
        r"page\s*(\d+)",
    ):
        m = re.search(pat, title, flags=re.IGNORECASE)
        if m:
            page = int(m.group(1))
            if page > 0:
                return page
    return None


def parse_pdf_name_from_title(title: str) -> str | None:
    text = title.strip()
    m = re.match(r"^(.+?\.pdf)\s*[-–—]", text, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.match(r"^(.+?\.pdf)\b", text, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"([^\\/:*?\"<>|]+\.pdf)", text, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None


def parse_word_name_from_title(title: str) -> str | None:
    text = title.strip()
    for ext in (".docx", ".doc"):
        m = re.match(rf"^(.+?{re.escape(ext)})\s*[-–—]", text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
        m = re.match(rf"^(.+?{re.escape(ext)})\b", text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
        m = re.search(rf"([^\\/:*?\"<>|]+{re.escape(ext)})", text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None


def is_wps_window_title(title: str) -> bool:
    lower = title.lower()
    return (
        any(k in lower for k in ("wps", "w365", "kingsoft"))
        or lower.endswith(PDF_SUFFIXES + WORD_SUFFIXES)
        or ".pdf" in lower
        or ".docx" in lower
        or ".doc" in lower
    )


def search_file_by_name(file_name: str, search_roots: list[Path]) -> tuple[Path | None, str]:
    key = file_name.casefold()
    cached = _FILE_SEARCH_CACHE.get(key)
    if cached and cached.exists():
        return cached, "search_cache"

    matches: list[Path] = []
    suffix = Path(file_name).suffix.lower()
    for root in search_roots:
        if not root.exists():
            continue
        try:
            for path in root.rglob(file_name):
                if path.is_file() and path.name.casefold() == key:
                    matches.append(path.resolve())
        except Exception:
            continue
        if not matches and suffix:
            try:
                for path in root.rglob(f"*{suffix}"):
                    if path.is_file() and path.name.casefold() == key:
                        matches.append(path.resolve())
            except Exception:
                continue

    if not matches:
        return None, "search_not_found"

    best = sorted(matches, key=lambda p: (p.stat().st_mtime, -len(str(p))), reverse=True)[0]
    _FILE_SEARCH_CACHE[key] = best
    return best, "search_root"


def search_pdf_by_name(pdf_name: str, search_roots: list[Path]) -> tuple[Path | None, str]:
    return search_file_by_name(pdf_name, search_roots)


def collect_wps_pdf_names() -> list[str]:
    names: list[str] = []
    seen: set[str] = set()

    def _add(title: str) -> None:
        name = parse_pdf_name_from_title(title)
        if not name:
            return
        key = name.casefold()
        if key in seen:
            return
        seen.add(key)
        names.append(name)

    def _enum(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd).strip()
        if not title or not is_wps_window_title(title):
            return True
        _add(title)
        for text in collect_child_window_texts(hwnd, limit=120):
            lower = text.lower()
            if ".pdf" in lower or "wps pdf" in lower:
                _add(text)
        return True

    try:
        win32gui.EnumWindows(_enum, None)
    except Exception:
        pass
    return names


def collect_wps_word_names() -> list[str]:
    names: list[str] = []
    seen: set[str] = set()

    def _add(title: str) -> None:
        name = parse_word_name_from_title(title)
        if not name:
            return
        key = name.casefold()
        if key in seen:
            return
        seen.add(key)
        names.append(name)

    def _enum(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd).strip()
        if not title or not is_wps_window_title(title):
            return True
        _add(title)
        for text in collect_child_window_texts(hwnd, limit=120):
            lower = text.lower()
            if ".docx" in lower or ".doc" in lower:
                _add(text)
        return True

    try:
        win32gui.EnumWindows(_enum, None)
    except Exception:
        pass
    return names


def resolve_pdf_path_from_title(
    title: str,
    app=None,
    search_roots: list[Path] | None = None,
) -> tuple[Path | None, str]:
    pdf_name = parse_pdf_name_from_title(title)
    if not pdf_name:
        return None, "title_no_pdf"

    candidates: list[Path] = []
    if app is not None:
        try:
            doc = app.ActiveDocument
            folder = com_get_str(doc, "Path")
            if folder:
                candidates.append((Path(folder) / pdf_name).resolve())
            full = normalize_path(com_get_str(doc, "FullName"))
            if full and full.name.lower() == pdf_name.lower():
                candidates.append(full)
        except Exception:
            pass

    for cand in candidates:
        if cand.exists():
            return cand, "foreground_title+ActiveDocument"

    roots = search_roots or [DEFAULT_SEARCH_ROOT]
    found, src = search_pdf_by_name(pdf_name, roots)
    if found is not None:
        return found, src
    return None, f"search_miss:{pdf_name}"


def resolve_pdf_path_from_wps_windows(
    app=None,
    search_roots: list[Path] | None = None,
) -> tuple[Path | None, str]:
    candidates: list[tuple[Path, str, int]] = []

    def score_title(text: str) -> int:
        lower = text.lower()
        if "wps pdf" in lower:
            return 3
        if lower.endswith(".pdf"):
            return 2
        if ".pdf" in lower:
            return 1
        return 0

    def _enum(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd).strip()
        if not title or not is_wps_window_title(title):
            return True
        texts = [title, *collect_child_window_texts(hwnd, limit=120)]
        for text in texts:
            if score_title(text) <= 0:
                continue
            path, src = resolve_pdf_path_from_title(text, app, search_roots)
            if path is not None:
                candidates.append((path, f"{src}:{text}", score_title(text)))
        return True

    try:
        win32gui.EnumWindows(_enum, None)
    except Exception:
        pass

    for pdf_name in collect_wps_pdf_names():
        found, src = search_pdf_by_name(pdf_name, search_roots or [DEFAULT_SEARCH_ROOT])
        if found is not None:
            candidates.append((found, f"{src}:{pdf_name}", 4))

    if candidates:
        candidates.sort(key=lambda item: (item[2], item[0].stat().st_mtime), reverse=True)
        best = candidates[0]
        return best[0], best[1]
    return None, "wps_pdf_not_found"


def resolve_word_path_from_title(
    title: str,
    app=None,
    search_roots: list[Path] | None = None,
) -> tuple[Path | None, str]:
    word_name = parse_word_name_from_title(title)
    if not word_name:
        return None, "title_no_word"

    candidates: list[Path] = []
    if app is not None:
        try:
            doc = app.ActiveDocument
            folder = com_get_str(doc, "Path")
            if folder:
                candidates.append((Path(folder) / word_name).resolve())
            full = normalize_document_path(
                com_get_str(doc, "FullName"),
                allowed_suffixes=WORD_SUFFIXES,
            )
            if full and full.name.lower() == word_name.lower():
                candidates.append(full)
        except Exception:
            pass

    for cand in candidates:
        if cand.exists():
            return cand, "foreground_title+ActiveDocument"

    roots = search_roots or [DEFAULT_SEARCH_ROOT]
    found, src = search_file_by_name(word_name, roots)
    if found is not None:
        return found, src
    return None, f"search_miss:{word_name}"


def resolve_word_path_from_wps_windows(
    app=None,
    search_roots: list[Path] | None = None,
) -> tuple[Path | None, str]:
    candidates: list[tuple[Path, str, int]] = []

    def score_title(text: str) -> int:
        lower = text.lower()
        if lower.endswith(".docx"):
            return 3
        if lower.endswith(".doc"):
            return 2
        if ".docx" in lower or ".doc" in lower:
            return 1
        return 0

    def _enum(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd).strip()
        if not title or not is_wps_window_title(title):
            return True
        texts = [title, *collect_child_window_texts(hwnd, limit=120)]
        for text in texts:
            if score_title(text) <= 0:
                continue
            path, src = resolve_word_path_from_title(text, app, search_roots)
            if path is not None:
                candidates.append((path, f"{src}:{text}", score_title(text)))
        return True

    try:
        win32gui.EnumWindows(_enum, None)
    except Exception:
        pass

    roots = search_roots or [DEFAULT_SEARCH_ROOT]
    for word_name in collect_wps_word_names():
        found, src = search_file_by_name(word_name, roots)
        if found is not None:
            candidates.append((found, f"{src}:{word_name}", 4))

    if candidates:
        candidates.sort(key=lambda item: (item[2], item[0].stat().st_mtime), reverse=True)
        best = candidates[0]
        return best[0], best[1]
    return None, "wps_word_not_found"


def resolve_page_from_foreground() -> tuple[int | None, str]:
    title = get_foreground_window_title()
    if title:
        page = parse_page_from_title(title)
        if page:
            return page, f"foreground_title:{title}"

    try:
        hwnd = win32gui.GetForegroundWindow()
        texts = collect_child_window_texts(hwnd)
        page = parse_page_from_ui_texts([title, *texts])
        if page:
            return page, "foreground_ui_text"
    except Exception:
        pass

    return None, "foreground_not_found"


def guess_pdf_from_window_title(app) -> Path | None:
    try:
        title = str(app.ActiveWindow.Caption)
    except Exception:
        return None
    if not title:
        return None
    # e.g. "报告.pdf - WPS PDF" / "报告.pdf - WPS Office"
    m = re.match(r"^(.+?\.pdf)\s*[-–—]", title, flags=re.IGNORECASE)
    if not m:
        m = re.search(r"([^\\/:*?\"<>|]+\.pdf)", title, flags=re.IGNORECASE)
    if not m:
        return None
    name = m.group(1).strip()
    candidates: list[Path] = []
    try:
        doc = app.ActiveDocument
        folder = Path(str(doc.Path))
        if folder.exists():
            candidates.append((folder / name).resolve())
    except Exception:
        pass
    try:
        full = normalize_path(str(app.ActiveDocument.FullName))
        if full:
            candidates.append(full)
    except Exception:
        pass
    for cand in candidates:
        if cand.exists():
            return cand
    return None


def resolve_pdf_path(app, active_pdf) -> tuple[Path | None, str]:
    probes: list[tuple[str, str | None]] = []

    if active_pdf is not None:
        for attr in (
            "FullName",
            "FileName",
            "FilePath",
            "Path",
            "Source",
            "DocumentPath",
            "DocPath",
        ):
            try:
                probes.append((f"ActivePDF.{attr}", com_get_str(active_pdf, attr)))
            except Exception:
                pass

    try:
        doc = app.ActiveDocument
        probes.append(("ActiveDocument.FullName", com_get_str(doc, "FullName")))
        probes.append(("ActiveDocument.Name", com_get_str(doc, "Name")))
        try:
            folder = com_get_str(doc, "Path")
            name = com_get_str(doc, "Name")
            if folder and name:
                probes.append(("ActiveDocument.Path+Name", str(Path(folder) / name)))
        except Exception:
            pass
    except Exception:
        pass

    for label, raw in probes:
        path = normalize_path(raw)
        if path is not None:
            return path, label

    guessed = guess_pdf_from_window_title(app)
    if guessed is not None:
        return guessed, "ActiveWindow.Caption"

    return None, "not_found"


def resolve_word_path_from_app(app) -> tuple[Path | None, str]:
    try:
        doc = app.ActiveDocument
        full = normalize_document_path(com_get_str(doc, "FullName"), allowed_suffixes=WORD_SUFFIXES)
        if full is not None:
            return full, "ActiveDocument.FullName"
        folder = com_get_str(doc, "Path")
        name = com_get_str(doc, "Name")
        if folder and name:
            cand = normalize_document_path(str(Path(folder) / name), allowed_suffixes=WORD_SUFFIXES)
            if cand is not None:
                return cand, "ActiveDocument.Path+Name"
    except Exception:
        pass
    return None, "word_not_found"


def resolve_current_page(app, active_pdf) -> tuple[int | None, str]:
    if active_pdf is not None:
        page = com_get_int(active_pdf, "CurrentPage", "currentPage")
        if page and page > 0:
            return page, "ActivePDF.CurrentPage"
        try:
            pages = active_pdf.ShowPages
            if pages:
                if isinstance(pages, (list, tuple)):
                    first = int(pages[0])
                else:
                    first = int(pages)
                if first > 0:
                    return first, "ActivePDF.ShowPages[0]"
        except Exception:
            pass
    return None, "not_found"


def save_image_blob(image_ref: str, out_png: Path) -> bool:
    ref = image_ref.strip().strip('"')
    src = normalize_path(ref)
    if src and src.exists():
        out_png.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, out_png)
        return True
    if Path(ref).exists():
        out_png.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ref, out_png)
        return True
    return False


def export_via_wps_render(active_pdf, page: int, out_png: Path) -> tuple[bool, str, float]:
    t0 = time.perf_counter()

    # Method 1: direct page render API
    for call in (
        lambda: active_pdf.GetPageRenderPicture(page),
        lambda: active_pdf.GetPageRenderPicture(Index=page),
        lambda: active_pdf.GetPageRenderPicture({"Index": page}),
    ):
        try:
            result = call()
            if isinstance(result, str) and save_image_blob(result, out_png):
                return True, "ActivePDF.GetPageRenderPicture", time.perf_counter() - t0
        except Exception:
            pass

    # Method 2: export current / single page as image
    enum_candidates = []
    try:
        enum_obj = active_pdf.Enum
        enum_candidates.append(enum_obj)
    except Exception:
        pass
    try:
        enum_candidates.append(getattr(active_pdf, "Application", None).Enum)
    except Exception:
        pass

    for enum_obj in enum_candidates:
        if enum_obj is None:
            continue
        for range_name in ("ImgTypeCurrent", "TypeCurrent", "Current"):
            for fmt_name in ("TypeIMG", "ImgTypePNG", "PNG"):
                try:
                    fixed = getattr(enum_obj.FixedFormatType, fmt_name, None)
                    range_type = getattr(enum_obj.RangeType, range_name, None)
                    img_format = getattr(enum_obj.ExportImgFormatType, "ImgTypePNG", None)
                    kwargs = {}
                    if fixed is not None:
                        kwargs["FixedFormatType"] = fixed
                    if img_format is not None:
                        kwargs["ImgFormat"] = img_format
                    if range_type is not None:
                        kwargs["RangeType"] = range_type
                    kwargs["Dpi"] = int(96 * ZOOM)
                    result = active_pdf.ExportAsFixedFormat(**kwargs)
                    if isinstance(result, str) and save_image_blob(result, out_png):
                        return True, "ActivePDF.ExportAsFixedFormat", time.perf_counter() - t0
                except Exception:
                    pass

    return False, "wps_render_failed", time.perf_counter() - t0


def export_via_wps_word(app, page: int, out_png: Path) -> tuple[bool, str, float]:
    t0 = time.perf_counter()
    if Image is None:
        return False, "pillow_required", time.perf_counter() - t0

    try:
        doc = app.ActiveDocument
    except Exception as exc:
        return False, f"no_active_document:{exc}", time.perf_counter() - t0

    page_obj = None
    try:
        page_obj = doc.ActiveWindow.ActivePane.Pages(page)
    except Exception:
        pass

    if page_obj is None:
        try:
            for wi in range(1, doc.Windows.Count + 1):
                win = doc.Windows(wi)
                for pi in range(1, win.Panes.Count + 1):
                    try:
                        page_obj = win.Panes(pi).Pages(page)
                        break
                    except Exception:
                        continue
                if page_obj is not None:
                    break
        except Exception:
            pass

    if page_obj is None:
        return False, "word_page_not_found", time.perf_counter() - t0

    try:
        data = bytes(page_obj.EnhMetaFileBits)
    except Exception as exc:
        return False, f"enhmetafile_failed:{exc}", time.perf_counter() - t0

    tmp_emf = out_png.with_suffix(".emf")
    try:
        out_png.parent.mkdir(parents=True, exist_ok=True)
        tmp_emf.write_bytes(data)
        Image.open(str(tmp_emf)).save(str(out_png), "PNG")
        return True, f"WPS-Word(page={page})", time.perf_counter() - t0
    except Exception as exc:
        return False, f"word_export_failed:{exc}", time.perf_counter() - t0
    finally:
        try:
            tmp_emf.unlink(missing_ok=True)
        except Exception:
            pass


def export_via_pymupdf(pdf_path: Path, page: int, out_png: Path) -> tuple[bool, str, float]:
    t0 = time.perf_counter()
    doc = fitz.open(str(pdf_path))
    try:
        idx = max(0, min(page - 1, doc.page_count - 1))
        pix = doc[idx].get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM))
        out_png.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(out_png))
        return True, f"PyMuPDF(page={page}, idx={idx})", time.perf_counter() - t0
    finally:
        doc.close()


def probe_wps() -> int:
    pythoncom.CoInitialize()
    try:
        title = get_foreground_window_title()
        log(f"[PROBE] 前台窗口标题: {title or '(空)'}")
        try:
            hwnd = win32gui.GetForegroundWindow()
            texts = collect_child_window_texts(hwnd, limit=80)
            if texts:
                log("[PROBE] 前台子控件文本(前20条):")
                for text in texts[:20]:
                    log(f"        {text}")
        except Exception as exc:
            log(f"[PROBE] 读取子控件失败: {exc}")

        for prog_id in WPS_PROG_IDS:
            try:
                app = win32com.client.GetActiveObject(prog_id)
            except Exception:
                continue
            log(f"[PROBE] 已连接: {prog_id}")
            for name in (
                "Name",
                "Version",
                "ActiveWindow",
                "ActiveDocument",
                "ActivePDF",
                "Documents",
            ):
                try:
                    val = getattr(app, name)
                    if name == "Documents":
                        log(f"        Documents.Count = {val.Count}")
                    else:
                        log(f"        {name} = {val}")
                except Exception as exc:
                    log(f"        {name} ERR = {exc}")
            active_pdf = com_get_active_pdf(app)
            if active_pdf is not None:
                page = com_get_int(active_pdf, "CurrentPage")
                log(f"        ActivePDF.CurrentPage = {page}")
            pdf_path, src = resolve_pdf_path(app, active_pdf)
            log(f"        PDF路径探测 = {pdf_path} ({src})")
        return 0
    finally:
        pythoncom.CoUninitialize()


def resolve_page_from_wps_windows() -> tuple[int | None, str]:
    """WPS PDF desktop viewer usually does not expose page text."""
    return None, "manual_page_required"


def detect_wps_context(
    pdf_override: Path | None = None,
    page_override: int | None = None,
    search_roots: list[Path] | None = None,
    quiet: bool = False,
) -> tuple[object | None, object | None, str, Path | None, str, int | None, str, str]:
    """Return WPS app, active_pdf, prog_id, pdf_path, path_src, page, page_src, fg_title."""
    app, prog_id = connect_wps(quiet=quiet)
    active_pdf = com_get_active_pdf(app)
    page, page_src = resolve_current_page(app, active_pdf)
    fg_title = get_foreground_window_title()
    roots = [DEFAULT_SEARCH_ROOT] if search_roots is None else search_roots

    pdf_path: Path | None = None
    path_src = "not_found"

    pdf_com, pdf_com_src = resolve_pdf_path(app, active_pdf)
    if pdf_com is not None:
        pdf_path, path_src = pdf_com, pdf_com_src

    if pdf_path is None:
        word_com, word_com_src = resolve_word_path_from_app(app)
        if word_com is not None:
            pdf_path, path_src = word_com, word_com_src

    if pdf_path is None and roots:
        pdf_win, pdf_win_src = resolve_pdf_path_from_wps_windows(app, roots)
        if pdf_win is not None:
            pdf_path, path_src = pdf_win, pdf_win_src

    if pdf_path is None and roots:
        word_win, word_win_src = resolve_word_path_from_wps_windows(app, roots)
        if word_win is not None:
            pdf_path, path_src = word_win, word_win_src

    if pdf_path is None and pdf_override is not None:
        pdf_path = pdf_override.resolve()
        path_src = "manual_document"

    if page is None:
        page, page_src = resolve_page_from_wps_windows()
    if page_override is not None:
        page = page_override
        page_src = "manual_page"

    return app, active_pdf, prog_id, pdf_path, path_src, page, page_src, fg_title


def capture_core(
    out_dir: Path,
    pdf_override: Path | None = None,
    page_override: int | None = None,
    search_roots: list[Path] | None = None,
    quiet: bool = False,
) -> CaptureResult:
    total_t0 = time.perf_counter()
    pythoncom.CoInitialize()
    try:
        app, active_pdf, prog_id, pdf_path, path_src, page, page_src, fg_title = detect_wps_context(
            pdf_override,
            page_override,
            search_roots=search_roots,
            quiet=quiet,
        )

        if not quiet:
            log(f"[INFO] COM ProgID     : {prog_id}")
            log(f"[INFO] 前台窗口       : {fg_title or '(空)'}")
            log(f"[INFO] 当前页         : {page} ({page_src})")
            log(f"[INFO] 文档路径       : {pdf_path} ({path_src})")

        if page is None:
            msg = "请在小窗口「页码」里填写 WPS 当前页（例如 12），再点截当前页。"
            if not quiet:
                log(f"[ERROR] {msg}")
            return CaptureResult(
                False,
                msg,
                elapsed=time.perf_counter() - total_t0,
                pdf_path=pdf_path,
            )

        kind = document_kind(pdf_path)
        out_png = build_output_png(out_dir, pdf_path, page)

        if kind == "word":
            active_word, _ = resolve_word_path_from_app(app)
            if active_word is None:
                msg = "未检测到 WPS 中打开的 Word 文档。请先在 WPS 打开 Word 文件。"
                if not quiet:
                    log(f"[ERROR] {msg}")
                return CaptureResult(
                    False,
                    msg,
                    elapsed=time.perf_counter() - total_t0,
                    page=page,
                    pdf_path=pdf_path,
                )
            if pdf_path is None or pdf_path.resolve() != active_word.resolve():
                pdf_path = active_word
                out_png = build_output_png(out_dir, pdf_path, page)
            ok, method, elapsed = export_via_wps_word(app, page, out_png)
            if ok:
                msg = f"Word 导出成功 ({elapsed:.2f}s)  第{page}页"
                if not quiet:
                    log(f"[OK] {msg} ({method})")
                    log(f"[OK] 已保存: {out_png}")
                return CaptureResult(
                    True,
                    msg,
                    out_png=out_png,
                    elapsed=time.perf_counter() - total_t0,
                    page=page,
                    pdf_path=pdf_path,
                )
            msg = f"Word 页面导出失败: {method}"
            if not quiet:
                log(f"[ERROR] {msg}")
            return CaptureResult(
                False,
                msg,
                elapsed=time.perf_counter() - total_t0,
                page=page,
                pdf_path=pdf_path,
            )

        if active_pdf is not None:
            ok, method, elapsed = export_via_wps_render(active_pdf, page, out_png)
            if ok:
                msg = f"WPS 导出成功 ({elapsed:.2f}s)"
                if not quiet:
                    log(f"[OK] {msg}")
                    log(f"[OK] 已保存: {out_png}")
                return CaptureResult(
                    True,
                    msg,
                    out_png=out_png,
                    elapsed=time.perf_counter() - total_t0,
                    page=page,
                    pdf_path=pdf_path,
                )
            if not quiet:
                log(f"[WARN] WPS 直接导出失败 ({elapsed:.2f}s)，改用 PyMuPDF。")

        if pdf_path is None or not pdf_path.exists():
            pdf_names = collect_wps_pdf_names()
            word_names = collect_wps_word_names()
            if pdf_names:
                msg = (
                    f"已识别 WPS 打开的 PDF 文件名，但在搜索目录未找到：{pdf_names[0]}。"
                    f"请检查「搜索目录」或手动浏览选择文件。"
                )
            elif word_names:
                msg = (
                    f"已识别 WPS 打开的 Word 文件名，但在搜索目录未找到：{word_names[0]}。"
                    f"请检查「搜索目录」或手动浏览选择文件。"
                )
            else:
                msg = "未找到文档。请先在 WPS 打开 PDF/Word，或手动浏览选择文件。"
            if not quiet:
                log(f"[ERROR] {msg}")
            return CaptureResult(
                False,
                msg,
                elapsed=time.perf_counter() - total_t0,
                page=page,
                pdf_path=pdf_path,
            )

        ok, method, elapsed = export_via_pymupdf(pdf_path, page, out_png)
        if not ok:
            msg = "PyMuPDF 导出失败。"
            if not quiet:
                log(f"[ERROR] {msg}")
            return CaptureResult(False, msg, elapsed=time.perf_counter() - total_t0)

        msg = f"导出成功 ({elapsed:.2f}s)  第{page}页"
        if not quiet:
            log(f"[OK] PyMuPDF 导出成功 ({elapsed:.2f}s, {method})")
            log(f"[OK] 已保存: {out_png}")
            log(f"[OK] 总耗时: {time.perf_counter() - total_t0:.2f}s")
        return CaptureResult(
            True,
            msg,
            out_png=out_png,
            elapsed=time.perf_counter() - total_t0,
            page=page,
            pdf_path=pdf_path,
        )
    except Exception as exc:
        msg = f"截图失败: {exc}"
        if not quiet:
            log(f"[ERROR] {msg}")
        return CaptureResult(False, msg, elapsed=time.perf_counter() - total_t0)
    finally:
        pythoncom.CoUninitialize()


def refresh_detection(
    pdf_override: Path | None = None,
    page_override: int | None = None,
    search_roots: list[Path] | None = None,
) -> tuple[Path | None, int | None, str]:
    pythoncom.CoInitialize()
    try:
        _, _, _, pdf_path, path_src, page, page_src, _ = detect_wps_context(
            pdf_override,
            page_override,
            search_roots=search_roots,
            quiet=True,
        )
        bits = []
        kind = document_kind(pdf_path)
        if pdf_path:
            label = "Word" if kind == "word" else "PDF"
            bits.append(f"{label}: {pdf_path.name}")
        else:
            bits.append("文档: 未找到")
        if page_override is not None:
            bits.append(f"页码: {page_override}")
        else:
            bits.append("页码: 请手动填写")
        bits.append(f"({path_src})")
        return pdf_path, page_override, "  |  ".join(bits)
    except Exception as exc:
        return pdf_override, page_override, f"刷新失败: {exc}"
    finally:
        pythoncom.CoUninitialize()


def capture_once(
    out_dir: Path,
    pdf_override: Path | None,
    page_override: int | None,
    search_roots: list[Path] | None = None,
) -> int:
    result = capture_core(
        out_dir,
        pdf_override,
        page_override,
        search_roots=search_roots,
        quiet=False,
    )
    return 0 if result.ok else 1


def run_gui(out_dir: Path, pdf_override: Path | None, page_override: int | None) -> int:
    import ctypes
    import tkinter as tk
    from tkinter import filedialog, ttk

    cfg = AppConfig.load()
    if page_override is None:
        page_override = cfg.last_page

    root = tk.Tk()
    root.title("PDF 截图助手")
    width, height = 500, 430
    root.resizable(False, False)
    root.attributes("-topmost", True)

    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    pos_x = max(20, screen_w - width - 40)
    pos_y = max(20, screen_h - height - 120)
    root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

    def bring_to_front() -> None:
        root.deiconify()
        root.lift()
        root.attributes("-topmost", True)
        root.focus_force()
        try:
            hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
            ctypes.windll.user32.SetForegroundWindow(hwnd)
        except Exception:
            pass

    state = {"busy": False}

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)

    status_var = tk.StringVar(value="● 运行中")
    detect_var = tk.StringVar(value="正在检测 WPS / PDF …")
    result_var = tk.StringVar(value="1) WPS 打开 PDF/Word  2) 填写页码  3) 点绿色按钮截图")

    ttk.Label(
        main,
        text="【截图工具窗口】保持本窗口不要关",
        font=("Microsoft YaHei UI", 10, "bold"),
        foreground="#1565c0",
    ).pack(anchor="w")
    ttk.Label(main, textvariable=status_var, font=("Microsoft YaHei UI", 11, "bold")).pack(
        anchor="w", pady=(6, 0)
    )

    ttk.Label(main, textvariable=detect_var, wraplength=460, foreground="#444").pack(
        anchor="w", pady=(4, 4)
    )

    search_row = ttk.Frame(main)
    search_row.pack(fill="x", pady=(0, 2))
    search_var = tk.StringVar(value=str(cfg.search_root))
    ttk.Label(search_row, text="搜索目录:").pack(side="left")
    ttk.Entry(search_row, textvariable=search_var).pack(
        side="left", fill="x", expand=True, padx=(6, 6)
    )

    def browse_search_root() -> None:
        path = filedialog.askdirectory(title="选择 PDF 搜索根目录")
        if path:
            search_var.set(path)

    ttk.Button(search_row, text="浏览…", command=browse_search_root, width=8).pack(side="left")

    search_warn_var = tk.StringVar(value="")
    ttk.Label(main, textvariable=search_warn_var, wraplength=460, foreground="#c62828").pack(
        anchor="w", pady=(0, 6)
    )

    path_row = ttk.Frame(main)
    path_row.pack(fill="x", pady=(0, 6))
    pdf_var = tk.StringVar(value="")
    ttk.Label(path_row, text="文档:").pack(side="left")
    pdf_entry = ttk.Entry(path_row, textvariable=pdf_var)
    pdf_entry.pack(side="left", fill="x", expand=True, padx=(6, 6))
    ttk.Label(path_row, text="(自动识别)", foreground="#666").pack(side="right", padx=(0, 4))

    def browse_pdf() -> None:
        path = filedialog.askopenfilename(
            title="选择 PDF 或 Word 文件",
            filetypes=[
                ("PDF / Word", "*.pdf;*.docx;*.doc"),
                ("PDF 文件", "*.pdf"),
                ("Word 文件", "*.docx;*.doc"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            pdf_var.set(path)

    ttk.Button(path_row, text="浏览…", command=browse_pdf, width=8).pack(side="left")

    page_row = ttk.Frame(main)
    page_row.pack(fill="x", pady=(0, 6))
    ttk.Label(page_row, text="页码:").pack(side="left")
    page_var = tk.StringVar(value=str(page_override) if page_override else "1")
    ttk.Entry(page_row, textvariable=page_var, width=8).pack(side="left", padx=(6, 8))
    ttk.Label(page_row, text="(单页，点绿色按钮)").pack(side="left")

    batch_row = ttk.Frame(main)
    batch_row.pack(fill="x", pady=(0, 10))
    ttk.Label(batch_row, text="批量:").pack(side="left")
    batch_from_var = tk.StringVar(value="")
    batch_to_var = tk.StringVar(value="")
    ttk.Label(batch_row, text="从").pack(side="left", padx=(6, 4))
    ttk.Entry(batch_row, textvariable=batch_from_var, width=6).pack(side="left")
    ttk.Label(batch_row, text="到").pack(side="left", padx=(4, 4))
    ttk.Entry(batch_row, textvariable=batch_to_var, width=6).pack(side="left", padx=(0, 8))
    batch_btn = ttk.Button(batch_row, text="批量截图", width=10)
    batch_btn.pack(side="left")

    btn_row = ttk.Frame(main)
    btn_row.pack(fill="x", pady=(0, 8))

    capture_btn = tk.Button(
        btn_row,
        text="截当前页",
        width=14,
        height=2,
        font=("Microsoft YaHei UI", 12, "bold"),
        bg="#2e7d32",
        fg="white",
        activebackground="#1b5e20",
        activeforeground="white",
        relief="raised",
        cursor="hand2",
    )
    capture_btn.pack(side="left")

    def set_busy(busy: bool, status: str) -> None:
        state["busy"] = busy
        status_var.set(status)
        btn_state = "disabled" if busy else "normal"
        capture_btn.configure(state=btn_state)
        batch_btn.configure(state=btn_state)

    def get_batch_range() -> tuple[int, int] | None:
        from_text = batch_from_var.get().strip()
        to_text = batch_to_var.get().strip()
        if not from_text or not to_text:
            return None
        try:
            start = int(from_text)
            end = int(to_text)
            if start > 0 and end > 0 and start <= end:
                return start, end
        except ValueError:
            pass
        return None

    def get_search_roots() -> list[Path]:
        text = search_var.get().strip().strip('"')
        if not text:
            return [DEFAULT_SEARCH_ROOT]
        path = Path(text)
        return [path] if path.exists() and path.is_dir() else []

    def get_search_root_warning() -> str | None:
        text = search_var.get().strip().strip('"')
        if not text:
            return None
        path = Path(text)
        if path.exists() and path.is_dir():
            return None
        return f"搜索目录不存在: {text}"

    def get_manual_pdf() -> Path | None:
        text = pdf_var.get().strip().strip('"')
        if not text:
            return None
        path = Path(text)
        return path if path.exists() and path.suffix.lower() in DOCUMENT_SUFFIXES else None

    def get_manual_page() -> int | None:
        text = page_var.get().strip()
        if not text:
            return None
        try:
            page = int(text)
            return page if page > 0 else None
        except ValueError:
            return None

    def save_config(pdf_path: Path | None, page: int | None) -> None:
        text = search_var.get().strip().strip('"')
        if text:
            cfg.search_root = Path(text)
        roots = get_search_roots()
        if not text and roots:
            cfg.search_root = roots[0]
        if pdf_path:
            cfg.last_pdf = pdf_path
        if page:
            cfg.last_page = page
        cfg.save()

    def update_search_warn() -> None:
        search_warn_var.set(get_search_root_warning() or "")

    def apply_refresh_result(pdf_path: Path | None, _page: int | None, text: str) -> None:
        update_search_warn()
        detect_var.set(text)
        pdf_var.set(str(pdf_path) if pdf_path else "")

    def do_refresh() -> None:
        if state["busy"]:
            return

        def worker() -> None:
            pdf_path, page, text = refresh_detection(
                None,
                get_manual_page(),
                search_roots=get_search_roots() or None,
            )
            root.after(0, lambda: apply_refresh_result(pdf_path, page, text))

        threading.Thread(target=worker, daemon=True).start()

    def do_capture() -> None:
        if state["busy"]:
            return
        manual_page = get_manual_page()
        if manual_page is None:
            result_var.set("请先在「页码」里填写当前页，再点截当前页。")
            return
        set_busy(True, "● 截图中…")
        result_var.set("正在导出，请稍候…")

        def worker() -> None:
            result = capture_core(
                out_dir,
                get_manual_pdf(),
                manual_page,
                search_roots=get_search_roots() or None,
                quiet=True,
            )

            def finish() -> None:
                if result.ok:
                    set_busy(False, "● 运行中")
                    if result.pdf_path:
                        pdf_var.set(str(result.pdf_path))
                    save_config(result.pdf_path, manual_page)
                    result_var.set(
                        f"{result.message}\n已保存: {result.out_png}\n耗时: {result.elapsed:.2f}s"
                    )
                else:
                    set_busy(False, "● 运行中（上次失败）")
                    result_var.set(result.message)
                    save_config(result.pdf_path or get_manual_pdf(), manual_page)
                do_refresh()

            root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def do_batch_capture() -> None:
        if state["busy"]:
            return
        batch_range = get_batch_range()
        if batch_range is None:
            result_var.set("请填写批量「从」和「到」页码，且起始页不能大于结束页。")
            return
        start_page, end_page = batch_range
        total = end_page - start_page + 1
        set_busy(True, f"● 批量截图中 0/{total}")
        result_var.set(f"正在批量导出第 {start_page}–{end_page} 页…")

        def worker() -> None:
            ok_count = 0
            last_pdf: Path | None = None
            errors: list[str] = []
            t0 = time.perf_counter()

            for index, page in enumerate(range(start_page, end_page + 1), 1):
                result = capture_core(
                    out_dir,
                    get_manual_pdf(),
                    page,
                    search_roots=get_search_roots() or None,
                    quiet=True,
                )
                if result.ok:
                    ok_count += 1
                    if result.pdf_path:
                        last_pdf = result.pdf_path
                else:
                    errors.append(f"第{page}页: {result.message}")

                root.after(
                    0,
                    lambda i=index, n=ok_count: status_var.set(f"● 批量截图中 {i}/{total}（成功 {n}）"),
                )

            elapsed = time.perf_counter() - t0

            def finish() -> None:
                if ok_count == total:
                    status_var.set("● 运行中")
                    summary = f"批量完成 {ok_count}/{total} 页，耗时 {elapsed:.1f}s"
                elif ok_count > 0:
                    status_var.set("● 运行中（部分失败）")
                    summary = f"批量完成 {ok_count}/{total} 页，耗时 {elapsed:.1f}s"
                else:
                    status_var.set("● 运行中（上次失败）")
                    summary = f"批量失败 0/{total} 页，耗时 {elapsed:.1f}s"

                if errors:
                    summary += "\n失败:\n" + "\n".join(errors[:5])
                    if len(errors) > 5:
                        summary += f"\n…另有 {len(errors) - 5} 页失败"

                if last_pdf:
                    pdf_var.set(str(last_pdf))
                save_config(last_pdf, start_page)
                result_var.set(summary)
                set_busy(False, status_var.get())
                do_refresh()

            root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    search_var.trace_add("write", lambda *_: update_search_warn())

    capture_btn.configure(command=do_capture)
    batch_btn.configure(command=do_batch_capture)
    ttk.Button(btn_row, text="刷新识别", command=do_refresh, width=10).pack(side="left", padx=(8, 0))

    def open_out_dir() -> None:
        out_dir.mkdir(parents=True, exist_ok=True)
        import os

        os.startfile(out_dir)

    ttk.Button(btn_row, text="打开输出文件夹", command=open_out_dir).pack(side="right")

    ttk.Label(main, textvariable=result_var, wraplength=460, foreground="#006400").pack(
        anchor="w", pady=(8, 0)
    )
    ttk.Label(main, text=f"输出目录: {out_dir}", wraplength=460, foreground="#666").pack(
        anchor="w", pady=(8, 0)
    )

    root.after(100, bring_to_front)
    root.after(800, bring_to_front)
    root.after(100, update_search_warn)
    root.after(300, do_refresh)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()
    return 0


def run_hotkey(
    out_dir: Path,
    hotkey: str,
    pdf_override: Path | None,
    page_override: int | None,
) -> int:
    try:
        import keyboard
    except ImportError:
        log("热键模式需要 keyboard 库: pip install keyboard")
        return 1

    log(f"[INFO] 热键监听中: {hotkey}")
    log("[INFO] 在 WPS 中翻到目标页后按热键；Ctrl+C 退出。")

    def on_hotkey():
        log("\n--- 热键触发 ---")
        capture_once(out_dir, pdf_override, page_override)

    keyboard.add_hotkey(hotkey, on_hotkey)
    keyboard.wait()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="WPS PDF 当前页截图测试")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUT,
        help=f"输出目录（默认: {DEFAULT_OUT}）",
    )
    parser.add_argument(
        "--hotkey",
        action="store_true",
        help="启用热键监听（默认 Ctrl+Shift+P）",
    )
    parser.add_argument(
        "--bind",
        default="ctrl+shift+p",
        help="热键组合，默认 ctrl+shift+p",
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="手动指定 PDF 完整路径（自动识别失败时使用）",
    )
    parser.add_argument(
        "--page",
        type=int,
        default=None,
        help="手动指定页码（1 起；自动识别失败时使用）",
    )
    parser.add_argument(
        "--probe",
        action="store_true",
        help="只诊断 WPS / 前台窗口信息，不导出图片",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="打开置顶小窗口（默认模式）",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="命令行单次截图，不打开窗口",
    )
    parser.add_argument(
        "--search-root",
        type=Path,
        default=DEFAULT_SEARCH_ROOT,
        help=f"PDF 文件名搜索根目录（默认: {DEFAULT_SEARCH_ROOT}）",
    )
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    search_roots = [args.search_root]

    if args.probe:
        return probe_wps()
    if args.hotkey:
        return run_hotkey(args.output, args.bind, args.pdf, args.page)
    if args.once:
        return capture_once(args.output, args.pdf, args.page, search_roots=search_roots)
    return run_gui(args.output, args.pdf, args.page)


if __name__ == "__main__":
    raise SystemExit(main())
