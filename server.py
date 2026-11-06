#!/usr/bin/env python3
"""
Универсальный MCP-сервер для работы с PDF и Office-файлами.
PDF: текст, таблицы, чертежи, векторная геометрия, слои, измерения, DXF, рендеринг.
Office (DOCX, XLSX, PPTX): текст, таблицы, встроенные картинки, метаданные.
CAD-инструменты (слои, геометрия, измерения, DXF) работают ТОЛЬКО с PDF.
Зависимости: mcp, pymupdf, pdfplumber, Pillow, python-docx, openpyxl, python-pptx, ezdxf (опц.)
"""
import sys
# КРИТИЧНО для Windows: переключение потоков в UTF-8 до начала работы MCP
for _stream in (sys.stdin, sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

import base64
import io
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pymupdf as fitz
import pdfplumber
from mcp.server.fastmcp import FastMCP
from mcp.types import ImageContent, TextContent
from PIL import Image

# Office-библиотеки — опциональные, чтобы сервер падал gracefully
try:
    import docx as _docx
    _HAS_DOCX = True
except ImportError:
    _HAS_DOCX = False

try:
    import openpyxl as _openpyxl
    _HAS_XLSX = True
except ImportError:
    _HAS_XLSX = False

try:
    import pptx as _pptx
    from pptx.util import Inches, Pt as _PptPt
    _HAS_PPTX = True
except ImportError:
    _HAS_PPTX = False

try:
    import ezdxf
    import logging as _logging
    _logging.getLogger("ezdxf").setLevel(_logging.WARNING)
    _HAS_EZDXF = True
except ImportError:
    _HAS_EZDXF = False

mcp = FastMCP("mcp-server-for-docs")

# ---------------------------------------------------------------------------
# Настройки / безопасность
# ---------------------------------------------------------------------------
_default_roots = [str(Path.home())]
_env_roots = os.environ.get("PDF_MCP_ROOT")
if _env_roots:
    sep = ";" if ";" in _env_roots else os.pathsep
    ALLOWED_ROOTS = [Path(p).expanduser().resolve() for p in _env_roots.split(sep) if p.strip()]
else:
    ALLOWED_ROOTS = [Path(p).expanduser().resolve() for p in _default_roots]

MAX_RENDER_DPI = 400
DEFAULT_DPI = 150

SUPPORTED_EXTS = {".pdf", ".docx", ".xlsx", ".pptx", ".txt", ".md"}

# Состояние сессии
_active_file: Path | None = None
_active_file_type: str | None = None  # "pdf", "docx", "xlsx", "pptx", "text"
_scale_cache: dict[str, tuple[float, str]] = {}


def _check_writable(path: Path) -> Path:
    p = path.expanduser().resolve()
    if ALLOWED_ROOTS and not any(str(p).startswith(str(root)) for root in ALLOWED_ROOTS):
        raise PermissionError(f"Путь {p} вне разрешённых каталогов.")
    return p


_CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")


def _sanitize_text(s: str | None) -> str:
    if not s:
        return s or ""
    s = s.encode("utf-8", "replace").decode("utf-8")
    try:
        s = s.encode("utf-16", "surrogatepass").decode("utf-16", "replace")
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    return _CONTROL_CHARS_RE.sub("", s)


def _sanitize_json_value(value: Any) -> Any:
    if isinstance(value, str):
        return _sanitize_text(value)
    if isinstance(value, dict):
        return {k: _sanitize_json_value(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_sanitize_json_value(v) for v in value]
    return value


def _point_to_list(pt: "fitz.Point") -> list[float]:
    return [round(pt.x, 2), round(pt.y, 2)]


def _serialize_drawing_items(items: list) -> list[dict]:
    out = []
    for it in items:
        op = it[0]
        if op == "l":
            out.append({"op": "line", "points": [_point_to_list(it[1]), _point_to_list(it[2])]})
        elif op == "re":
            r = it[1]
            out.append({"op": "rect", "rect": [round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2)]})
        elif op == "c":
            out.append({"op": "curve", "points": [_point_to_list(p) for p in it[1:]]})
        elif op == "qu":
            q = it[1]
            out.append({"op": "quad", "points": [_point_to_list(q.ul), _point_to_list(q.ur), _point_to_list(q.lr), _point_to_list(q.ll)]})
        else:
            out.append({"op": str(op), "raw": str(it)})
    return out


def _detect_file_type(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".pdf": return "pdf"
    if ext == ".docx": return "docx"
    if ext == ".xlsx": return "xlsx"
    if ext == ".pptx": return "pptx"
    if ext in {".txt", ".md"}: return "text"
    return "unknown"


def _resolve_file(path: str | None) -> Path:
    """Универсальный резолвер для ЛЮБОГО поддерживаемого файла."""
    global _active_file, _active_file_type
    if not path:
        if _active_file is not None:
            return _active_file
        raise ValueError(
            "Путь к файлу не указан, и нет закреплённого файла. "
            "Сначала вызови open_document(path) или open_pdf(path)."
        )
    
    # Нормализуем путь: убираем кавычки, пробелы по краям
    path = path.strip().strip('"').strip("'")
    raw = Path(path).expanduser()
    
    # Если путь абсолютный — используем напрямую
    if raw.is_absolute():
        p = raw.resolve()
    else:
        # Относительный путь: пробуем несколько вариантов
        candidates = []
        
        # 1. Относительно текущей рабочей директории
        cwd_candidate = (Path.cwd() / raw).resolve()
        if cwd_candidate.exists():
            candidates.append(cwd_candidate)
        
        # 2. Относительно домашней директории
        home_candidate = (Path.home() / raw).resolve()
        if home_candidate.exists() and home_candidate not in candidates:
            candidates.append(home_candidate)
        
        # 3. Поиск в разрешённых корневых директориях
        for root in ALLOWED_ROOTS:
            if not root.exists():
                continue
            # Прямое объединение путей
            direct_candidate = (root / raw).resolve()
            if direct_candidate.exists() and direct_candidate not in candidates:
                candidates.append(direct_candidate)
            # Поиск по имени файла (если указан только имя)
            if raw.parent == Path("."):
                for match in root.rglob(raw.name):
                    if match.exists() and match not in candidates:
                        candidates.append(match)
        
        if len(candidates) == 1:
            p = candidates[0]
        elif len(candidates) > 1:
            listing = "\n".join(f"  - {m}" for m in candidates)
            raise FileNotFoundError(
                f"Найдено несколько файлов по пути '{path}'. Уточни полный путь:\n{listing}"
            )
        else:
            # Файл не найден — даём полезную ошибку
            raise FileNotFoundError(
                f"Файл '{path}' не найден.\n"
                f"Проверь:\n"
                f"  1. Путь указан правильно (абсолютный или относительный)\n"
                f"  2. Файл существует\n"
                f"  3. Если путь относительный, попробуй указать абсолютный путь\n"
                f"Текущая рабочая директория: {Path.cwd()}"
            )

    if not p.exists():
        raise FileNotFoundError(f"Файл не найден: {p}")
    
    # Проверка безопасности: файл должен быть в разрешённых директориях
    # Но если ALLOWED_ROOTS пустой или содержит только домашнюю директорию,
    # разрешаем доступ к любому существующему файлу (для удобства)
    if ALLOWED_ROOTS and len(ALLOWED_ROOTS) > 0:
        # Если есть явные ограничения через PDF_MCP_ROOT — проверяем строго
        if _env_roots:
            if not any(str(p).startswith(str(root)) for root in ALLOWED_ROOTS):
                raise PermissionError(
                    f"Путь {p} вне разрешённых каталогов.\n"
                    f"Разрешённые каталоги: {', '.join(str(r) for r in ALLOWED_ROOTS)}\n"
                    f"Задай PDF_MCP_ROOT для расширения доступа."
                )
        # Иначе (по умолчанию) — разрешаем доступ ко всем файлам
    
    if p.suffix.lower() not in SUPPORTED_EXTS:
        raise ValueError(
            f"Неподдерживаемый формат: {p.suffix}.\n"
            f"Допустимы: {', '.join(sorted(SUPPORTED_EXTS))}"
        )
    
    _active_file = p
    _active_file_type = _detect_file_type(p)
    return p


def _ensure_pdf(path: str | None) -> Path:
    """Строгий резолвер: возвращает путь ТОЛЬКО если это PDF. Для CAD-инструментов."""
    p = _resolve_file(path)
    if _active_file_type != "pdf":
        raise ValueError(
            f"Этот инструмент работает ТОЛЬКО с PDF-файлами (чертежами). "
            f"Сейчас открыт: {p.name} (тип: {_active_file_type}). "
            f"Открой PDF через open_pdf(path) и повтори."
        )
    return p


def _parse_page_range(page_range: str | None, total_pages: int) -> list[int]:
    if not page_range:
        return list(range(total_pages))
    pages: set[int] = set()
    for part in page_range.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            pages.update(range(int(a) - 1, int(b)))
        else:
            pages.add(int(part) - 1)
    return sorted(p for p in pages if 0 <= p < total_pages)


# ---------------------------------------------------------------------------
# Вспомогательные функции для Office
# ---------------------------------------------------------------------------
def _extract_docx_images(path: Path, min_size_px: int) -> list[tuple[str, bytes, int, int]]:
    """Извлекает картинки из DOCX (они лежат в word/media/ внутри ZIP)."""
    out = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.startswith("word/media/"):
                continue
            data = z.read(name)
            try:
                img = Image.open(io.BytesIO(data))
            except Exception:
                continue
            if img.width < min_size_px or img.height < min_size_px:
                continue
            out.append((name, data, img.width, img.height))
    return out


def _extract_xlsx_images(path: Path, min_size_px: int) -> list[tuple[str, bytes, int, int]]:
    """Извлекает картинки из XLSX (xl/media/)."""
    out = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.startswith("xl/media/"):
                continue
            data = z.read(name)
            try:
                img = Image.open(io.BytesIO(data))
            except Exception:
                continue
            if img.width < min_size_px or img.height < min_size_px:
                continue
            out.append((name, data, img.width, img.height))
    return out


def _extract_pptx_images(path: Path, min_size_px: int) -> list[tuple[str, bytes, int, int]]:
    """Извлекает картинки из PPTX (ppt/media/)."""
    out = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.startswith("ppt/media/"):
                continue
            data = z.read(name)
            try:
                img = Image.open(io.BytesIO(data))
            except Exception:
                continue
            if img.width < min_size_px or img.height < min_size_px:
                continue
            out.append((name, data, img.width, img.height))
    return out


# ---------------------------------------------------------------------------
# Инструменты: навигация и состояние
# ---------------------------------------------------------------------------
@mcp.tool()
def open_document(file_path: str) -> str:
    """Закрепить файл (PDF/DOCX/XLSX/PPTX/TXT/MD) как активный для сессии.
    После этого остальные инструменты можно вызывать без указания пути.
    Args:
        file_path: путь к файлу.
    """
    path = _resolve_file(file_path)
    f_type = _active_file_type
    info: dict[str, Any] = {"status": "закреплён", "file": str(path), "type": f_type}

    try:
        if f_type == "pdf":
            doc = fitz.open(path)
            info["page_count"] = doc.page_count
            doc.close()
        elif f_type == "docx":
            if not _HAS_DOCX:
                raise RuntimeError("python-docx не установлен. Выполни: pip install python-docx")
            doc = _docx.Document(path)
            info["paragraphs_count"] = len(doc.paragraphs)
            info["tables_count"] = len(doc.tables)
        elif f_type == "xlsx":
            if not _HAS_XLSX:
                raise RuntimeError("openpyxl не установлен. Выполни: pip install openpyxl")
            wb = _openpyxl.load_workbook(path, read_only=True, data_only=True)
            info["sheets"] = wb.sheetnames
            info["sheets_count"] = len(wb.sheetnames)
            wb.close()
        elif f_type == "pptx":
            if not _HAS_PPTX:
                raise RuntimeError("python-pptx не установлен. Выполни: pip install python-pptx")
            prs = _pptx.Presentation(path)
            info["slides_count"] = len(prs.slides)
        elif f_type == "text":
            info["size_bytes"] = path.stat().st_size
    except Exception as e:
        info["warning"] = f"Файл закреплён, но не удалось получить детали: {e}"

    info["note"] = "Файл закреплён. Используй read_text, extract_tables, extract_embedded_images."
    return json.dumps(info, indent=2, ensure_ascii=False)


@mcp.tool()
def open_pdf(pdf_path: str) -> str:
    """Закрепить PDF-файл (алиас open_document для обратной совместимости)."""
    return open_document(pdf_path)


@mcp.tool()
def close_document() -> str:
    """Снять закрепление текущего файла."""
    global _active_file, _active_file_type
    was = str(_active_file) if _active_file else None
    _active_file = None
    _active_file_type = None
    return json.dumps({"status": "откреплён", "was": was})


@mcp.tool()
def close_pdf() -> str:
    """Алиас close_document для обратной совместимости."""
    return close_document()


@mcp.tool()
def list_documents(directory: str) -> str:
    """Найти все поддерживаемые файлы (PDF/DOCX/XLSX/PPTX/TXT/MD) в каталоге."""
    d = Path(directory).expanduser().resolve()
    if ALLOWED_ROOTS and not any(str(d).startswith(str(root)) for root in ALLOWED_ROOTS):
        raise PermissionError(f"Каталог {d} вне разрешённых корней.")
    if not d.exists():
        raise FileNotFoundError(f"Каталог не найден: {d}")
    files = []
    for ext in SUPPORTED_EXTS:
        files.extend(d.rglob(f"*{ext}"))
    files = sorted(str(p) for p in files)
    return json.dumps({"count": len(files), "files": files}, indent=2)


@mcp.tool()
def list_pdfs(directory: str) -> str:
    """Алиас list_documents для обратной совместимости."""
    return list_documents(directory)


@mcp.tool()
def get_document_info(file_path: str | None = None) -> str:
    """Метаданные закреплённого/указанного файла.
    Для PDF: страницы, размеры, слои, чертежи.
    Для Office: число абзацев/листов/слайдов, таблицы.
    """
    path = _resolve_file(file_path)
    f_type = _active_file_type

    if f_type == "pdf":
        doc = fitz.open(path)
        try:
            toc = doc.get_toc()
            pages_info = []
            for i, page in enumerate(doc):
                r = page.rect
                pages_info.append({
                    "page": i + 1,
                    "width_pt": round(r.width, 1),
                    "height_pt": round(r.height, 1),
                    "rotation": page.rotation,
                    "has_images": len(page.get_images()) > 0,
                    "has_drawings": len(page.get_drawings()) > 0,
                })
            info = {
                "file": str(path), "type": "pdf", "page_count": doc.page_count,
                "metadata": _sanitize_json_value(doc.metadata),
                "table_of_contents": [{"level": l, "title": _sanitize_text(t), "page": p} for l, t, p in toc],
                "pages": pages_info,
            }
            return json.dumps(info, indent=2, ensure_ascii=False)
        finally:
            doc.close()

    if f_type == "docx":
        if not _HAS_DOCX:
            raise RuntimeError("python-docx не установлен.")
        doc = _docx.Document(path)
        core = doc.core_properties
        return json.dumps({
            "file": str(path), "type": "docx",
            "paragraphs_count": len(doc.paragraphs),
            "tables_count": len(doc.tables),
            "sections_count": len(doc.sections),
            "metadata": {
                "title": _sanitize_text(core.title),
                "author": _sanitize_text(core.author),
                "created": str(core.created) if core.created else None,
                "modified": str(core.modified) if core.modified else None,
            },
        }, indent=2, ensure_ascii=False)

    if f_type == "xlsx":
        if not _HAS_XLSX:
            raise RuntimeError("openpyxl не установлен.")
        wb = _openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheets_info = []
        for name in wb.sheetnames:
            ws = wb[name]
            sheets_info.append({
                "name": name,
                "dimensions": ws.dimensions,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
            })
        props = wb.properties
        result = {
            "file": str(path), "type": "xlsx",
            "sheets_count": len(wb.sheetnames),
            "sheets": sheets_info,
            "metadata": {
                "title": _sanitize_text(props.title),
                "creator": _sanitize_text(props.creator),
            },
        }
        wb.close()
        return json.dumps(result, indent=2, ensure_ascii=False)

    if f_type == "pptx":
        if not _HAS_PPTX:
            raise RuntimeError("python-pptx не установлен.")
        prs = _pptx.Presentation(path)
        slides_info = []
        for i, slide in enumerate(prs.slides):
            shapes_summary = {"text": 0, "tables": 0, "images": 0, "other": 0}
            for shp in slide.shapes:
                if shp.has_text_frame:
                    shapes_summary["text"] += 1
                elif shp.has_table:
                    shapes_summary["tables"] += 1
                elif shp.shape_type == 13:  # Picture
                    shapes_summary["images"] += 1
                else:
                    shapes_summary["other"] += 1
            slides_info.append({
                "slide": i + 1,
                "layout": slide.slide_layout.name if slide.slide_layout else None,
                "shapes": shapes_summary,
                "has_notes": bool(slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip()),
            })
        w_in = prs.slide_width / 914400 if prs.slide_width else None
        h_in = prs.slide_height / 914400 if prs.slide_height else None
        return json.dumps({
            "file": str(path), "type": "pptx",
            "slides_count": len(prs.slides),
            "slide_size_inches": {"width": round(w_in, 2) if w_in else None, "height": round(h_in, 2) if h_in else None},
            "slides": slides_info,
        }, indent=2, ensure_ascii=False)

    if f_type == "text":
        return json.dumps({
            "file": str(path), "type": "text",
            "size_bytes": path.stat().st_size,
        }, indent=2, ensure_ascii=False)

    return json.dumps({"error": f"Неизвестный тип файла: {f_type}"})


@mcp.tool()
def get_pdf_info(pdf_path: str | None = None) -> str:
    """Алиас get_document_info для PDF (обратная совместимость)."""
    return get_document_info(pdf_path)


# ---------------------------------------------------------------------------
# Инструменты: текст
# ---------------------------------------------------------------------------
@mcp.tool()
def read_text(file_path: str | None = None, page_range: str | None = None) -> str:
    """Извлечь текст из PDF/DOCX/XLSX/PPTX/TXT.
    Для PDF: постранично. Для Office: по абзацам/листам/слайдам.
    """
    path = _resolve_file(file_path)
    f_type = _active_file_type

    if f_type == "pdf":
        doc = fitz.open(path)
        try:
            indices = _parse_page_range(page_range, doc.page_count)
            parts = []
            for i in indices:
                text = _sanitize_text(doc[i].get_text("text")).strip()
                parts.append(f"--- Страница {i + 1} ---\n{text if text else '(текст не найден — возможно скан, см. render_page_image)'}")
            return "\n".join(parts)
        finally:
            doc.close()

    if f_type == "docx":
        if not _HAS_DOCX:
            raise RuntimeError("python-docx не установлен.")
        doc = _docx.Document(path)
        return "\n".join([_sanitize_text(p.text) for p in doc.paragraphs if p.text.strip()])

    if f_type == "xlsx":
        if not _HAS_XLSX:
            raise RuntimeError("openpyxl не установлен.")
        wb = _openpyxl.load_workbook(path, data_only=True)
        parts = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            parts.append(f"--- Лист: {sheet_name} ---")
            for row in ws.iter_rows(values_only=True):
                row_text = " | ".join([str(c) if c is not None else "" for c in row])
                if row_text.strip("| "):
                    parts.append(row_text)
        wb.close()
        return "\n".join(parts)

    if f_type == "pptx":
        if not _HAS_PPTX:
            raise RuntimeError("python-pptx не установлен.")
        prs = _pptx.Presentation(path)
        parts = []
        for i, slide in enumerate(prs.slides):
            slide_text = []
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip():
                    slide_text.append(shape.text.strip())
            parts.append(f"--- Слайд {i + 1} ---\n" + "\n".join(slide_text))
        return "\n".join(parts)

    if f_type == "text":
        return _sanitize_text(path.read_text(encoding="utf-8", errors="replace"))

    return "Формат файла не поддерживается для чтения текста."


@mcp.tool()
def search_text(query: str, file_path: str | None = None, case_sensitive: bool = False) -> str:
    """Найти текст в PDF/DOCX/PPTX. Возвращает страницы/слайды с контекстом."""
    path = _resolve_file(file_path)
    f_type = _active_file_type
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(query), flags)

    if f_type == "pdf":
        doc = fitz.open(path)
        try:
            results = []
            for i, page in enumerate(doc):
                rects = page.search_for(query, quads=False)
                if rects:
                    text = _sanitize_text(page.get_text("text"))
                    results.append({"page": i + 1, "matches": len(rects), "snippet": text[:300].strip()})
            return json.dumps({"query": query, "found_on_pages": len(results), "results": results}, indent=2, ensure_ascii=False)
        finally:
            doc.close()

    if f_type == "docx":
        if not _HAS_DOCX:
            raise RuntimeError("python-docx не установлен.")
        doc = _docx.Document(path)
        matches = 0
        snippets = []
        for p in doc.paragraphs:
            if pattern.search(p.text):
                matches += 1
                if len(snippets) < 10:
                    snippets.append(_sanitize_text(p.text)[:300])
        return json.dumps({"query": query, "matches": matches, "snippets": snippets}, indent=2, ensure_ascii=False)

    if f_type == "pptx":
        if not _HAS_PPTX:
            raise RuntimeError("python-pptx не установлен.")
        prs = _pptx.Presentation(path)
        results = []
        for i, slide in enumerate(prs.slides):
            slide_matches = 0
            slide_text = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        if pattern.search(para.text):
                            slide_matches += 1
                            slide_text.append(para.text.strip())
            if slide_matches:
                results.append({"slide": i + 1, "matches": slide_matches, "snippets": [_sanitize_text(t)[:200] for t in slide_text[:5]]})
        return json.dumps({"query": query, "found_on_slides": len(results), "results": results}, indent=2, ensure_ascii=False)

    return json.dumps({"error": f"Поиск не поддерживается для формата {f_type}"})


@mcp.tool()
def extract_tables(file_path: str | None = None, page_range: str | None = None) -> str:
    """Извлечь таблицы из PDF/DOCX/XLSX."""
    path = _resolve_file(file_path)
    f_type = _active_file_type

    if f_type == "pdf":
        with pdfplumber.open(path) as pdf:
            indices = _parse_page_range(page_range, len(pdf.pages))
            out = []
            for i in indices:
                tables = pdf.pages[i].extract_tables()
                if tables:
                    out.append({"page": i + 1, "tables": _sanitize_json_value(tables)})
            return json.dumps({"pages_with_tables": len(out), "data": out}, indent=2, ensure_ascii=False)

    if f_type == "docx":
        if not _HAS_DOCX:
            raise RuntimeError("python-docx не установлен.")
        doc = _docx.Document(path)
        out = []
        for i, table in enumerate(doc.tables):
            rows = []
            for row in table.rows:
                rows.append([_sanitize_text(cell.text) for cell in row.cells])
            out.append({"table_index": i, "rows": rows})
        return json.dumps({"tables_found": len(out), "data": out}, indent=2, ensure_ascii=False)

    if f_type == "xlsx":
        if not _HAS_XLSX:
            raise RuntimeError("openpyxl не установлен.")
        wb = _openpyxl.load_workbook(path, data_only=True)
        out = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append([str(c) if c is not None else "" for c in row])
            if rows:
                out.append({"sheet": sheet_name, "rows": rows})
        wb.close()
        return json.dumps({"sheets_with_data": len(out), "data": out}, indent=2, ensure_ascii=False)

    return json.dumps({"error": f"Извлечение таблиц для формата {f_type} не реализовано."})


# ---------------------------------------------------------------------------
# CAD-инструменты (СТРОГО ДЛЯ PDF)
# ---------------------------------------------------------------------------
@mcp.tool()
def summarize_vector_graphics(
    page_number: int,
    pdf_path: str | None = None,
    x0: float | None = None, y0: float | None = None,
    x1: float | None = None, y1: float | None = None,
) -> str:
    """Обзор векторной графики PDF БЕЗ полных координат. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    doc = fitz.open(path)
    try:
        if not (1 <= page_number <= doc.page_count):
            raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
        page = doc[page_number - 1]
        region = fitz.Rect(x0, y0, x1, y1) if None not in (x0, y0, x1, y1) else None
        drawings = page.get_drawings(extended=True)
        paths_by_layer: dict[str, int] = {}
        items_by_type: dict[str, int] = {}
        overall_bbox: fitz.Rect | None = None
        total_paths = total_items = 0
        for d in drawings:
            rect = d.get("rect")
            if region is not None and rect is not None and not region.intersects(rect):
                continue
            layer_name = d.get("layer") or "(без слоя)"
            paths_by_layer[layer_name] = paths_by_layer.get(layer_name, 0) + 1
            total_paths += 1
            for it in d["items"]:
                op = {"l": "line", "re": "rect", "c": "curve", "qu": "quad"}.get(it[0], str(it[0]))
                items_by_type[op] = items_by_type.get(op, 0) + 1
                total_items += 1
            if rect is not None:
                overall_bbox = rect if overall_bbox is None else (overall_bbox | rect)
        result = {
            "page": page_number, "total_paths": total_paths, "total_items": total_items,
            "paths_by_layer": paths_by_layer, "items_by_type": items_by_type,
            "overall_bbox": [round(overall_bbox.x0, 2), round(overall_bbox.y0, 2), round(overall_bbox.x1, 2), round(overall_bbox.y1, 2)] if overall_bbox else None,
        }
        if total_items > 300:
            result["note"] = "Геометрии много. Вызывай extract_vector_graphics ТОЧЕЧНО (по слою или области)."
        return json.dumps(result, indent=2, ensure_ascii=False)
    finally:
        doc.close()


@mcp.tool()
def extract_vector_graphics(
    page_number: int,
    pdf_path: str | None = None,
    x0: float | None = None, y0: float | None = None,
    x1: float | None = None, y1: float | None = None,
    layers: list[str] | None = None,
    max_paths: int = 200,
) -> str:
    """Векторная геометрия PDF с координатами. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    doc = fitz.open(path)
    try:
        if not (1 <= page_number <= doc.page_count):
            raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
        page = doc[page_number - 1]
        region = fitz.Rect(x0, y0, x1, y1) if None not in (x0, y0, x1, y1) else None
        drawings = page.get_drawings(extended=True)
        paths = []
        counts: dict[str, int] = {}
        truncated = False
        for d in drawings:
            layer_name = d.get("layer") or None
            if layers and layer_name not in layers:
                continue
            rect = d.get("rect")
            if region is not None and rect is not None and not region.intersects(rect):
                continue
            if len(paths) >= max_paths:
                truncated = True
                break
            items = _serialize_drawing_items(d["items"])
            for it in items:
                counts[it["op"]] = counts.get(it["op"], 0) + 1
            paths.append({
                "layer": layer_name,
                "kind": {"s": "stroke", "f": "fill", "fs": "fill+stroke"}.get(d.get("type"), d.get("type")),
                "color": d.get("color"), "fill": d.get("fill"),
                "width": d.get("width"), "dashes": d.get("dashes"),
                "closed": d.get("closePath"),
                "bbox": [round(rect.x0, 2), round(rect.y0, 2), round(rect.x1, 2), round(rect.y1, 2)] if rect else None,
                "items": items,
            })
        result = {"page": page_number, "total_paths_returned": len(paths), "truncated": truncated, "element_counts": counts, "paths": paths}
        if truncated:
            result["note"] = f"Вывод обрезан по max_paths={max_paths}. Уточни область или слой."
        return json.dumps(result, indent=2, ensure_ascii=False)
    finally:
        doc.close()


@mcp.tool()
def list_layers(pdf_path: str | None = None) -> str:
    """Список слоёв (OCG) в PDF. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    doc = fitz.open(path)
    try:
        ocgs = doc.get_ocgs()
        layers = [{"xref": xref, "name": _sanitize_text(info.get("name")), "visible": info.get("on", True)} for xref, info in ocgs.items()]
        result = {"layer_count": len(layers), "layers": layers}
        if not layers:
            result["note"] = "В этом PDF слои (OCG) не обнаружены."
        return json.dumps(result, indent=2, ensure_ascii=False)
    finally:
        doc.close()


@mcp.tool()
def get_text_positions(
    page_number: int,
    pdf_path: str | None = None,
    x0: float | None = None, y0: float | None = None,
    x1: float | None = None, y1: float | None = None,
) -> str:
    """Слова страницы с координатами. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    doc = fitz.open(path)
    try:
        if not (1 <= page_number <= doc.page_count):
            raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
        page = doc[page_number - 1]
        region = fitz.Rect(x0, y0, x1, y1) if None not in (x0, y0, x1, y1) else None
        words = page.get_text("words")
        out = []
        for w in words:
            wx0, wy0, wx1, wy1, text = w[0], w[1], w[2], w[3], w[4]
            if region is not None and not region.intersects(fitz.Rect(wx0, wy0, wx1, wy1)):
                continue
            out.append({"text": _sanitize_text(text), "bbox": [round(wx0, 2), round(wy0, 2), round(wx1, 2), round(wy1, 2)]})
        return json.dumps({"page": page_number, "word_count": len(out), "words": out}, indent=2, ensure_ascii=False)
    finally:
        doc.close()


@mcp.tool()
def extract_annotations(page_number: int | None = None, pdf_path: str | None = None) -> str:
    """PDF-аннотации (markup). ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    doc = fitz.open(path)
    try:
        if page_number is not None:
            if not (1 <= page_number <= doc.page_count):
                raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
            indices = [page_number - 1]
        else:
            indices = list(range(doc.page_count))
        out = []
        for i in indices:
            page = doc[i]
            for annot in page.annots() or []:
                info = annot.info
                r = annot.rect
                out.append({
                    "page": i + 1, "type": annot.type[1],
                    "content": _sanitize_text(info.get("content", "")),
                    "author": _sanitize_text(info.get("title", "")),
                    "bbox": [round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2)],
                    "created": info.get("creationDate", ""), "modified": info.get("modDate", ""),
                })
        return json.dumps({"annotation_count": len(out), "annotations": out}, indent=2, ensure_ascii=False)
    finally:
        doc.close()


@mcp.tool()
def calibrate_scale(
    x0: float, y0: float, x1: float, y1: float,
    real_length: float, unit: str = "мм",
    pdf_path: str | None = None,
) -> str:
    """Откалибровать масштаб чертежа. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    if real_length <= 0:
        raise ValueError("real_length должен быть положительным.")
    dist_pt = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5
    points_per_unit = dist_pt / real_length
    _scale_cache[str(path)] = (points_per_unit, unit)
    return json.dumps({
        "file": str(path), "measured_points": round(dist_pt, 2),
        "real_length": real_length, "unit": unit,
        "points_per_unit": round(points_per_unit, 4),
    }, indent=2, ensure_ascii=False)


@mcp.tool()
def set_scale(points_per_unit: float, unit: str = "мм", pdf_path: str | None = None) -> str:
    """Задать масштаб напрямую. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    _scale_cache[str(path)] = (points_per_unit, unit)
    return json.dumps({"file": str(path), "points_per_unit": points_per_unit, "unit": unit}, ensure_ascii=False)


@mcp.tool()
def measure_distance(x0: float, y0: float, x1: float, y1: float, pdf_path: str | None = None) -> str:
    """Измерить расстояние между точками. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    dist_pt = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5
    result: dict[str, Any] = {"distance_points": round(dist_pt, 2)}
    scale = _scale_cache.get(str(path))
    if scale:
        ppu, unit = scale
        result["distance_real"] = round(dist_pt / ppu, 3)
        result["unit"] = unit
    else:
        result["note"] = "Масштаб не задан — вызови calibrate_scale()."
    return json.dumps(result, indent=2, ensure_ascii=False)


@mcp.tool()
def export_page_svg(page_number: int, pdf_path: str | None = None, out_path: str | None = None) -> str:
    """Экспорт страницы PDF в SVG. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    doc = fitz.open(path)
    try:
        if not (1 <= page_number <= doc.page_count):
            raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
        page = doc[page_number - 1]
        svg = page.get_svg_image()
        out = _check_writable(Path(out_path)) if out_path else path.with_name(f"{path.stem}_p{page_number}.svg")
        out.write_text(svg, encoding="utf-8")
        return json.dumps({"saved_to": str(out), "size_bytes": len(svg.encode("utf-8"))}, indent=2, ensure_ascii=False)
    finally:
        doc.close()


@mcp.tool()
def export_page_dxf(
    page_number: int,
    pdf_path: str | None = None,
    out_path: str | None = None,
    layers: list[str] | None = None,
    x0: float | None = None, y0: float | None = None,
    x1: float | None = None, y1: float | None = None,
) -> str:
    """Экспорт геометрии PDF в DXF. ТОЛЬКО для PDF."""
    if not _HAS_EZDXF:
        raise RuntimeError("ezdxf не установлен. Выполни: pip install ezdxf")
    path = _ensure_pdf(pdf_path)
    doc = fitz.open(path)
    try:
        if not (1 <= page_number <= doc.page_count):
            raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
        page = doc[page_number - 1]
        page_height = page.rect.height
        region = fitz.Rect(x0, y0, x1, y1) if None not in (x0, y0, x1, y1) else None
        drawings = page.get_drawings(extended=True)
        ppu, _unit = _scale_cache.get(str(path), (1.0, "pt"))

        def conv(x: float, y: float) -> tuple[float, float]:
            return (x / ppu, (page_height - y) / ppu)

        dxf = ezdxf.new("R2010")
        msp = dxf.modelspace()
        seen_layers: set[str] = set()
        entity_count = 0
        for d in drawings:
            layer_name = d.get("layer") or "0"
            if layers and (d.get("layer") or None) not in layers:
                continue
            rect = d.get("rect")
            if region is not None and rect is not None and not region.intersects(rect):
                continue
            if layer_name not in seen_layers:
                if layer_name not in dxf.layers:
                    dxf.layers.add(layer_name)
                seen_layers.add(layer_name)
            for it in d["items"]:
                op = it[0]
                if op == "l":
                    p1, p2 = conv(it[1].x, it[1].y), conv(it[2].x, it[2].y)
                    msp.add_line(p1, p2, dxfattribs={"layer": layer_name})
                    entity_count += 1
                elif op == "re":
                    r = it[1]
                    pts = [conv(r.x0, r.y0), conv(r.x1, r.y0), conv(r.x1, r.y1), conv(r.x0, r.y1)]
                    msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer_name})
                    entity_count += 1
                elif op == "c":
                    pts = [conv(p.x, p.y) for p in it[1:]]
                    msp.add_spline(fit_points=pts, dxfattribs={"layer": layer_name})
                    entity_count += 1
                elif op == "qu":
                    q = it[1]
                    pts = [conv(q.ul.x, q.ul.y), conv(q.ur.x, q.ur.y), conv(q.lr.x, q.lr.y), conv(q.ll.x, q.ll.y)]
                    msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer_name})
                    entity_count += 1
        out = _check_writable(Path(out_path)) if out_path else path.with_name(f"{path.stem}_p{page_number}.dxf")
        dxf.saveas(str(out))
        return json.dumps({
            "saved_to": str(out), "entity_count": entity_count, "layers": sorted(seen_layers),
            "scale_used": f"1 unit DXF = {ppu} pt PDF" if str(path) in _scale_cache else "масштаб не задан — координаты в points PDF",
        }, indent=2, ensure_ascii=False)
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# Визуальные инструменты
# ---------------------------------------------------------------------------
@mcp.tool()
def render_page_image(page_number: int, pdf_path: str | None = None, dpi: int = DEFAULT_DPI) -> list[ImageContent | TextContent]:
    """Отрендерить страницу PDF в PNG. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    dpi = max(50, min(dpi, MAX_RENDER_DPI))
    doc = fitz.open(path)
    try:
        if not (1 <= page_number <= doc.page_count):
            raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
        page = doc[page_number - 1]
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        png_bytes = pix.tobytes("png")
        b64 = base64.b64encode(png_bytes).decode("ascii")
        return [
            TextContent(type="text", text=f"Страница {page_number}, рендер {pix.width}x{pix.height}px @ {dpi} DPI."),
            ImageContent(type="image", data=b64, mimeType="image/png"),
        ]
    finally:
        doc.close()


@mcp.tool()
def render_region_image(
    page_number: int,
    x0: float, y0: float, x1: float, y1: float,
    pdf_path: str | None = None, dpi: int = 300,
) -> list[ImageContent | TextContent]:
    """Отрендерить область страницы PDF. ТОЛЬКО для PDF."""
    path = _ensure_pdf(pdf_path)
    dpi = max(50, min(dpi, MAX_RENDER_DPI))
    doc = fitz.open(path)
    try:
        if not (1 <= page_number <= doc.page_count):
            raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
        page = doc[page_number - 1]
        clip = fitz.Rect(x0, y0, x1, y1)
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
        png_bytes = pix.tobytes("png")
        b64 = base64.b64encode(png_bytes).decode("ascii")
        return [
            TextContent(type="text", text=f"Фрагмент страницы {page_number} [{x0},{y0},{x1},{y1}], {pix.width}x{pix.height}px @ {dpi} DPI."),
            ImageContent(type="image", data=b64, mimeType="image/png"),
        ]
    finally:
        doc.close()


@mcp.tool()
def extract_embedded_images(
    page_number: int | None = None,
    file_path: str | None = None,
    min_size_px: int = 100,
) -> list[ImageContent | TextContent]:
    """Извлечь встроенные картинки.
    PDF: со страницы. DOCX/XLSX/PPTX: все картинки из файла (через ZIP-архив).
    """
    path = _resolve_file(file_path)
    f_type = _active_file_type
    content: list[ImageContent | TextContent] = []

    if f_type == "pdf":
        doc = fitz.open(path)
        try:
            if page_number is None:
                indices = list(range(doc.page_count))
            else:
                if not (1 <= page_number <= doc.page_count):
                    raise ValueError(f"Страница {page_number} вне диапазона 1..{doc.page_count}")
                indices = [page_number - 1]
            kept = 0
            for i in indices:
                page = doc[i]
                for img in page.get_images(full=True):
                    xref = img[0]
                    base = doc.extract_image(xref)
                    img_bytes = base["image"]
                    pil_img = Image.open(io.BytesIO(img_bytes))
                    if pil_img.width < min_size_px or pil_img.height < min_size_px:
                        continue
                    buf = io.BytesIO()
                    pil_img.convert("RGB").save(buf, format="PNG")
                    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                    content.append(TextContent(type="text", text=f"Стр. {i+1}, изображение {pil_img.width}x{pil_img.height}px (xref {xref})"))
                    content.append(ImageContent(type="image", data=b64, mimeType="image/png"))
                    kept += 1
            if kept == 0:
                content.append(TextContent(type="text", text="Подходящих изображений не найдено."))
            return content
        finally:
            doc.close()

    # Office-файлы: извлекаем из ZIP-архива
    images_list: list[tuple[str, bytes, int, int]] = []
    if f_type == "docx":
        if not _HAS_DOCX:
            raise RuntimeError("python-docx не установлен.")
        images_list = _extract_docx_images(path, min_size_px)
    elif f_type == "xlsx":
        if not _HAS_XLSX:
            raise RuntimeError("openpyxl не установлен.")
        images_list = _extract_xlsx_images(path, min_size_px)
    elif f_type == "pptx":
        if not _HAS_PPTX:
            raise RuntimeError("python-pptx не установлен.")
        images_list = _extract_pptx_images(path, min_size_px)
    else:
        return [TextContent(type="text", text=f"Извлечение картинок для формата {f_type} не поддерживается.")]

    if not images_list:
        return [TextContent(type="text", text="Подходящих изображений не найдено.")]

    for name, data, w, h in images_list:
        # Конвертируем в PNG для единообразия
        pil_img = Image.open(io.BytesIO(data))
        buf = io.BytesIO()
        pil_img.convert("RGB").save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        content.append(TextContent(type="text", text=f"Изображение из {name} ({w}x{h}px)"))
        content.append(ImageContent(type="image", data=b64, mimeType="image/png"))
    return content


@mcp.tool()
def extract_pptx_notes(file_path: str | None = None) -> str:
    """Извлечь заметки докладчика из PPTX. ТОЛЬКО для PPTX."""
    path = _resolve_file(file_path)
    if _active_file_type != "pptx":
        raise ValueError(f"Этот инструмент только для PPTX. Сейчас открыт: {path.name} ({_active_file_type})")
    if not _HAS_PPTX:
        raise RuntimeError("python-pptx не установлен.")
    prs = _pptx.Presentation(path)
    parts = []
    for i, slide in enumerate(prs.slides):
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
            if notes_text:
                parts.append(f"--- Слайд {i + 1} ---\n{_sanitize_text(notes_text)}")
    if not parts:
        return json.dumps({"note": "Заметки докладчика не найдены."})
    return "\n".join(parts)


if __name__ == "__main__":
    mcp.run(transport="stdio")