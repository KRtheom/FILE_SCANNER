"""FILE SCANNER 핵심 엔진 함수 모음."""

from __future__ import annotations

import csv
import json
import logging
import os
import re
import threading
import xml.etree.ElementTree as ET
import zipfile
import zlib
from datetime import datetime
from typing import Any

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover - optional dependency
    fitz = None  # type: ignore[assignment]

try:
    import pdfplumber
except Exception:  # pragma: no cover - optional dependency
    pdfplumber = None  # type: ignore[assignment]

try:
    import olefile
except Exception:  # pragma: no cover - optional dependency
    olefile = None  # type: ignore[assignment]

try:
    import openpyxl
except Exception:  # pragma: no cover - optional dependency
    openpyxl = None  # type: ignore[assignment]

try:
    import xlrd
except Exception:  # pragma: no cover - optional dependency
    xlrd = None  # type: ignore[assignment]

try:
    from docx import Document
except Exception:  # pragma: no cover - optional dependency
    Document = None  # type: ignore[assignment]


logger = logging.getLogger(__name__)

DEFAULT_EXTENSIONS: set[str] = {
    ".pdf",
    ".xlsx",
    ".xls",
    ".docx",
    ".doc",
    ".hwp",
    ".hwpx",
    ".csv",
    ".txt",
}

EXCLUDED_DIR_NAMES: set[str] = {
    "windows",
    "program files",
    "program files (x86)",
    "$recycle.bin",
    "system volume information",
    "programdata",
    "appdata",
    "venv",
    "node_modules",
    ".git",
    "__pycache__",
    "dist",
    "build",
}

TXT_FALLBACK_ENCODINGS: tuple[str, ...] = ("utf-8", "cp949", "euc-kr", "mbcs")
QUICK_CHECK_ENCODINGS: tuple[str, ...] = ("utf-8", "cp949", "euc-kr", "utf-16-le")
QUICK_CHECK_MAX_BYTES = 100 * 1024 * 1024

_SEARCH_FILE_META: dict[str, tuple[bool, bool]] = {}
_SEARCH_FILE_META_LOCK = threading.Lock()


def scan_files(paths: list[str], extensions: set[str]) -> list[str]:
    """지정 경로를 재귀 탐색해 대상 확장자 파일의 절대경로 목록을 반환한다."""
    if not paths:
        return []

    normalized_extensions = _normalize_extensions(extensions or DEFAULT_EXTENSIONS)
    if not normalized_extensions:
        return []

    file_paths: set[str] = set()
    visited_dirs: set[str] = set()

    def walk_dir(root_dir: str) -> None:
        normalized_root = os.path.normcase(os.path.abspath(root_dir))
        if normalized_root in visited_dirs:
            return
        visited_dirs.add(normalized_root)

        try:
            with os.scandir(root_dir) as entries:
                for entry in entries:
                    entry_path = os.path.abspath(entry.path)
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            if _is_excluded_dir(entry.name):
                                continue
                            walk_dir(entry_path)
                        elif entry.is_file(follow_symlinks=False):
                            _, ext = os.path.splitext(entry.name)
                            if ext.lower() in normalized_extensions:
                                file_paths.add(entry_path)
                    except (PermissionError, FileNotFoundError, OSError) as exc:
                        logger.debug("항목 접근 스킵: %s (%s)", entry_path, exc)
        except (PermissionError, FileNotFoundError, NotADirectoryError, OSError) as exc:
            logger.debug("디렉터리 접근 스킵: %s (%s)", root_dir, exc)

    for raw_path in paths:
        if not raw_path:
            continue

        abs_path = os.path.abspath(raw_path)
        if os.path.isfile(abs_path):
            _, ext = os.path.splitext(abs_path)
            if ext.lower() in normalized_extensions:
                file_paths.add(abs_path)
            continue

        if os.path.isdir(abs_path):
            base_name = os.path.basename(abs_path.rstrip("\\/")).lower()
            if base_name and _is_excluded_dir(base_name):
                continue
            walk_dir(abs_path)

    return sorted(file_paths)


def extract_text(filepath: str) -> list[tuple[str, str]]:
    """파일 확장자별 텍스트를 추출해 (위치, 텍스트) 목록으로 반환한다."""
    text_items, _ = _extract_text_with_status(filepath)
    return text_items


def quick_check(filepath: str, keywords: list[str]) -> bool:
    """바이너리 레벨에서 키워드 포함 여부를 사전 확인한다."""
    normalized_keywords = _normalize_keyword_list(keywords)
    if not normalized_keywords:
        return False

    try:
        if os.path.getsize(filepath) > QUICK_CHECK_MAX_BYTES:
            return False
    except Exception as exc:
        logger.debug("파일 크기 확인 실패: %s (%s)", filepath, exc)
        return True

    patterns = _build_quick_check_patterns(normalized_keywords)
    if not patterns:
        return False

    max_pattern_len = max(len(pattern) for pattern in patterns)
    overlap_len = max(0, max_pattern_len - 1)

    try:
        with open(filepath, "rb") as file:
            tail = b""
            while True:
                chunk = file.read(1024 * 1024)
                if not chunk:
                    break

                buffer = tail + chunk
                for pattern in patterns:
                    if pattern in buffer:
                        return True

                tail = buffer[-overlap_len:] if overlap_len > 0 else b""
    except Exception as exc:
        logger.debug("빠른 필터링 실패(정밀검사 포함 처리): %s (%s)", filepath, exc)
        return True

    return False


def search_file(filepath: str, keywords: list[str]) -> list[dict[str, str]]:
    """1단계 quick_check 후 통과 파일만 정밀 추출/검색해 결과를 반환한다."""
    normalized_keywords = _normalize_keyword_list(keywords)
    skipped = False
    failed = False

    try:
        if not normalized_keywords:
            skipped = True
            return []

        if not quick_check(filepath, normalized_keywords):
            skipped = True
            return []

        text_items, failed = _extract_text_with_status(filepath)
        if failed:
            return []

        matches = search_keywords(text_items, normalized_keywords)
        return [
            {
                "keyword": str(item.get("keyword", "")),
                "file": filepath,
                "location": str(item.get("location", "")),
                "context": str(item.get("context", "")),
            }
            for item in matches
        ]
    except Exception as exc:
        logger.debug("파일 검색 실패: %s (%s)", filepath, exc)
        failed = True
        return []
    finally:
        _set_search_file_meta(filepath, skipped=skipped, failed=failed)


def consume_search_file_meta(filepath: str) -> dict[str, bool]:
    """search_file 실행 메타(스킵/실패)를 반환하고 내부 저장소에서 제거한다."""
    normalized = os.path.normcase(os.path.abspath(filepath))
    with _SEARCH_FILE_META_LOCK:
        skipped, failed = _SEARCH_FILE_META.pop(normalized, (False, False))
    return {"skipped": skipped, "failed": failed}


def _extract_text_with_status(filepath: str) -> tuple[list[tuple[str, str]], bool]:
    """확장자별 텍스트 추출 결과와 실패 여부를 함께 반환한다."""
    _, ext = os.path.splitext(filepath)
    ext = ext.lower()

    try:
        if ext == ".pdf":
            return _extract_pdf(filepath), False
        if ext == ".xlsx":
            return _extract_xlsx(filepath), False
        if ext == ".xls":
            return _extract_xls(filepath), False
        if ext == ".docx":
            return _extract_docx(filepath), False
        if ext == ".doc":
            return _extract_doc(filepath), False
        if ext == ".hwp":
            return _extract_hwp(filepath), False
        if ext == ".hwpx":
            return _extract_hwpx(filepath), False
        if ext == ".csv":
            return _extract_csv(filepath), False
        if ext == ".txt":
            return _extract_txt(filepath), False
    except Exception as exc:
        logger.debug("텍스트 추출 실패: %s (%s)", filepath, exc)
        return [], True

    return [], False


def search_keywords(text_items: list[tuple[str, str]], keywords: list[str]) -> list[dict[str, str]]:
    """텍스트 목록에서 키워드를 검색해 위치와 문맥 정보를 반환한다."""
    normalized_keywords = _normalize_keyword_list(keywords)
    if not normalized_keywords or not text_items:
        return []

    results: list[dict[str, str]] = []

    for location, text in text_items:
        if not text:
            continue
        lowered_text = text.lower()

        for keyword in normalized_keywords:
            lowered_keyword = keyword.lower()
            start_idx = 0

            while True:
                found_idx = lowered_text.find(lowered_keyword, start_idx)
                if found_idx < 0:
                    break

                context_start = max(0, found_idx - 50)
                context_end = min(len(text), found_idx + len(keyword) + 50)
                context = text[context_start:context_end].strip()

                results.append(
                    {
                        "keyword": keyword,
                        "location": location,
                        "context": context,
                    }
                )
                start_idx = found_idx + max(1, len(lowered_keyword))

    return results


def load_keywords(path: str = "keywords.json") -> list[str]:
    """키워드 JSON 파일을 읽어 키워드 목록을 반환한다."""
    if not os.path.exists(path):
        return []

    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
    except (OSError, json.JSONDecodeError) as exc:
        logger.warning("키워드 로드 실패: %s (%s)", path, exc)
        return []

    if not isinstance(data, list):
        logger.warning("키워드 형식 오류: %s (list 아님)", path)
        return []

    string_items = [item for item in data if isinstance(item, str)]
    return _normalize_keyword_list(string_items)


def save_keywords(keywords: list[str], path: str = "keywords.json") -> None:
    """키워드 목록을 JSON 파일로 저장한다."""
    normalized = _normalize_keyword_list(keywords)

    try:
        parent = os.path.dirname(os.path.abspath(path))
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(path, "w", encoding="utf-8") as file:
            json.dump(normalized, file, ensure_ascii=False, indent=2)
    except OSError as exc:
        logger.warning("키워드 저장 실패: %s (%s)", path, exc)


def save_report(results: list[dict[str, Any]], output_path: str) -> None:
    """검색 결과를 Excel 리포트 파일로 저장한다."""
    if openpyxl is None:
        logger.warning("openpyxl 미설치로 리포트 저장 불가: %s", output_path)
        return

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "검색결과"

    worksheet.append(["키워드", "파일경로", "위치", "해당문장", "검색일시"])
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for row in results:
        keyword = str(row.get("keyword", "") or "")
        file_path = str(row.get("file_path", row.get("filepath", "")) or "")
        location = str(row.get("location", "") or "")
        context = str(row.get("context", "") or "")
        searched_at = str(row.get("searched_at", now) or now)
        worksheet.append([keyword, file_path, location, context, searched_at])

    try:
        parent = os.path.dirname(os.path.abspath(output_path))
        if parent:
            os.makedirs(parent, exist_ok=True)
        workbook.save(output_path)
    except OSError as exc:
        logger.warning("리포트 저장 실패: %s (%s)", output_path, exc)
    finally:
        workbook.close()


def _extract_pdf(filepath: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []

    if fitz is not None:
        try:
            with fitz.open(filepath) as document:
                for page_num, page in enumerate(document, start=1):
                    raw_text = page.get_text("text") or ""
                    for line_num, line in enumerate(raw_text.splitlines(), start=1):
                        cleaned = line.strip()
                        if cleaned:
                            items.append((f"P{page_num} L{line_num}", cleaned))
            return items
        except Exception as exc:
            logger.debug("PyMuPDF 추출 실패, fallback 시도: %s (%s)", filepath, exc)
            items.clear()

    if pdfplumber is None:
        raise RuntimeError("PDF 추출 라이브러리(fitz/pdfplumber) 없음")

    with pdfplumber.open(filepath) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            raw_text = page.extract_text() or ""
            for line_num, line in enumerate(raw_text.splitlines(), start=1):
                cleaned = line.strip()
                if cleaned:
                    items.append((f"P{page_num} L{line_num}", cleaned))

    return items


def _extract_xlsx(filepath: str) -> list[tuple[str, str]]:
    if openpyxl is None:
        raise RuntimeError("openpyxl 미설치")

    items: list[tuple[str, str]] = []
    workbook = openpyxl.load_workbook(filepath, read_only=True, data_only=True)

    try:
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    value = cell.value
                    if value is None:
                        continue
                    text = str(value).strip()
                    if text:
                        items.append((f"{sheet.title} {cell.coordinate}", text))
    finally:
        workbook.close()

    return items


def _extract_xls(filepath: str) -> list[tuple[str, str]]:
    if xlrd is None:
        raise RuntimeError("xlrd 미설치")

    items: list[tuple[str, str]] = []
    workbook = xlrd.open_workbook(filepath, on_demand=True)

    try:
        for sheet in workbook.sheets():
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    value = sheet.cell_value(row_idx, col_idx)
                    if value in (None, ""):
                        continue
                    text = str(value).strip()
                    if text:
                        items.append((f"{sheet.name} R{row_idx + 1}C{col_idx + 1}", text))
    finally:
        try:
            workbook.release_resources()
        except Exception:
            pass

    return items


def _extract_docx(filepath: str) -> list[tuple[str, str]]:
    if Document is None:
        raise RuntimeError("python-docx 미설치")

    items: list[tuple[str, str]] = []
    document = Document(filepath)

    for idx, paragraph in enumerate(document.paragraphs, start=1):
        text = paragraph.text.strip()
        if text:
            items.append((f"문단{idx}", text))

    return items


def _extract_doc(filepath: str) -> list[tuple[str, str]]:
    if olefile is None:
        raise RuntimeError("olefile 미설치")

    items: list[tuple[str, str]] = []

    with olefile.OleFileIO(filepath) as ole:
        if not ole.exists("WordDocument"):
            return []

        try:
            raw = ole.openstream("WordDocument").read()
        except OSError:
            return []

        candidates = _extract_doc_text_candidates(raw)
        for line_no, text in enumerate(candidates, start=1):
            items.append((f"DOC L{line_no}", text))

    return items


def _extract_hwp(filepath: str) -> list[tuple[str, str]]:
    if olefile is None:
        raise RuntimeError("olefile 미설치")

    items: list[tuple[str, str]] = []
    line_no = 1

    with olefile.OleFileIO(filepath) as ole:
        compressed = _is_hwp_compressed(ole)
        section_streams = sorted(_iter_hwp_section_streams(ole))

        for stream_name in section_streams:
            try:
                raw_bytes = ole.openstream(stream_name).read()
            except OSError as exc:
                logger.debug("HWP 섹션 읽기 실패: %s (%s)", stream_name, exc)
                continue

            content = _decompress_hwp_stream(raw_bytes, compressed)
            if not content:
                continue

            for paragraph_text in _parse_hwp_para_text(content):
                items.append((f"L{line_no}", paragraph_text))
                line_no += 1

        if items:
            return items

        # 섹션 파싱 실패 시 미리보기 텍스트 스트림으로 최소한의 텍스트를 시도한다.
        if ole.exists("PrvText"):
            preview = ole.openstream("PrvText").read().decode("utf-16le", errors="ignore")
            for line in preview.splitlines():
                cleaned = _clean_text(line)
                if cleaned:
                    items.append((f"L{line_no}", cleaned))
                    line_no += 1

    return items


def _extract_hwpx(filepath: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    line_no = 1

    with zipfile.ZipFile(filepath, "r") as archive:
        xml_files = [
            name
            for name in archive.namelist()
            if name.lower().endswith(".xml") and "section" in name.lower()
        ]
        if not xml_files:
            xml_files = [name for name in archive.namelist() if name.lower().endswith(".xml")]

        for xml_name in sorted(xml_files):
            try:
                xml_bytes = archive.read(xml_name)
                root = ET.fromstring(xml_bytes)
            except (KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
                logger.debug("HWPX XML 파싱 스킵: %s (%s)", xml_name, exc)
                continue

            for raw_text in root.itertext():
                cleaned = _clean_text(raw_text)
                if cleaned:
                    items.append((f"L{line_no}", cleaned))
                    line_no += 1

    return items


def _extract_csv(filepath: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []

    for encoding in TXT_FALLBACK_ENCODINGS:
        try:
            with open(filepath, "r", encoding=encoding, newline="") as file:
                reader = csv.reader(file)
                for row_idx, row in enumerate(reader, start=1):
                    for col_idx, value in enumerate(row, start=1):
                        cleaned = value.strip()
                        if cleaned:
                            items.append((f"R{row_idx}C{col_idx}", cleaned))
            return items
        except UnicodeDecodeError:
            items.clear()
            continue

    return items


def _extract_txt(filepath: str) -> list[tuple[str, str]]:
    for encoding in TXT_FALLBACK_ENCODINGS:
        try:
            with open(filepath, "r", encoding=encoding) as file:
                lines = file.readlines()
            return [
                (f"L{line_no}", line.strip())
                for line_no, line in enumerate(lines, start=1)
                if line.strip()
            ]
        except UnicodeDecodeError:
            continue

    return []


def _normalize_extensions(extensions: set[str]) -> set[str]:
    normalized: set[str] = set()
    for extension in extensions:
        if not extension:
            continue
        cleaned = extension.strip().lower()
        if not cleaned:
            continue
        if not cleaned.startswith("."):
            cleaned = f".{cleaned}"
        normalized.add(cleaned)
    return normalized


def _build_quick_check_patterns(keywords: list[str]) -> list[bytes]:
    patterns: list[bytes] = []
    seen: set[bytes] = set()

    for keyword in keywords:
        text_variants = [keyword, keyword.lower(), keyword.upper()]
        for encoding in QUICK_CHECK_ENCODINGS:
            for text in text_variants:
                try:
                    encoded = text.encode(encoding)
                except UnicodeEncodeError:
                    continue

                if not encoded or encoded in seen:
                    continue
                seen.add(encoded)
                patterns.append(encoded)

    return patterns


def _set_search_file_meta(filepath: str, skipped: bool, failed: bool) -> None:
    normalized = os.path.normcase(os.path.abspath(filepath))
    with _SEARCH_FILE_META_LOCK:
        _SEARCH_FILE_META[normalized] = (skipped, failed)


def _normalize_keyword_list(keywords: list[str]) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()

    for keyword in keywords:
        if not isinstance(keyword, str):
            continue
        cleaned = keyword.strip()
        if not cleaned:
            continue
        lowered = cleaned.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        normalized.append(cleaned)

    return normalized


def _is_excluded_dir(dir_name: str) -> bool:
    return dir_name.lower() in EXCLUDED_DIR_NAMES


def _is_hwp_compressed(ole: Any) -> bool:
    if not ole.exists("FileHeader"):
        return False

    file_header = ole.openstream("FileHeader").read()
    if len(file_header) < 40:
        return False

    flags = int.from_bytes(file_header[36:40], byteorder="little", signed=False)
    return bool(flags & 0x01)


def _iter_hwp_section_streams(ole: Any) -> list[str]:
    stream_names: list[str] = []
    for path_parts in ole.listdir(streams=True, storages=False):
        if len(path_parts) >= 2 and path_parts[0] == "BodyText" and path_parts[1].startswith("Section"):
            stream_names.append("/".join(path_parts))
    return stream_names


def _decompress_hwp_stream(data: bytes, compressed: bool) -> bytes:
    if not compressed:
        return data

    try:
        return zlib.decompress(data, -15)
    except zlib.error:
        try:
            return zlib.decompress(data)
        except zlib.error:
            return b""


def _parse_hwp_para_text(data: bytes) -> list[str]:
    paragraphs: list[str] = []
    offset = 0
    data_len = len(data)

    while offset + 4 <= data_len:
        header = int.from_bytes(data[offset : offset + 4], byteorder="little", signed=False)
        offset += 4

        record_type = header & 0x3FF
        record_size = (header >> 20) & 0xFFF

        if record_size == 0xFFF:
            if offset + 4 > data_len:
                break
            record_size = int.from_bytes(data[offset : offset + 4], byteorder="little", signed=False)
            offset += 4

        end_offset = offset + record_size
        if end_offset > data_len:
            break

        payload = data[offset:end_offset]
        offset = end_offset

        if record_type != 67 or not payload:
            continue

        decoded = payload.decode("utf-16le", errors="ignore")
        cleaned = _clean_text(decoded)
        if cleaned:
            paragraphs.append(cleaned)

    return paragraphs


def _clean_text(text: str) -> str:
    cleaned = text.replace("\x00", " ")
    cleaned = re.sub(r"[\x01-\x08\x0B\x0C\x0E-\x1F]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _extract_doc_text_candidates(raw: bytes) -> list[str]:
    candidates: list[str] = []
    seen: set[str] = set()

    def collect(decoded_text: str) -> None:
        for token in re.split(r"[\r\n\t]+", decoded_text):
            cleaned = _clean_text(token)
            if len(cleaned) < 2:
                continue
            if not re.search(r"[0-9A-Za-z가-힣]", cleaned):
                continue
            lowered = cleaned.lower()
            if lowered in seen:
                continue
            seen.add(lowered)
            candidates.append(cleaned)

    # WordDocument 스트림은 바이너리이므로 UTF-16LE/CP949 등으로 후보 문자열을 추출한다.
    collect(raw.decode("utf-16le", errors="ignore"))
    collect(raw.decode("cp949", errors="ignore"))
    collect(raw.decode("latin-1", errors="ignore"))

    return candidates
