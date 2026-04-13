from __future__ import annotations

import argparse
import csv
import logging
import re
import sys
import time
import unicodedata
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="replace")

MST_PATTERN = re.compile(r"\d{10}(?:-\d{3})?")
DEFAULT_GDT_SOURCE = "https://tracuunnt.gdt.gov.vn/tcnnt/mstdn.jsp"

OUTPUT_COLUMNS = ["STT", "MST", "Ten nguoi nop thue", "Dia chi", "Co quan thue quan ly", "Trang thai MST"]
CAPTCHA_ERROR_MARKERS = ("nhap dung ma xac nhan", "ma xac nhan khong dung", "captcha khong dung")
NOT_FOUND_MARKERS = ("khong tim thay", "khong co thong tin")


def normalize_text(value: str) -> str:
    text = unicodedata.normalize("NFD", value or "")
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    text = re.sub(r"[^0-9A-Za-z]+", " ", text)
    return re.sub(r"\s+", " ", text).strip().lower()


def coerce_mst(value: object, numeric_from_excel: bool = False) -> Optional[str]:
    if value is None:
        return None
    if numeric_from_excel and isinstance(value, (int, float)):
        if isinstance(value, float) and not value.is_integer():
            return None
        text = str(int(value))
        if len(text) < 10:
            text = text.zfill(10)
    else:
        text = str(value).strip().replace("\ufeff", "")
    match = MST_PATTERN.search(re.sub(r"[^\d-]", "", text))
    return match.group(0) if match else None


def normalize_mst(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def unique_keep_order(values: Iterable[str]) -> List[str]:
    seen = set()
    result = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def read_text_or_csv(path: Path) -> List[str]:
    msts: List[str] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        if path.suffix.lower() == ".csv":
            reader = csv.reader(handle)
            for row in reader:
                for cell in row:
                    mst = coerce_mst(cell)
                    if mst:
                        msts.append(mst)
        else:
            for line in handle:
                mst = coerce_mst(line)
                if mst:
                    msts.append(mst)
    return msts


def resolve_excel_column(sheet, column: Optional[str]) -> Optional[int]:
    if not column:
        return None
    if column.isdigit():
        index = int(column)
        if index <= 0:
            raise ValueError("input column must be 1-based")
        return index
    target = normalize_text(column)
    for cell in next(sheet.iter_rows(min_row=1, max_row=1), []):
        if normalize_text(str(cell.value or "")) == target:
            return cell.column
    raise ValueError(f"Could not find Excel column header: {column}")


def read_excel(path: Path, input_column: Optional[str]) -> List[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    column_index = resolve_excel_column(sheet, input_column)
    msts: List[str] = []
    for row in sheet.iter_rows():
        cells = row if column_index is None else [row[column_index - 1]]
        for cell in cells:
            mst = coerce_mst(cell.value, numeric_from_excel=True)
            if mst:
                msts.append(mst)
    workbook.close()
    return msts


def read_msts_from_file(path: Path, input_column: Optional[str]) -> List[str]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return read_excel(path, input_column)
    if suffix in {".txt", ".csv"}:
        return read_text_or_csv(path)
    raise ValueError("Input file must be .txt, .csv, .xlsx, or .xlsm")


def load_gdt_client() -> Tuple[Any, str]:
    try:
        from api_client_v2 import GdtTaxLookupClientV2, TAX_LOOKUP_URL
    except ImportError as exc:
        raise RuntimeError("Could not import api_client_v2") from exc
    return GdtTaxLookupClientV2, TAX_LOOKUP_URL or DEFAULT_GDT_SOURCE


def set_log_level(debug: bool) -> None:
    level = logging.INFO if debug else logging.ERROR
    logging.getLogger().setLevel(level)
    logging.getLogger("api_client_v2").setLevel(level)


def html_has_marker(html: Optional[str], markers: Iterable[str]) -> bool:
    normalized = normalize_text(html or "")
    return any(marker in normalized for marker in markers)


def select_taxpayer(mst: str, taxpayers: List[Any]) -> Optional[Any]:
    if not taxpayers:
        return None
    target = normalize_mst(mst)
    for taxpayer in taxpayers:
        if normalize_mst(getattr(taxpayer, "tax_id", "")) == target:
            return taxpayer
    return taxpayers[0]


def taxpayer_to_result(input_mst: str, taxpayer: Any, count: int, source_url: str) -> Dict[str, str]:
    return {
        "input_mst": input_mst,
        "mst": getattr(taxpayer, "tax_id", input_mst),
        "company_name": getattr(taxpayer, "name", ""),
        "address": getattr(taxpayer, "address", ""),
        "managed_by": getattr(taxpayer, "tax_office", ""),
        "status": getattr(taxpayer, "status", ""),
        "source_url": source_url,
    }


def lookup_mst_with_client(client: Any, mst: str, retries: int, source_url: str, save_html_dir: Optional[str] = None) -> Dict[str, str]:
    try:
        if not client.init_session():
            return {"input_mst": mst, "mst": mst, "error": "Could not initialize GDT session", "source_url": source_url}
    except Exception as exc:
        return {"input_mst": mst, "mst": mst, "error": f"Could not initialize GDT session: {exc}", "source_url": source_url}

    for attempt in range(1, retries + 1):
        try:
            captcha = client.auto_solve_captcha()
            print(f"  CAPTCHA attempt {attempt}/{retries}: {captcha}")
            taxpayers = client.lookup_tax_id(mst, captcha, save_html_path=None)
            selected = select_taxpayer(mst, taxpayers)
            if selected:
                return taxpayer_to_result(mst, selected, len(taxpayers), source_url)
            html = getattr(client, "last_html", "")
            if html_has_marker(html, NOT_FOUND_MARKERS):
                return {"input_mst": mst, "mst": mst, "error": "Not found", "source_url": source_url}
            if attempt < retries:
                reason = "wrong CAPTCHA" if html_has_marker(html, CAPTCHA_ERROR_MARKERS) else "no result yet"
                print(f"  -> {reason}, retrying...")
                time.sleep(1)
                continue
            return {"input_mst": mst, "mst": mst, "error": "No result after CAPTCHA retries", "source_url": source_url}
        except Exception as exc:
            if attempt == retries:
                return {"input_mst": mst, "mst": mst, "error": str(exc), "source_url": source_url}
            print(f"  -> attempt {attempt}/{retries} error: {exc}. Retrying...")
            time.sleep(1)
            try:
                client.init_session()
            except Exception:
                pass
    return {"input_mst": mst, "mst": mst, "error": "Lookup failed", "source_url": source_url}


def build_row(index: int, result: Dict[str, str]) -> List[object]:
    status = result.get("status") or result.get("error") or ""
    return [index, result.get("mst") or result.get("input_mst", ""), result.get("company_name", ""), result.get("address", ""), result.get("managed_by", ""), status]


def write_excel(results: List[Dict[str, str]], output_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Tra cuu thue"
    sheet.append(OUTPUT_COLUMNS)
    for index, result in enumerate(results, start=1):
        sheet.append(build_row(index, result))
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_alignment = Alignment(vertical="top", wrap_text=True)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = body_alignment
    widths = {"A": 8, "B": 18, "C": 42, "D": 58, "E": 36, "F": 28}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(output_path)


def get_msts(args: argparse.Namespace) -> List[str]:
    raise NotImplementedError("CLI input parsing is not used by the Streamlit app")
