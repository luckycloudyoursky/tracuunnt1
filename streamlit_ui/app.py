from __future__ import annotations

import contextlib
import importlib
import logging
import re
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Dict, List, Optional

import streamlit as st

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tra_cuu_thue_batch_excel as batch

batch = importlib.reload(batch)

APP_TITLE = "Tra Cuu Thue Batch"
APP_BUILD = "build-20260414-requests-fallback-v4"
DEFAULT_OUTPUT_NAME = "ket_qua_tra_cuu_thue.xlsx"
DISPLAY_COLUMNS = ["STT", "MST", "Ten nguoi nop thue", "Dia chi", "Co quan thue quan ly", "Trang thai MST"]


class StreamlitLog:
    def __init__(self, placeholder, max_chars: int = 20000):
        self.placeholder = placeholder
        self.max_chars = max_chars
        self.parts: List[str] = []

    def append(self, text: str) -> None:
        if not text:
            return
        self.parts.append(text)
        rendered = "".join(self.parts)
        if len(rendered) > self.max_chars:
            rendered = rendered[-self.max_chars:]
        self.placeholder.code(rendered or " ", language="text")

    def write(self, text: str) -> None:
        self.append(text)

    def flush(self) -> None:
        pass


class StreamlitLogHandler(logging.Handler):
    def __init__(self, log: StreamlitLog):
        super().__init__()
        self.log = log

    def emit(self, record: logging.LogRecord) -> None:
        self.log.append(self.format(record) + "\n")


def extract_msts_from_text(text: str) -> List[str]:
    return re.findall(r"\d{10}(?:-\d{3})?", text or "")


def read_msts_from_upload(uploaded_file, input_column: Optional[str]) -> List[str]:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix not in {".txt", ".csv", ".xlsx", ".xlsm"}:
        raise ValueError("Input file must be .txt, .csv, .xlsx, or .xlsm")
    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = Path(temp_dir) / f"input{suffix}"
        input_path.write_bytes(uploaded_file.getvalue())
        return batch.read_msts_from_file(input_path, input_column or None)


def collect_msts(text_input: str, uploaded_file, input_column: Optional[str], dedupe: bool) -> List[str]:
    msts: List[str] = []
    msts.extend(extract_msts_from_text(text_input))
    if uploaded_file is not None:
        msts.extend(read_msts_from_upload(uploaded_file, input_column))
    if dedupe:
        msts = batch.unique_keep_order(msts)
    if not msts:
        raise ValueError("No valid MST values found")
    return msts


def result_rows(results: List[Dict[str, str]]) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    for index, result in enumerate(results, start=1):
        rows.append(dict(zip(DISPLAY_COLUMNS, batch.build_row(index, result))))
    return rows


def build_excel_bytes(results: List[Dict[str, str]]) -> bytes:
    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = Path(temp_dir) / DEFAULT_OUTPUT_NAME
        batch.write_excel(results, output_path)
        return output_path.read_bytes()


def run_batch_lookup(
    msts: List[str],
    retries: int,
    delay: float,
    debug: bool,
    status_placeholder,
    progress_bar,
    log: StreamlitLog,
) -> List[Dict[str, str]]:
    batch.set_log_level(debug)
    handler = StreamlitLogHandler(log)
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    handler.setLevel(logging.INFO if debug else logging.ERROR)
    logging.getLogger().addHandler(handler)
    results: List[Dict[str, str]] = []
    try:
        with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
            client_class, source_url = batch.load_gdt_client()
            batch.set_log_level(debug)
            total = len(msts)
            with client_class() as client:
                for index, mst in enumerate(msts, start=1):
                    status_placeholder.info(f"Running {index}/{total}: {mst}")
                    log.append(f"[{index}/{total}] Looking up MST {mst}...\n")
                    result = batch.lookup_mst_with_client(client, mst, retries, source_url, save_html_dir=None)
                    results.append(result)
                    if result.get("error"):
                        log.append(f"  -> ERROR: {result['error']}\n")
                    else:
                        log.append(f"  -> OK: {result.get('company_name', '')}\n")
                    progress_bar.progress(index / total)
                    if index < total and delay > 0:
                        import time
                        time.sleep(delay)
    finally:
        logging.getLogger().removeHandler(handler)
    return results


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="TT", layout="wide")
    st.title(APP_TITLE)
    with st.sidebar:
        st.caption(APP_BUILD)
        st.header("Options")
        retries = st.number_input("CAPTCHA retries", min_value=1, max_value=100, value=15, step=1)
        delay = st.number_input("Delay between MST", min_value=0.0, value=1.0, step=0.5)
        dedupe = st.checkbox("Remove duplicated MST", value=True)
        debug = st.checkbox("Debug log", value=False)
        input_column = st.text_input(
            "MST column in Excel",
            value="",
            help="Use a 1-based column index or header name. Leave blank to scan all cells.",
        ).strip()
    uploaded_file = st.file_uploader("MST file", type=["txt", "csv", "xlsx", "xlsm"])
    text_input = st.text_area("MST list", height=180, placeholder="0100109106\n0101243150\n0301446666")
    run_clicked = st.button("Run lookup", type="primary")
    if run_clicked:
        st.session_state.pop("lookup_results", None)
        st.session_state.pop("excel_bytes", None)
        try:
            msts = collect_msts(text_input, uploaded_file, input_column or None, dedupe)
        except Exception as exc:
            st.error(str(exc))
            st.stop()
        status_placeholder = st.empty()
        progress_bar = st.progress(0)
        log_placeholder = st.empty()
        log = StreamlitLog(log_placeholder)
        try:
            results = run_batch_lookup(msts, int(retries), float(delay), debug, status_placeholder, progress_bar, log)
            st.session_state["lookup_results"] = results
            st.session_state["excel_bytes"] = build_excel_bytes(results)
            status_placeholder.success(f"Done {len(results)}/{len(msts)} MST")
        except Exception:
            st.error("Lookup failed")
            st.code(traceback.format_exc(), language="text")
    results = st.session_state.get("lookup_results")
    excel_bytes = st.session_state.get("excel_bytes")
    if results:
        st.subheader("Results")
        st.dataframe(result_rows(results), use_container_width=True, hide_index=True)
        if excel_bytes:
            st.download_button(
                "Download Excel",
                data=excel_bytes,
                file_name=DEFAULT_OUTPUT_NAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
