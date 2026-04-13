#!/usr/bin/env python3
"""GDT taxpayer lookup client using Playwright, with a requests fallback for cloud hosts."""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from typing import List, Optional

try:
    import ddddocr
    import requests
    from bs4 import BeautifulSoup
    from playwright.sync_api import sync_playwright
    from playwright_stealth.stealth import Stealth
except ImportError as exc:
    print(f"Required dependency missing: {exc}")
    print("Install with: pip install -r requirements.txt")
    sys.exit(1)

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

BASE_URL = "https://tracuunnt.gdt.gov.vn"
TAX_LOOKUP_URL = f"{BASE_URL}/tcnnt/mstdn.jsp"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/133.0.0.0 Safari/537.36"
)


def _prefer_requests_first() -> bool:
    mode = os.environ.get("GDT_LOOKUP_MODE", "").strip().lower()
    if mode in {"requests", "http"}:
        return True
    if mode in {"playwright", "browser"}:
        return False
    current_file = os.path.abspath(__file__).replace("\\", "/")
    return os.name != "nt" and (
        current_file.startswith("/mount/src/")
        or os.environ.get("HOME") == "/home/appuser"
    )


def _resolve_chromium_executable() -> Optional[str]:
    candidates = [
        os.environ.get("CHROME_BIN"), os.environ.get("CHROMIUM_BIN"), shutil.which("chromium"),
        shutil.which("chromium-browser"), shutil.which("google-chrome"), shutil.which("google-chrome-stable"),
        "/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def _install_playwright_chromium_once() -> None:
    marker = "/tmp/tracuunnt_playwright_chromium_installed"
    if os.name == "nt" or os.path.exists(marker):
        return
    logger.info("Installing Playwright Chromium browser")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False, timeout=180)
    try:
        with open(marker, "w", encoding="utf-8") as handle:
            handle.write("ok")
    except OSError:
        pass


def _decode_response(response: requests.Response) -> str:
    if not response.encoding or response.encoding.lower() == "iso-8859-1":
        response.encoding = "utf-8"
    return response.text


@dataclass
class TaxPayerInfo:
    tax_id: str
    name: str
    address: str
    tax_office: str
    status: str
    branch_id: Optional[str] = None


class GdtTaxLookupClientV2:
    def __init__(self):
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None
        self.last_html: Optional[str] = None
        self.last_error: Optional[str] = None
        self._ocr = None
        self._requests_session: Optional[requests.Session] = None
        self._use_requests_fallback = False

    def __enter__(self):
        if not _prefer_requests_first():
            self._start_playwright()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def _launch_browser(self):
        launch_options = {"headless": True}
        chromium_path = _resolve_chromium_executable()
        if chromium_path:
            launch_options["executable_path"] = chromium_path
        if os.name != "nt":
            launch_options["args"] = ["--no-sandbox", "--disable-dev-shm-usage"]
        return self.pw.chromium.launch(**launch_options)

    def _start_playwright(self) -> None:
        if self.pw:
            return
        self.pw = sync_playwright().start()
        try:
            self.browser = self._launch_browser()
        except Exception as exc:
            logger.warning("Initial Chromium launch failed: %s", exc)
            _install_playwright_chromium_once()
            self.browser = self._launch_browser()
        self.context = self.browser.new_context(user_agent=USER_AGENT, ignore_https_errors=True)
        self.context.set_default_timeout(45000)
        self.context.set_default_navigation_timeout(60000)
        self.page = self.context.new_page()
        Stealth().apply_stealth_sync(self.page)
        logger.info("Playwright started with stealth enabled")

    def init_session(self) -> bool:
        if _prefer_requests_first():
            logger.info("Using requests-first mode for GDT session")
            if self._init_requests_session():
                return True
            logger.warning("Requests-first mode failed, trying Playwright fallback")
        try:
            self._use_requests_fallback = False
            self.last_error = None
            self._start_playwright()
            logger.info("Navigating to %s", TAX_LOOKUP_URL)
            self.page.goto(TAX_LOOKUP_URL, wait_until="domcontentloaded", timeout=15000)
            time.sleep(4)
            self.page.locator('input[name="mst"]').wait_for(state="attached", timeout=15000)
            self.page.locator('img[src*="captcha.png"]').wait_for(state="attached", timeout=15000)
            self.last_html = self.page.content()
            return True
        except Exception as exc:
            try:
                self.last_html = self.page.content() if self.page else None
            except Exception:
                self.last_html = None
            current_url = self.page.url if self.page else ""
            self.last_error = f"{type(exc).__name__}: {exc}"
            if current_url:
                self.last_error = f"{self.last_error} (url={current_url})"
            logger.error("Failed to initialize Playwright session: %s", self.last_error)
            logger.warning("Falling back to requests-based GDT session")
            return self._init_requests_session(self.last_error)

    def _init_requests_session(self, previous_error: Optional[str] = None) -> bool:
        try:
            if self._requests_session:
                self._requests_session.close()
            self._use_requests_fallback = True
            self._requests_session = requests.Session()
            self._requests_session.headers.update({
                "User-Agent": USER_AGENT,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
                "Referer": TAX_LOOKUP_URL,
            })
            response = self._requests_session.get(TAX_LOOKUP_URL, timeout=(8, 20))
            response.raise_for_status()
            self.last_html = _decode_response(response)
            if 'name="mst"' not in self.last_html or "captcha" not in self.last_html.lower():
                raise RuntimeError("GDT response did not contain lookup form/captcha")
            self.last_error = None
            logger.info("Requests-based GDT session initialized")
            return True
        except Exception as exc:
            fallback_error = f"{type(exc).__name__}: {exc}"
            if previous_error:
                fallback_error = f"playwright={previous_error}; requests={fallback_error}"
            self.last_error = fallback_error
            logger.error("Failed to initialize requests fallback session: %s", self.last_error)
            return False

    def get_captcha(self, save_path: Optional[str] = None) -> bytes:
        if self._use_requests_fallback:
            return self._get_captcha_requests(save_path)
        captcha_img = self.page.locator('img[src*="captcha.png"]')
        captcha_img.wait_for(state="visible", timeout=15000)
        image_bytes = captcha_img.screenshot()
        if save_path:
            with open(save_path, "wb") as handle:
                handle.write(image_bytes)
        return image_bytes

    def _get_captcha_requests(self, save_path: Optional[str] = None) -> bytes:
        if not self._requests_session:
            raise RuntimeError("Requests session is not initialized")
        response = self._requests_session.get(
            f"{BASE_URL}/tcnnt/captcha.png",
            params={"uid": str(int(time.time() * 1000))},
            timeout=(8, 20),
        )
        response.raise_for_status()
        image_bytes = response.content
        if save_path:
            with open(save_path, "wb") as handle:
                handle.write(image_bytes)
        return image_bytes

    def auto_solve_captcha(self, image_bytes: Optional[bytes] = None) -> str:
        if not self._ocr:
            logger.info("Initializing OCR engine")
            self._ocr = ddddocr.DdddOcr(show_ad=False)
        if image_bytes is None:
            image_bytes = self.get_captcha()
        result = self._ocr.classification(image_bytes)
        logger.info("Auto-solved CAPTCHA: %s", result)
        return result

    def _parse_response_bs4(self, html_content: str) -> List[TaxPayerInfo]:
        taxpayers: List[TaxPayerInfo] = []
        soup = BeautifulSoup(html_content, "html.parser")
        result_container = soup.find(id="resultContainer")
        table = result_container.find("table") if result_container else None
        if table is None:
            table = soup.find("table", class_="table_gdt") or soup.find("table")
        if table is None:
            return taxpayers
        for row in table.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) < 6 or not cols[0].get_text(strip=True).isdigit():
                continue
            tax_id = cols[1].get_text(strip=True)
            taxpayers.append(TaxPayerInfo(
                tax_id=tax_id,
                name=cols[2].get_text(separator=" ", strip=True),
                address=cols[3].get_text(separator=" ", strip=True),
                tax_office=cols[4].get_text(separator=" ", strip=True),
                status=cols[5].get_text(separator=" ", strip=True),
                branch_id=tax_id.split("-", 1)[1] if "-" in tax_id else None,
            ))
        return taxpayers

    def lookup_tax_id(self, tax_id: str, captcha: str, save_html_path: Optional[str] = "last_search_result.html") -> List[TaxPayerInfo]:
        if self._use_requests_fallback:
            return self._lookup_tax_id_requests(tax_id, captcha, save_html_path)
        if not tax_id or not captcha:
            raise ValueError("Tax ID and CAPTCHA are required")
        logger.info("Submitting lookup for MST: %s", tax_id)
        self.page.fill('input[name="mst"]', tax_id.strip())
        self.page.fill('input[name="captcha"]', captcha.strip())
        self.page.locator("input.subBtn").click()
        time.sleep(3)
        self.last_html = self.page.content()
        if save_html_path:
            with open(save_html_path, "w", encoding="utf-8") as handle:
                handle.write(self.last_html)
        return self._parse_response_bs4(self.last_html)

    def _lookup_tax_id_requests(self, tax_id: str, captcha: str, save_html_path: Optional[str] = "last_search_result.html") -> List[TaxPayerInfo]:
        if not tax_id or not captcha:
            raise ValueError("Tax ID and CAPTCHA are required")
        if not self._requests_session:
            raise RuntimeError("Requests session is not initialized")
        logger.info("Submitting requests fallback lookup for MST: %s", tax_id)
        response = self._requests_session.post(
            TAX_LOOKUP_URL,
            data={"cm": "cm", "mst": tax_id.strip(), "fullname": "", "address": "", "cmt": "", "captcha": captcha.strip()},
            timeout=(8, 30),
        )
        response.raise_for_status()
        self.last_html = _decode_response(response)
        if save_html_path:
            with open(save_html_path, "w", encoding="utf-8") as handle:
                handle.write(self.last_html)
        return self._parse_response_bs4(self.last_html)

    def close(self) -> None:
        if self.browser:
            self.browser.close()
            self.browser = None
        if self.pw:
            self.pw.stop()
            self.pw = None
        if self._requests_session:
            self._requests_session.close()
            self._requests_session = None
        logger.info("GDT client closed")
