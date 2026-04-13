#!/usr/bin/env python3
"""GDT taxpayer lookup client using Playwright + OCR."""

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


def _resolve_chromium_executable() -> Optional[str]:
    candidates = [
        os.environ.get("CHROME_BIN"),
        os.environ.get("CHROMIUM_BIN"),
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        shutil.which("google-chrome"),
        shutil.which("google-chrome-stable"),
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
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
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=False,
        timeout=180,
    )
    try:
        with open(marker, "w", encoding="utf-8") as handle:
            handle.write("ok")
    except OSError:
        pass


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
        self._ocr = None

    def __enter__(self):
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
        self.context = self.browser.new_context(user_agent=USER_AGENT)
        self.page = self.context.new_page()
        Stealth().apply_stealth_sync(self.page)
        logger.info("Playwright started with stealth enabled")

    def init_session(self) -> bool:
        try:
            self._start_playwright()
            logger.info("Navigating to %s", TAX_LOOKUP_URL)
            self.page.goto(TAX_LOOKUP_URL, wait_until="networkidle", timeout=60000)
            time.sleep(4)
            self.last_html = self.page.content()
            return True
        except Exception as exc:
            logger.error("Failed to initialize session: %s", exc)
            return False

    def get_captcha(self, save_path: Optional[str] = None) -> bytes:
        captcha_img = self.page.locator('img[src*="captcha.png"]')
        captcha_img.wait_for(state="visible", timeout=15000)
        image_bytes = captcha_img.screenshot()
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
            taxpayers.append(
                TaxPayerInfo(
                    tax_id=tax_id,
                    name=cols[2].get_text(separator=" ", strip=True),
                    address=cols[3].get_text(separator=" ", strip=True),
                    tax_office=cols[4].get_text(separator=" ", strip=True),
                    status=cols[5].get_text(separator=" ", strip=True),
                    branch_id=tax_id.split("-", 1)[1] if "-" in tax_id else None,
                )
            )
        return taxpayers

    def lookup_tax_id(self, tax_id: str, captcha: str, save_html_path: Optional[str] = "last_search_result.html") -> List[TaxPayerInfo]:
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

    def close(self) -> None:
        if self.browser:
            self.browser.close()
            self.browser = None
        if self.pw:
            self.pw.stop()
            self.pw = None
        logger.info("GDT client closed")
