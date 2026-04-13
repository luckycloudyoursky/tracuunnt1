#!/usr/bin/env python3
"""GDT taxpayer lookup client using Playwright and OCR."""

from __future__ import annotations

import logging
import os
import shutil
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
    print("Install with: pip install beautifulsoup4 pillow playwright playwright-stealth ddddocr")
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


@dataclass
class TaxPayerInfo:
    tax_id: str
    name: str
    address: str
    tax_office: str
    status: str
    branch_id: Optional[str] = None

    def __str__(self) -> str:
        branch = f"\nBranch ID: {self.branch_id}" if self.branch_id else ""
        return (
            f"Tax ID: {self.tax_id}\n"
            f"Name: {self.name}\n"
            f"Address: {self.address}\n"
            f"Tax Office: {self.tax_office}\n"
            f"Status: {self.status}"
            f"{branch}"
        )


class GdtTaxLookupClientV2:
    def __init__(self) -> None:
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None
        self.last_html: Optional[str] = None
        self.last_error: Optional[str] = None
        self._ocr = None

    def __enter__(self) -> "GdtTaxLookupClientV2":
        self._start_playwright()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
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
        self.browser = self._launch_browser()
        self.context = self.browser.new_context(user_agent=USER_AGENT, ignore_https_errors=True)
        self.context.set_default_timeout(45000)
        self.context.set_default_navigation_timeout(45000)
        self.page = self.context.new_page()
        Stealth().apply_stealth_sync(self.page)
        logger.info("Playwright started with stealth enabled")

    def init_session(self) -> bool:
        try:
            self.last_error = None
            self._start_playwright()
            logger.info("Navigating to %s", TAX_LOOKUP_URL)
            self.page.goto(TAX_LOOKUP_URL, wait_until="domcontentloaded", timeout=45000)
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
            logger.error("Failed to initialize session: %s", self.last_error)
            return False

    def get_captcha(self, save_path: Optional[str] = None) -> bytes:
        captcha_img = self.page.locator('img[src*="captcha.png"]')
        captcha_img.wait_for(state="visible", timeout=10000)
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

    def lookup_tax_id(
        self,
        tax_id: str,
        captcha: str,
        save_html_path: Optional[str] = "last_search_result.html",
    ) -> List[TaxPayerInfo]:
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


if __name__ == "__main__":
    with GdtTaxLookupClientV2() as client:
        if client.init_session():
            client.get_captcha("test_captcha.png")
            print("CAPTCHA saved to test_captcha.png")
            mst = input("MST: ")
            code = input("CAPTCHA code: ")
            for result in client.lookup_tax_id(mst, code):
                print(f"---\n{result}")
