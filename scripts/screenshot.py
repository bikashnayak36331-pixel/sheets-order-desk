"""Render docs/*.html to docs/images/*.png with Playwright (pip install playwright)."""
from pathlib import Path

from playwright.sync_api import sync_playwright

docs = Path(__file__).resolve().parent.parent / "docs"
jobs = [("sample-invoice.html", "invoice.png", 820), ("sheet-preview.html", "sheet-preview.png", 1500)]
with sync_playwright() as p:
    browser = p.chromium.launch()
    for src, out, width in jobs:
        page = browser.new_page(viewport={"width": width, "height": 600}, device_scale_factor=1.5)
        page.route("http*://**/*", lambda route: route.abort())
        page.goto((docs / src).as_uri())
        page.screenshot(path=str(docs / "images" / out), full_page=True)
        page.close()
    browser.close()
print("Saved docs/images/invoice.png and docs/images/sheet-preview.png")
