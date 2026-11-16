"""Screenshots of KRONOS Offline running the three demo recipes, for the Word guide (D88).

    python packaging/build_offline.py
    python docs/demo/take_screenshots.py [--browser <chromium executable>] [--out <folder>]

It unpacks the built zip, starts its own server/serve.py on localhost:5181 as the launcher would, and drives the page
with Playwright (packaging/offline_test/, installed only where screenshots are made): for each recipe it chooses
NCR_Register_Demo.xlsx, pastes the recipe, checks it, picks the sheet and makes the changes. It saves the parts of the
page the guide shows. Nothing here runs in CI.
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--browser", help="a Chromium executable already on this computer")
    ap.add_argument("--out", default=str(HERE / "shots"))
    args = ap.parse_args()
    info = json.loads((ROOT / "build" / "offline" / "build-info.json").read_text(encoding="utf-8"))
    folder = Path(tempfile.mkdtemp(prefix="kronos demo "))
    with zipfile.ZipFile(ROOT / "build" / "offline" / info["zip"]) as archive:
        archive.extractall(folder)
    app = folder / "KRONOS-Offline"
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    log = open(folder / "server.log", "wb")
    server = subprocess.Popen([sys.executable, os.path.join("server", "serve.py")], cwd=app, stdout=log, stderr=log,
                              stdin=subprocess.DEVNULL)
    try:
        for _ in range(50):
            try:
                urllib.request.urlopen("http://localhost:5181/", timeout=2)
                break
            except OSError:
                time.sleep(0.2)
        from playwright.sync_api import sync_playwright
        data = (HERE / "NCR_Register_Demo.xlsx").read_bytes()
        with sync_playwright() as p:
            browser = p.chromium.launch(**({"executable_path": args.browser} if args.browser else {}))
            for n, name in enumerate(("recipe_1_highlight.vcl", "recipe_2_tidy.vcl", "recipe_3_summary.vcl"), start=1):
                page = browser.new_page(viewport={"width": 1100, "height": 900}, device_scale_factor=2)
                page.goto("http://localhost:5181/")
                page.wait_for_function("document.getElementById('python').textContent.indexOf('ready') >= 0",
                                       timeout=180000)
                page.set_input_files("#book-file", files={"name": "NCR_Register_Demo.xlsx", "mimeType": XLSX,
                                                          "buffer": data})
                page.wait_for_selector("#book-info:not([hidden])", timeout=180000)
                if n == 1:
                    page.locator("#book-card").screenshot(path=str(out / "workbook.png"))
                page.fill("#recipe", (HERE / name).read_text(encoding="utf-8"))
                page.click("#check")
                page.wait_for_selector("#banner", timeout=180000)
                page.wait_for_selector("#apply-card:not([hidden])", timeout=180000)
                page.locator("#result").screenshot(path=str(out / f"recipe{n}-proven.png"))
                page.select_option("#arg-0" if n < 3 else "#arg-1", "NCR Register")
                page.locator("#apply-card").screenshot(path=str(out / f"recipe{n}-apply.png"))
                page.click("#apply")
                page.wait_for_selector("#applied-banner", timeout=180000)
                page.locator("#applied").screenshot(path=str(out / f"recipe{n}-copy.png"))
                print(n, page.text_content("#applied-banner"))
                page.close()
            browser.close()
    finally:
        server.terminate()
        server.wait(timeout=10)
        log.close()


if __name__ == "__main__":
    main()
