"""
Daily Report Bot - Social Eagle Gen AI Architect Program, Week 1 Day 3 (Assignment 1).

Uses PyAutoGUI to control the real mouse and keyboard like a person would:
  1. Opens Chrome on a public site (BBC News) and copies the top headline.
  2. Opens a spreadsheet and writes one row: date & time, the fetched data, a comment.
  3. Saves the file as daily_report_YYYY-MM-DD.xlsx.
  4. Takes a screenshot of the final sheet.

Do NOT move the mouse or type while the bot runs - it is controlling them.
"""

import subprocess
import sys
import time
import webbrowser
from datetime import datetime

import pyautogui
import pyperclip
from openpyxl import Workbook
from openpyxl import load_workbook

SOURCE_URL = "https://www.bbc.com/news"
COMMENT = "Top headline of the morning - filed automatically."


def fetch_headline_with_browser(url: str) -> str:
    """Open Chrome, select the page text around the top headline and copy it."""
    print(f"Opening {url} in Chrome ...")
    webbrowser.open(url)  # default browser; keep Chrome your default for this run
    time.sleep(8)  # let the page load; slow networks need a moment

    # Ctrl+A then Ctrl+C grabs the visible page text; we keep the first
    # meaningful line, which on BBC News is the top headline block.
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(1)
    page_text = pyperclip.paste()

    headline = "Could not read headline"
    for line in page_text.splitlines():
        line = line.strip()
        # first line long enough to be a headline, not a nav label
        if len(line) > 30:
            headline = line
            break
    print("Fetched:", headline)
    return headline


def write_report_row(headline: str) -> str:
    """Create/append the daily report workbook and return its filename."""
    today = datetime.now().strftime("%Y-%m-%d")
    filename = f"daily_report_{today}.xlsx"

    try:
        wb = load_workbook(filename)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Daily Report"
        ws.append(["Date & Time", "Fetched Data", "Comment"])

    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), headline, COMMENT])
    wb.save(filename)
    print("Saved row to", filename)
    return filename


def open_and_screenshot(filename: str) -> str:
    """Open the saved workbook like a person would and screenshot the sheet."""
    if sys.platform.startswith("win"):
        subprocess.Popen(["cmd", "/c", "start", "", filename], shell=False)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", filename])
    else:
        subprocess.Popen(["xdg-open", filename])
    time.sleep(8)  # let Excel / Numbers / Calc open and render

    shot = f"screenshot_{datetime.now().strftime('%Y-%m-%d')}.png"
    pyautogui.screenshot(shot)
    print("Screenshot saved:", shot)
    return shot


def main():
    headline = fetch_headline_with_browser(SOURCE_URL)
    filename = write_report_row(headline)
    open_and_screenshot(filename)
    print("Done. Files produced:")
    print(" -", filename)
    print(" - screenshot PNG of the final sheet")


if __name__ == "__main__":
    main()
