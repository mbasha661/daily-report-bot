# Daily Report Bot (PyAutoGUI) - Social Eagle Week 1 Day 3

A PyAutoGUI bot that prepares a daily status report like a person would:
it opens Chrome on a public news site, copies the top headline, writes one row
(date & time, the headline, a short comment) into an Excel file named with
today's date, and saves a screenshot of the final sheet.

## Setup

```bash
mkdir daily-report-bot && cd daily-report-bot
python -m venv venv
source venv/bin/activate          # Windows PowerShell: venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Linux only: `sudo apt install scrot` (PyAutoGUI screenshots need it).

## Run

```bash
python daily_report_bot.py
```

Do not touch the mouse or keyboard while it runs. When it finishes you will have:

- `daily_report_YYYY-MM-DD.xlsx` - one new row: date & time, headline, comment
- `screenshot_YYYY-MM-DD.png` - screenshot of the final sheet

## Notes / decisions

- The date, time and filename are generated at run time, never typed by hand.
- The bot waits with timed pauses after opening Chrome and the spreadsheet so
  slow app launches do not break the clicks.
- `openpyxl` is used to write the row safely; PyAutoGUI then opens the real
  spreadsheet app for the screenshot, which is the part a person would do.
- To film the submission video, record the screen from `python daily_report_bot.py`
  until the screenshot is saved (30-60 seconds).
