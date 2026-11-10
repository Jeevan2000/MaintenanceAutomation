# Ferro Towers Maintenance Ledger Bot

The Telegram bot records maintenance payments as new rows in the workbook's existing ledger. It recognizes only `/mnt` commands; other messages are ignored.

## Telegram command

```text
/mnt <flat number> <amount> [date]
```

Examples:

- `/mnt 101 3000`
- `/mnt 202 4500 16-03-2026`

The optional date accepts `YYYY-MM-DD`, `DD-MM-YYYY`, or `DD/MM/YYYY`. If omitted, the bot uses the Telegram message timestamp converted to India time (`Asia/Kolkata`). Excel stores it as a date and displays it using Long Date format.

Send `/report` to receive a PNG image containing a table for all 11 flats. The image has a centered two-line title: “Ferro Tower” followed by “Maintenance Report”; the table header is bold with a light-blue background. For each flat it uses the latest worksheet entry and calculates the reported balance as that row's Balance plus ₹1,500 for each elapsed month after Paid Till. A negative balance covers whole months at ₹1,500 per month; only elapsed months beyond that credit count as pending. The balance heading includes the report date. Paid Till in the current or a future month therefore adds no monthly maintenance or overdue months. Report calculations do not modify the workbook.

Supported flats and owners:

| Flat | Owner |
| --- | --- |
| 101 | SANIDHYA  |
| 102 | SUNIL  |
| 201 | SANIDHYA  |
| 202 | SAMPATLAL  |
| 301 | SANDIP  |
| 302 | SWAWAND  |
| 401 | DIPAK  |
| 402 | VIJAY  |
| 501 | MACCHINDRA  |
| 502 | SUDHIR  |
| 601 | PRACHEER  |

## Ledger calculations

Each accepted payment is appended as a new row; earlier rows and the master roster are retained. The bot finds the latest payment row for that flat and calculates:

- **Pending Balance:** previous payment's Balance. For a flat's first appended payment, use the current Balance in its roster row; if that is blank, use its Pending Balance, or zero if both are blank.
- **Pending Balance, Amount, and Balance:** stored as numbers and displayed in Indian rupees (₹) with two decimal places. Balance is Pending Balance minus Amount; an overpayment is retained as a negative balance (credit).
- **Paid Date and Paid Till:** stored as dates and displayed in Excel Long Date format.
- **Paid Till:** a calendar date displayed in Excel's Long Date format. For an existing date, add the whole number of months covered by the payment (`floor(amount / 1500)`) to the prior Paid Till date, including carrying into the next year (for example, 01-Nov-2026 + 3 months = 01-Feb-2027). When no prior date exists, the payment's month is the first covered month, so the date advances by `months - 1` from the first day of that month. If the amount is less than ₹1,500, Paid Till remains unchanged. Non-multiples of 1500 are allowed; their remainder still affects Balance.

The workbook must contain the ledger columns `FLAT NO.`, `OWNERS`, `Pending Balance`, `Paid Date`, `Amount`, `Balance`, and `Paid Till` on one worksheet. The bot searches the workbook for this header row and leaves unrelated sheets untouched.

## Setup and run

Use Python 3.9 or later. From the project folder, create and activate a virtual environment named `.venv` (PowerShell):

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell prevents activation because script execution is disabled, allow scripts for this terminal session and activate again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

Install the project's dependencies into the active virtual environment:

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If a dependency is missing, make sure `.venv` is activated and install the package using its pip package name:

```powershell
python -m pip install <package-name>
```

For a dependency needed by the project, add it to `requirements.txt` as well so it can be installed the next time with `python -m pip install -r requirements.txt`. The tests also require `pytest`, which can be installed with `python -m pip install pytest`.

Create your local `.env` file from the example and add your Telegram bot token:

```powershell
Copy-Item .env.example .env
```

Edit `.env` and set `BOT_TOKEN`. The other settings shown in `.env.example` are optional. Keep `.venv` and `.env` local to your machine; do not copy the virtual environment between machines.

With `.venv` activated and `.env` configured, run:

```powershell
python run_bot.py FerroTowerMainteance.xlsx
```

If no workbook path is supplied, `FerroTowerMainteance.xlsx` in the current directory is used. `STATE_FILE` and `LOG_FILE` can optionally be set in `.env`; defaults are `read_state.json` and `messages.log`. The state file stores the last processed Telegram `update_id` (used to resume without re-reading handled messages) and `last_read_timestamp` (UTC time when that update was processed). The timestamp is preserved during idle polling.

The bot prints timestamped logs to the console and appends them to `messages.log` (or `LOG_FILE`). Logs show startup, each Telegram update/message received, command processing or rejection, successful workbook saves with the row number, acknowledgement status, and update-state completion. If the log file cannot be written, the error is still printed to the console.

Run tests with `pytest`.
