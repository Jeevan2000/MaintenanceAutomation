from datetime import date, datetime, timezone
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from PIL import Image

from maintenance_utils import OWNER_MASTER, append_payment, get_maintenance_report
import run_bot
from run_bot import parse_mnt_command, render_maintenance_report


def make_workbook(path: Path):
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Sheet1"
    ws["A1"] = "FEERO TOWERS"
    headers = ["FLAT NO.", "OWNERS", "Pending Balance", "Paid Date", "Amount", "Balance", "Paid Till"]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=3, column=col, value=header)
    for row, (flat, owner) in enumerate(OWNER_MASTER.items(), start=4):
        ws.cell(row=row, column=1, value=int(flat))
        ws.cell(row=row, column=2, value=owner)
        ws.cell(row=row, column=6).number_format = r'"₹"\ #,##0.00'
        ws.cell(row=row, column=7).number_format = r"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"
    # Opening values exercise the first-payment path.
    ws.cell(row=4, column=6, value=20000)
    ws.cell(row=4, column=7, value=datetime(2026, 1, 1))
    other = workbook.create_sheet("Other")
    other["A1"] = "preserve me"
    workbook.save(path)


def test_parse_command_with_amount_and_explicit_date():
    parsed = parse_mnt_command("/mnt 101 3,000 16-03-2026")
    assert parsed == {"flat": "101", "amount": 3000, "date": date(2026, 3, 16)}


def test_parse_command_defaults_to_telegram_date_in_india_timezone():
    sent_at = datetime(2026, 10, 16, 20, 0, tzinfo=timezone.utc)
    parsed = parse_mnt_command("/mnt@FerroTowerBot 102 1500", sent_at)
    assert parsed["date"] == date(2026, 10, 17)


def test_parse_ignores_non_command_and_rejects_invalid_inputs():
    assert parse_mnt_command("101 paid 1500") is None
    for text in ("/mnt 101 0", "/mnt 999 1500", "/mnt 101 1500 31-02-2026", "/mnt 101"):
        with pytest.raises(ValueError):
            parse_mnt_command(text)


def test_append_carries_balance_and_accumulates_whole_paid_months(tmp_path: Path):
    file_path = tmp_path / "ledger.xlsx"
    make_workbook(file_path)

    first = append_payment(file_path, "101", 3000, date(2026, 3, 16))
    second = append_payment(file_path, "101", 4500, date(2026, 4, 16))

    assert (first["pending"], first["balance"], first["paid_till"]) == (20000, 17000, datetime(2026, 3, 1))
    assert (second["pending"], second["balance"], second["paid_till"]) == (17000, 12500, datetime(2026, 6, 1))
    workbook = load_workbook(file_path, data_only=False)
    ws = workbook["Sheet1"]
    assert ws.cell(first["row"], 4).value == datetime(2026, 3, 16)
    assert ws.cell(first["row"], 4).number_format == r"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"
    assert ws.cell(first["row"], 1).value == 101
    assert ws.cell(second["row"], 3).value == 17000
    assert ws.cell(first["row"], 3).number_format == r'"₹"\ #,##0.00'
    assert ws.cell(first["row"], 5).number_format == r'"₹"\ #,##0.00'
    assert ws.cell(first["row"], 6).number_format == r'"₹"\ #,##0.00'
    assert ws.cell(first["row"], 7).value == datetime(2026, 3, 1)
    assert ws.cell(first["row"], 7).number_format == r"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"
    assert workbook["Other"]["A1"].value == "preserve me"
    assert ws.cell(4, 6).value == 20000


def test_paid_till_advances_across_calendar_year(tmp_path: Path):
    file_path = tmp_path / "ledger.xlsx"
    make_workbook(file_path)
    workbook = load_workbook(file_path)
    workbook["Sheet1"].cell(4, 7).value = datetime(2026, 11, 1)
    workbook.save(file_path)

    result = append_payment(file_path, "101", 4500, date(2026, 11, 15))

    assert result["paid_till"] == datetime(2027, 2, 1)
    saved = load_workbook(file_path, data_only=False)["Sheet1"]
    assert saved.cell(result["row"], 7).value == datetime(2027, 2, 1)


def test_first_payment_defaults_blank_roster_balance_to_zero_and_allows_credit(tmp_path: Path):
    file_path = tmp_path / "ledger.xlsx"
    make_workbook(file_path)
    workbook = load_workbook(file_path)
    ws = workbook["Sheet1"]
    ws.cell(5, 3).value = None
    ws.cell(5, 6).value = None
    ws.cell(5, 7).value = None
    workbook.save(file_path)

    result = append_payment(file_path, "102", 4000, date(2026, 3, 16))
    assert (result["pending"], result["balance"], result["paid_till"]) == (0, -4000, datetime(2026, 4, 1))


def test_partial_month_payment_keeps_cents_and_adds_only_whole_months(tmp_path: Path):
    file_path = tmp_path / "ledger.xlsx"
    make_workbook(file_path)
    workbook = load_workbook(file_path)
    workbook["Sheet1"].cell(4, 6).value = 0.5
    workbook.save(file_path)

    result = append_payment(file_path, "101", "0.30", date(2026, 3, 16))
    assert result["balance"] == 0.2
    assert result["paid_till"] == datetime(2026, 1, 1)


def test_unknown_flat_is_not_appended(tmp_path: Path):
    file_path = tmp_path / "ledger.xlsx"
    make_workbook(file_path)
    with pytest.raises(ValueError, match="Unknown flat number"):
        append_payment(file_path, "777", 1500, date(2026, 3, 16))


def test_append_log_writes_to_console_and_log_file(tmp_path: Path, monkeypatch, capsys):
    log_path = tmp_path / "bot.log"
    monkeypatch.setattr(run_bot, "LOG_FILE", str(log_path))

    run_bot.append_log("Workbook updated successfully")

    assert "Workbook updated successfully" in capsys.readouterr().out
    assert "Workbook updated successfully" in log_path.read_text(encoding="utf-8")


def test_save_state_preserves_last_read_timestamp_and_update_id(tmp_path: Path, monkeypatch):
    state_path = tmp_path / "read_state.json"
    monkeypatch.setattr(run_bot, "STATE_FILE", str(state_path))
    state = {
        "update_id": 123,
        "last_read_timestamp": "2026-10-17T18:08:32+00:00",
    }

    run_bot.save_state(state)

    saved_state = run_bot.load_state()
    assert saved_state["update_id"] == 123
    assert saved_state["last_read_timestamp"] == "2026-10-17T18:08:32+00:00"
    assert "last_checked" in saved_state


def test_report_uses_latest_row_and_accrues_monthly_maintenance(tmp_path: Path):
    file_path = tmp_path / "ledger.xlsx"
    make_workbook(file_path)
    workbook = load_workbook(file_path)
    ws = workbook["Sheet1"]
    # Flat 102 has only a roster entry: use its existing Balance as the fallback.
    ws.cell(5, 6).value = 6000
    ws.cell(5, 7).value = datetime(2026, 1, 1)
    ws.cell(7, 6).value = -4500
    ws.cell(7, 7).value = datetime(2026, 1, 1)
    # Two entries for flat 101 prove the report uses Balance from its latest row.
    ws.append([101, OWNER_MASTER["101"], 9000, datetime(2026, 8, 1), 3000, 6000, datetime(2026, 8, 1)])
    ws.append([101, OWNER_MASTER["101"], 12345, datetime(2026, 9, 1), 1500, 10845, datetime(2026, 9, 1)])
    workbook.save(file_path)

    report = get_maintenance_report(file_path, date(2026, 10, 17))

    assert len(report["rows"]) == 11
    assert report["rows"][0] == {
        "flat": "101",
        "owner": OWNER_MASTER["101"],
        "pending_balance": 12345,
        "paid_till": date(2026, 9, 1),
        "pending_months": 1,
    }
    assert report["rows"][1]["pending_balance"] == 19500
    assert report["rows"][1]["pending_months"] == 9
    assert report["rows"][3]["pending_balance"] == 9000
    assert report["rows"][3]["pending_months"] == 6

    rendered = render_maintenance_report(report)
    rendered_image = Image.open(rendered)
    assert rendered_image.format == "PNG"
    assert rendered.name == "maintenance-report-2026-10-17.png"
    header_y = 28 * 2 + 112 * 2 + 5
    assert rendered_image.getpixel((60, header_y)) == (214, 234, 248)


def test_report_credit_covers_elapsed_months_before_reporting_overdue(tmp_path: Path):
    file_path = tmp_path / "ledger.xlsx"
    make_workbook(file_path)
    workbook = load_workbook(file_path)
    ws = workbook["Sheet1"]
    ws.cell(7, 6).value = -4500
    ws.cell(7, 7).value = datetime(2026, 1, 1)
    workbook.save(file_path)

    fully_covered = get_maintenance_report(file_path, date(2026, 4, 17))["rows"][3]
    first_overdue = get_maintenance_report(file_path, date(2026, 5, 17))["rows"][3]

    assert (fully_covered["pending_balance"], fully_covered["pending_months"]) == (0, 0)
    assert (first_overdue["pending_balance"], first_overdue["pending_months"]) == (1500, 1)
    assert load_workbook(file_path, data_only=True)["Sheet1"].cell(7, 6).value == -4500
