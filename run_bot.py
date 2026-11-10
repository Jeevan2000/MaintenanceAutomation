"""Telegram poller for the Ferro Towers append-only maintenance ledger."""
from __future__ import annotations

import asyncio
from io import BytesIO
import json
import os
import re
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zoneinfo import ZoneInfo

from dotenv import load_dotenv

from maintenance_utils import OWNER_MASTER, append_payment, get_maintenance_report, parse_date

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
STATE_FILE = os.getenv("STATE_FILE", "read_state.json")
LOG_FILE = os.getenv("LOG_FILE", "messages.log")
COMMAND_RE = re.compile(
    r"^\s*/mnt(?:@[A-Za-z0-9_]+)?\s+(\d{2,})\s+((?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,2})?)"
    r"(?:\s+(\d{4}-\d{2}-\d{2}|\d{2}-\d{2}-\d{4}|\d{2}/\d{2}/\d{4}))?\s*$",
    re.IGNORECASE,
)


def parse_mnt_command(text: str, sent_at: datetime | None = None) -> dict | None:
    """Parse /mnt <flat> <amount> [date]; return None for unrelated messages."""
    stripped = (text or "").strip()
    if not stripped.lower().startswith("/mnt"):
        return None
    match = COMMAND_RE.fullmatch(stripped)
    if not match:
        raise ValueError("Use: /mnt <flat no> <amount> [date, e.g. 16-03-2026]")

    flat, amount_text, date_text = match.groups()
    if flat not in OWNER_MASTER:
        raise ValueError(f"Unknown flat {flat}. Please use one of: {', '.join(OWNER_MASTER)}")
    try:
        amount = Decimal(amount_text.replace(",", ""))
    except InvalidOperation as exc:
        raise ValueError("Amount must be a valid number") from exc
    if not amount.is_finite() or amount <= 0:
        raise ValueError("Amount must be greater than zero")
    if date_text:
        paid_date = parse_date(date_text)
        if paid_date is None:
            raise ValueError("Invalid date. Use YYYY-MM-DD, DD-MM-YYYY, or DD/MM/YYYY")
    else:
        sent_at = sent_at or datetime.now(timezone.utc)
        if sent_at.tzinfo is None:
            sent_at = sent_at.replace(tzinfo=timezone.utc)
        paid_date = sent_at.astimezone(ZoneInfo("Asia/Kolkata")).date()
    return {"flat": flat, "amount": amount, "date": paid_date}


def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as state_file:
            return json.load(state_file)
    except (OSError, json.JSONDecodeError):
        return {}


def save_state(state: dict):
    updated = dict(state)
    updated["last_checked"] = datetime.now(timezone.utc).isoformat()
    state_path = Path(STATE_FILE)
    temp_path = state_path.with_name(state_path.name + ".tmp")
    with temp_path.open("w", encoding="utf-8") as state_file:
        json.dump(updated, state_file)
        state_file.flush()
        os.fsync(state_file.fileno())
    os.replace(temp_path, state_path)


def append_log(entry: str):
    timestamp = datetime.now(timezone.utc).isoformat()
    line = f"[{timestamp}] {entry}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as log_file:
            log_file.write(line + "\n")
    except OSError as exc:
        print(f"[{timestamp}] Could not write log file {LOG_FILE}: {exc}", flush=True)


def _message_sent_at(message) -> datetime:
    value = getattr(message, "date", None)
    return value if isinstance(value, datetime) else datetime.now(timezone.utc)


def render_maintenance_report(report: dict) -> BytesIO:
    """Render the maintenance report as a styled PNG image for Telegram."""
    from PIL import Image, ImageDraw, ImageFont

    report_date = report["as_of"]
    headers = (
        "Flat No.",
        "Owner Name",
        f"Balance till {report_date:%d-%b-%Y}",
        "Maintenance Pending Months",
    )
    rows = []
    for entry in report["rows"]:
        try:
            balance_amount = float(entry["pending_balance"])
            sign = "-" if balance_amount < 0 else ""
            balance = f"{sign}₹{abs(balance_amount):,.2f}"
        except (TypeError, ValueError):
            balance = "₹0.00"
        rows.append((
            str(entry["flat"]),
            str(entry["owner"]),
            balance,
            str(entry["pending_months"]),
        ))

    def load_font(bold: bool, size: int):
        candidates = (
            (r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\arial.ttf")
            if bold else (r"C:\Windows\Fonts\arial.ttf",)
        )
        candidates += (
            ("DejaVuSans-Bold.ttf",) if bold else ("DejaVuSans.ttf",)
        )
        for candidate in candidates:
            try:
                return ImageFont.truetype(candidate, size)
            except OSError:
                continue
        return ImageFont.load_default()

    scale = 2
    body_font = load_font(False, 22 * scale)
    header_font = load_font(True, 21 * scale)
    title_font = load_font(True, 30 * scale)
    subtitle_font = load_font(True, 24 * scale)
    padding_x = 18 * scale
    row_height = 46 * scale
    header_height = 60 * scale
    title_height = 112 * scale
    border = (191, 205, 218)
    header_fill = (214, 234, 248)
    body_fill = (255, 255, 255)
    alternate_fill = (245, 249, 252)
    text_color = (32, 45, 58)

    measure_image = Image.new("RGB", (1, 1))
    measure = ImageDraw.Draw(measure_image)
    column_widths = [
        max(
            140 * scale,
            int(max(measure.textlength(headers[index], font=header_font), *(measure.textlength(row[index], font=body_font) for row in rows))) + padding_x * 2,
        )
        for index in range(len(headers))
    ]
    table_width = sum(column_widths)
    margin = 28 * scale
    image_width = table_width + margin * 2
    image_height = title_height + header_height + row_height * len(rows) + margin * 2
    image = Image.new("RGB", (image_width, image_height), (240, 245, 249))
    draw = ImageDraw.Draw(image)

    x_start = margin
    y_start = margin + title_height
    title_lines = (
        ("Ferro Tower", title_font),
        ("Maintenance Report", subtitle_font),
    )
    title_metrics = [
        (text, font, draw.textbbox((0, 0), text, font=font))
        for text, font in title_lines
    ]
    title_gap = 8 * scale
    title_block_height = sum(bounds[3] - bounds[1] for _, _, bounds in title_metrics) + title_gap
    title_y = (y_start - title_block_height) // 2
    for text, font, bounds in title_metrics:
        text_height = bounds[3] - bounds[1]
        text_width = draw.textlength(text, font=font)
        draw.text(
            (image_width // 2 - text_width / 2, title_y - bounds[1]),
            text,
            font=font,
            fill=(27, 54, 78),
        )
        title_y += text_height + title_gap

    x_positions = [x_start]
    for width in column_widths:
        x_positions.append(x_positions[-1] + width)

    for row_index in range(len(rows) + 1):
        y_top = y_start + (header_height if row_index else 0) + row_height * max(row_index - 1, 0)
        height = header_height if row_index == 0 else row_height
        fill = header_fill if row_index == 0 else (body_fill if row_index % 2 else alternate_fill)
        for col_index in range(len(headers)):
            left, right = x_positions[col_index], x_positions[col_index + 1]
            draw.rectangle((left, y_top, right, y_top + height), fill=fill, outline=border, width=2)
            value = headers[col_index] if row_index == 0 else rows[row_index - 1][col_index]
            font = header_font if row_index == 0 else body_font
            bounds = draw.textbbox((0, 0), value, font=font)
            text_height = bounds[3] - bounds[1]
            draw.text(
                (left + padding_x, y_top + (height - text_height) // 2 - bounds[1]),
                value,
                font=font,
                fill=text_color,
            )

    output = BytesIO()
    image.save(output, format="PNG", optimize=True)
    output.seek(0)
    output.name = f"maintenance-report-{report_date:%Y-%m-%d}.png"
    return output


async def main_async(path: Path):
    if not BOT_TOKEN:
        raise SystemExit("BOT_TOKEN must be set in environment (.env)")
    try:
        from telegram import Bot
    except ImportError as exc:
        raise SystemExit("python-telegram-bot is required. Install with: pip install python-telegram-bot") from exc

    bot = Bot(BOT_TOKEN)
    state = load_state()
    try:
        offset = int(state["update_id"]) + 1 if "update_id" in state else None
    except (TypeError, ValueError):
        offset = None

    append_log(f"Bot started. Workbook: {path.resolve()}; log file: {Path(LOG_FILE).resolve()}")
    append_log(f"Starting Telegram long-poll loop at offset {offset}")
    try:
        while True:
            try:
                updates = await bot.get_updates(timeout=30, offset=offset)
            except Exception as exc:
                append_log(f"Error fetching updates: {exc}")
                await asyncio.sleep(5)
                continue

            if not updates:
                save_state(state)
                await asyncio.sleep(1)
                continue

            for update in updates:
                message = getattr(update, "message", None)
                append_log(f"Received Telegram update_id={update.update_id}")
                if message is not None:
                    text = getattr(message, "text", "") or ""
                    append_log(
                        f"Received message_id={getattr(message, 'message_id', 'unknown')} "
                        f"chat_id={message.chat.id}: {text!r}"
                    )
                    command_token = text.strip().split(maxsplit=1)[0] if text.strip() else ""
                    command_name = command_token.split("@", maxsplit=1)[0].lower()
                    if command_name == "/report":
                        append_log(f"Processing /report command in update_id={update.update_id}")
                        try:
                            if text.strip() != command_token:
                                raise ValueError("Use /report with no arguments")
                            report_date = datetime.now(ZoneInfo("Asia/Kolkata")).date()
                            report = get_maintenance_report(path, report_date)
                            report_image = render_maintenance_report(report)
                            append_log(
                                f"Generated /report for {len(report['rows'])} flats as of "
                                f"{report_date:%d-%b-%Y} from {path.resolve()}"
                            )
                            await bot.send_photo(
                                chat_id=message.chat.id,
                                photo=report_image,
                                caption=f"Ferro Towers maintenance report · {report_date:%d-%b-%Y}",
                            )
                            append_log(f"Sent /report to chat_id={message.chat.id} for update_id={update.update_id}")
                        except ValueError as exc:
                            await bot.send_message(chat_id=message.chat.id, text=f"Report not generated: {exc}")
                            append_log(f"Rejected /report command: {exc}")
                        except Exception as exc:
                            append_log(f"Failed to generate/send /report for update_id={update.update_id}: {exc}")
                            try:
                                await bot.send_message(chat_id=message.chat.id, text="Report could not be generated. Check the bot log.")
                            except Exception as send_exc:
                                append_log(f"Failed to send /report error acknowledgement: {send_exc}")
                    elif command_name == "/mnt":
                        try:
                            append_log(f"Processing /mnt command in update_id={update.update_id}")
                            record = parse_mnt_command(text, _message_sent_at(message))
                            result = append_payment(path, record["flat"], record["amount"], record["date"])
                            ack = (
                                f"Payment Recorded for Flat {record['flat']} ({OWNER_MASTER[record['flat']]}).\n"
                                f"Paid Date: {record['date'].strftime('%d-%b-%Y')}\n"
                                f"Pending Balance: {result['pending']}\n"
                                f"Amount: {result['amount']}\n"
                                f"Balance: {result['balance']}\n"
                                f"Paid Till: {result['paid_till'].strftime('%A, %B %d, %Y') if result['paid_till'] else 'unchanged (less than one full month paid)'}"
                            )
                            append_log(
                                f"Workbook updated successfully: {path.resolve()}, sheet row {result['row']}, "
                                f"flat={record['flat']}, paid_date={record['date']:%d-%b-%Y}, "
                                f"amount={result['amount']}, balance={result['balance']}"
                            )
                        except ValueError as exc:
                            ack = f"Payment not recorded: {exc}"
                            append_log(f"Rejected /mnt command: {exc}")
                        except Exception as exc:
                            # Do not commit the update id: Telegram will retry this message.
                            append_log(f"Workbook update failed for Telegram update {update.update_id}: {exc}")
                            offset = int(state["update_id"]) + 1 if "update_id" in state else None
                            break
                        try:
                            await bot.send_message(chat_id=message.chat.id, text=ack)
                            append_log(f"Acknowledgement sent for update_id={update.update_id}")
                        except Exception as exc:
                            append_log(f"Failed to send acknowledgement: {exc}")
                    else:
                        append_log(f"No maintenance command in message_id={getattr(message, 'message_id', 'unknown')}; ignored")
                else:
                    append_log(f"Update_id={update.update_id} has no message; skipped")

                state["update_id"] = update.update_id
                state["last_read_timestamp"] = datetime.now(timezone.utc).isoformat()
                save_state(state)
                offset = update.update_id + 1
                append_log(f"Update_id={update.update_id} processing complete; state saved")
            await asyncio.sleep(0.5)
    except (KeyboardInterrupt, asyncio.CancelledError):
        append_log("Stopping poll loop")


def main():
    workbook = Path(os.sys.argv[1]) if len(os.sys.argv) > 1 else Path("FerroTowerMainteance.xlsx")
    if not workbook.exists():
        raise SystemExit(f"File not found: {workbook}")
    asyncio.run(main_async(workbook))


if __name__ == "__main__":
    main()
