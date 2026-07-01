#!/usr/bin/env python3
"""
Reads Google Form responses for each event, verifies payment slips via
Claude Vision, and writes confirmed participant lists to a Google Sheet.

Usage:
    uv run tools/process_registrations.py
    uv run tools/process_registrations.py --event "Standard"   # single event only
"""

import argparse
import base64
import hashlib
import io
import json
import os
import re
import sys
import time
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

try:
    import anthropic
    import gspread
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from PIL import Image
except ImportError as e:
    print(f"ERROR: Missing dependency — {e}")
    print("Run: pip install -r requirements.txt")
    sys.exit(1)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

CONFIG_PATH = Path("events_config.json")
SA_PATH     = Path("service_account.json")
TMP_DIR = Path(".tmp")
SEEN_SLIPS_PATH = TMP_DIR / "seen_slips.json"

# Keywords used to auto-detect column mapping from the form header row
COLUMN_KEYWORDS = {
    "game_name":      ["แข่งในวงการ", "ชื่อแข่ง", "game name", "ingame"],
    "openchat_name":  ["openchat", "open chat"],
    "facebook":       ["facebook", "เฟสบุค", "fb"],
    "slip_url":       ["สลิป", "slip", "หลักฐาน", "การชำระ"],
    "transfer_name":  ["ชื่อที่ใช้โอน", "ชื่อโอน", "transfer"],
}

OUTPUT_HEADER = [
    "#", "ชื่อที่ใช้แข่ง", "ชื่อใน OpenChat", "ชื่อเฟสบุค",
    "ชื่อที่โอนเงิน", "สถานะ", "รายละเอียด", "วันที่สลิป", "ธนาคาร",
]


# ── Duplicate Slip Tracking ──────────────────────────────────────────────────

def load_seen_slips() -> dict:
    if SEEN_SLIPS_PATH.exists():
        with open(SEEN_SLIPS_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_seen_slips(seen: dict) -> None:
    with open(SEEN_SLIPS_PATH, "w", encoding="utf-8") as f:
        json.dump(seen, f, ensure_ascii=False, indent=2)


# ── Google Auth ──────────────────────────────────────────────────────────────

def get_google_credentials() -> Credentials:
    return Credentials.from_service_account_file(str(SA_PATH), scopes=SCOPES)


# ── Column Detection ─────────────────────────────────────────────────────────

def detect_columns(header: list[str]) -> dict[str, int]:
    """Auto-detect column indices from the form header row."""
    cols: dict[str, int] = {}
    for col_key, keywords in COLUMN_KEYWORDS.items():
        for i, h in enumerate(header):
            if any(kw.lower() in h.lower() for kw in keywords):
                cols[col_key] = i
                break
    return cols


# ── Drive / Image Utilities ──────────────────────────────────────────────────

def extract_drive_file_id(url: str) -> str | None:
    for pattern in [r"/file/d/([a-zA-Z0-9_-]+)", r"[?&]id=([a-zA-Z0-9_-]+)"]:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return None


def download_drive_image(drive_service, file_id: str) -> bytes | None:
    try:
        request = drive_service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = dl.next_chunk()
        return buf.getvalue()
    except Exception as e:
        print(f"    WARNING: Cannot download file {file_id}: {e}")
        return None


def to_jpeg_bytes(raw: bytes) -> bytes:
    """Convert any image format to JPEG bytes for the API."""
    try:
        img = Image.open(io.BytesIO(raw))
        out = io.BytesIO()
        img.convert("RGB").save(out, format="JPEG", quality=90)
        return out.getvalue()
    except Exception:
        return raw


# ── AI Slip Reading ──────────────────────────────────────────────────────────

def read_slip_with_ai(client: anthropic.Anthropic, image_bytes: bytes) -> dict:
    jpeg = to_jpeg_bytes(image_bytes)
    b64 = base64.standard_b64encode(jpeg).decode()

    prompt = (
        "นี่คือสลิปโอนเงิน กรุณาดึงข้อมูลต่อไปนี้ออกมาในรูปแบบ JSON เท่านั้น:\n"
        '{"sender_name": "ชื่อผู้โอน", "amount": 0.00, "date": "DD/MM/YYYY", '
        '"bank": "ชื่อธนาคาร", "readable": true}\n'
        'ถ้าอ่านไม่ได้หรือไม่ใช่สลิปโอนเงิน ให้ตอบ {"readable": false}\n'
        "ตอบด้วย JSON เท่านั้น ห้ามมีข้อความอื่น"
    )

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=300,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                    {"type": "text", "text": prompt},
                ],
            }],
        )
        text = re.sub(r"```(?:json)?\s*|\s*```", "", response.content[0].text).strip()
        return json.loads(text)
    except Exception as e:
        print(f"    WARNING: AI slip read failed: {e}")
        return {"readable": False}


# ── Matching Logic ───────────────────────────────────────────────────────────

def normalize(name: str) -> str:
    return re.sub(r"[\s.\-_]", "", name).lower()


def match_status(slip: dict, form_transfer_name: str, expected_fee: float) -> tuple[str, str]:
    if not slip.get("readable", True):
        return "🔍", "อ่านสลิปไม่ได้"

    slip_name = slip.get("sender_name", "")
    slip_amount = float(slip.get("amount") or 0)

    n_form = normalize(form_transfer_name)
    n_slip = normalize(slip_name)
    name_ok = n_form in n_slip or n_slip in n_form
    amount_ok = abs(slip_amount - expected_fee) < 1.0

    if name_ok and amount_ok:
        return "✅", f"ยืนยันแล้ว | โอน {slip_amount:.0f}฿ | {slip_name}"
    if name_ok:
        return "⚠️", f"ยอดไม่ตรง | โอน {slip_amount:.0f}฿ (คาด {expected_fee:.0f}฿) | {slip_name}"
    if amount_ok:
        return "⚠️", f"ชื่อไม่ตรง | สลิป: {slip_name} | ฟอร์ม: {form_transfer_name}"
    return "❌", f"ไม่ตรงทั้งยอดและชื่อ | โอน {slip_amount:.0f}฿ | สลิป: {slip_name} | ฟอร์ม: {form_transfer_name}"


# ── Per-Event Processing ─────────────────────────────────────────────────────

def process_event(
    event: dict,
    gc: gspread.Client,
    drive_service,
    ai_client: anthropic.Anthropic,
    output_sheet: gspread.Spreadsheet,
    seen_slips: dict,
) -> int:
    name = event["name"]
    source_id = event["source_sheet_id"]
    expected_fee = float(event["fee"])

    print(f"\n{'='*55}")
    print(f"Event: {name}  |  Fee: {expected_fee:.0f}฿")

    rows = gc.open_by_key(source_id).sheet1.get_all_values()
    if len(rows) < 2:
        print("  No responses found — skipping.")
        return 0

    header, data_rows = rows[0], rows[1:]
    cols = detect_columns(header)

    missing = [k for k in COLUMN_KEYWORDS if k not in cols]
    if missing:
        print(f"  WARNING: Could not detect columns: {missing}")
        print(f"  Header row: {header}")

    # Prepare output worksheet
    try:
        ws = output_sheet.worksheet(name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = output_sheet.add_worksheet(title=name, rows=500, cols=len(OUTPUT_HEADER))

    ws.append_row(OUTPUT_HEADER)

    results = []
    for seq, row in enumerate(data_rows, start=1):
        def cell(key: str) -> str:
            i = cols.get(key)
            return row[i].strip() if i is not None and i < len(row) else ""

        game_name     = cell("game_name")
        openchat_name = cell("openchat_name")
        facebook      = cell("facebook")
        slip_url      = cell("slip_url")
        transfer_name = cell("transfer_name")

        if not game_name:
            continue

        print(f"  [{seq:02d}] {game_name:<25} | โอนชื่อ: {transfer_name}")

        slip_data: dict = {"readable": False}
        status = detail = slip_date = slip_bank = ""
        duplicate = False
        file_id = img_bytes = None

        if slip_url:
            file_id = extract_drive_file_id(slip_url)
            if file_id:
                img_bytes = download_drive_image(drive_service, file_id)
                if img_bytes:
                    slip_hash = hashlib.md5(img_bytes).hexdigest()
                    if slip_hash in seen_slips:
                        prev = seen_slips[slip_hash]
                        status = "🚫"
                        detail = f"สลิปซ้ำ | ใช้ไปแล้วโดย: {prev['game_name']} ({prev['event']})"
                        duplicate = True
                    else:
                        slip_data = read_slip_with_ai(ai_client, img_bytes)
                        time.sleep(0.3)
            else:
                print(f"       WARNING: Cannot parse Drive URL: {slip_url[:80]}")

        if not duplicate:
            status, detail = match_status(slip_data, transfer_name, expected_fee)
            slip_date = slip_data.get("date", "") if slip_data.get("readable") else ""
            slip_bank = slip_data.get("bank", "") if slip_data.get("readable") else ""
            if status == "✅" and img_bytes:
                seen_slips[hashlib.md5(img_bytes).hexdigest()] = {"game_name": game_name, "event": name}

        print(f"       {status} {detail}")

        results.append([seq, game_name, openchat_name, facebook, transfer_name,
                        status, detail, slip_date, slip_bank])

    if results:
        ws.append_rows(results)

    confirmed = sum(1 for r in results if r[5] == "✅")
    print(f"\n  Result: {confirmed}/{len(results)} ยืนยันแล้ว")
    return len(results)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Verify tournament registrations")
    parser.add_argument("--event", help="Process only this event name (default: all)")
    args = parser.parse_args()

    TMP_DIR.mkdir(exist_ok=True)

    if not CONFIG_PATH.exists():
        print("ERROR: events_config.json not found.")
        print("Copy events_config.example.json → events_config.json and fill in your Sheet IDs.")
        sys.exit(1)

    with open(CONFIG_PATH, encoding="utf-8") as f:
        config = json.load(f)

    events = config.get("events", [])
    output_sheet_id = config.get("output_sheet_id", "")

    if not events:
        print("ERROR: No events in events_config.json")
        sys.exit(1)
    if not output_sheet_id:
        print("ERROR: output_sheet_id not set in events_config.json")
        sys.exit(1)

    if args.event:
        events = [e for e in events if e["name"] == args.event]
        if not events:
            print(f"ERROR: Event '{args.event}' not found in config.")
            sys.exit(1)

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY not set in .env")
        sys.exit(1)

    print("Authenticating with Google...")
    creds = get_google_credentials()
    gc = gspread.authorize(creds)
    drive_service = build("drive", "v3", credentials=creds)
    ai_client = anthropic.Anthropic(api_key=api_key)
    output_sheet = gc.open_by_key(output_sheet_id)

    seen_slips = load_seen_slips()

    total = 0
    for event in events:
        total += process_event(event, gc, drive_service, ai_client, output_sheet, seen_slips)
        save_seen_slips(seen_slips)

    print(f"\n{'='*55}")
    print(f"เสร็จสิ้น! ประมวลผลทั้งหมด {total} รายการ")
    print(f"ดูผล: https://docs.google.com/spreadsheets/d/{output_sheet_id}")


if __name__ == "__main__":
    main()
