#!/usr/bin/env python3
"""
Matches Google Form responses against a bank CSV statement export.
No AI calls — completely free to run.

Usage:
    uv run tools/match_bank_csv.py --csv path/to/bank_statement.csv
    uv run tools/match_bank_csv.py --csv statement.csv --event "Standard"
"""

import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

try:
    import gspread
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError as e:
    print(f"ERROR: Missing dependency — {e}")
    print("Run: pip install -r requirements.txt")
    sys.exit(1)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

CONFIG_PATH = Path("events_config.json")
TOKEN_PATH = Path("token.json")
CREDENTIALS_PATH = Path("credentials.json")
TMP_DIR = Path(".tmp")
USED_TXNS_PATH = TMP_DIR / "used_transactions.json"

# Keywords to detect bank CSV columns
CSV_COLUMN_KEYWORDS = {
    "txn_id": ["เลขที่รายการ", "transaction", "ref no", "refno"],
    "date":   ["วันทำรายการ", "date", "วันที่"],
    "sender": ["รับเงินจาก", "sender", "จากบัญชี", "from"],
    "amount": ["จำนวนเงิน", "amount", "ยอด"],
    "bank":   ["source of fund", "source", "ธนาคารต้นทาง"],
}

# Keywords to detect Google Form response columns
FORM_COLUMN_KEYWORDS = {
    "game_name":     ["in game name", "ingame", "แข่งในวงการ", "ชื่อแข่ง"],
    "openchat_name": ["openchat", "open chat"],
    "facebook":      ["facebook", "เฟสบุค", "fb"],
    "transfer_name": ["ชื่อบัญชี", "ชื่อที่โอน", "ชื่อโอน", "ใช้โอน", "transfer name", "ชื่อเจ้าของบัญชี"],
}

OUTPUT_HEADER = [
    "#", "ชื่อที่ใช้แข่ง", "ชื่อใน OpenChat", "ชื่อเฟสบุค", "ชื่อบัญชีที่โอน",
    "สถานะ", "รายละเอียด", "ยอดที่พบ", "ชื่อในธนาคาร", "เลขที่รายการ", "วันที่โอน",
]


# ── Google Auth ──────────────────────────────────────────────────────────────

def get_google_credentials() -> Credentials:
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                print("ERROR: credentials.json not found.")
                print("See workflows/setup_google_auth.md for setup instructions.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return creds


# ── Used Transaction Tracking ─────────────────────────────────────────────────

def load_used_txns() -> set:
    if USED_TXNS_PATH.exists():
        with open(USED_TXNS_PATH, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_used_txns(used: set) -> None:
    with open(USED_TXNS_PATH, "w", encoding="utf-8") as f:
        json.dump(sorted(used), f, ensure_ascii=False, indent=2)


# ── Bank CSV Loading ──────────────────────────────────────────────────────────

def load_bank_csv(path: str) -> list[dict]:
    """Load bank statement CSV, trying common Thai encodings."""
    for encoding in ("utf-8-sig", "utf-8", "cp874", "tis-620"):
        try:
            with open(path, encoding=encoding, newline="") as f:
                reader = csv.reader(f)
                rows = list(reader)
            if rows:
                return _parse_csv_rows(rows)
        except (UnicodeDecodeError, Exception):
            continue
    print(f"ERROR: Cannot read CSV file: {path}")
    sys.exit(1)


def _parse_csv_rows(rows: list[list[str]]) -> list[dict]:
    """Find the header row and parse transactions."""
    header_idx = None
    cols: dict[str, int] = {}

    for i, row in enumerate(rows):
        row_lower = [c.lower().strip() for c in row]
        # detect header row by looking for key columns
        matches = sum(
            1 for keywords in CSV_COLUMN_KEYWORDS.values()
            if any(kw in cell for kw in keywords for cell in row_lower)
        )
        if matches >= 2:
            header_idx = i
            for col_key, keywords in CSV_COLUMN_KEYWORDS.items():
                for j, cell in enumerate(row_lower):
                    if any(kw in cell for kw in keywords):
                        cols[col_key] = j
                        break
            break

    if header_idx is None:
        print("ERROR: Cannot find header row in bank CSV.")
        print("Expected columns like: เลขที่รายการ, รับเงินจาก, จำนวนเงิน")
        sys.exit(1)

    transactions = []
    for row in rows[header_idx + 1:]:
        if not row or not any(row):
            continue
        def cell(key: str) -> str:
            i = cols.get(key)
            return row[i].strip() if i is not None and i < len(row) else ""

        amount_str = cell("amount").replace(",", "").strip()
        try:
            amount = float(amount_str)
        except ValueError:
            continue  # skip total/summary rows

        txn = {
            "txn_id": cell("txn_id"),
            "date":   cell("date"),
            "sender": cell("sender"),
            "amount": amount,
            "bank":   cell("bank"),
        }
        if txn["sender"] or txn["txn_id"]:
            transactions.append(txn)

    print(f"  Loaded {len(transactions)} transactions from bank CSV")
    return transactions


# ── Name Normalization & Matching ─────────────────────────────────────────────

def normalize(name: str) -> str:
    return re.sub(r"[\s.\-_]", "", name).lower()


def find_matching_txn(
    transfer_name: str,
    facebook_name: str,
    expected_fee: float,
    bank_rows: list[dict],
    used_txn_ids: set,
) -> tuple[dict | None, str, str]:
    """
    Search bank_rows for a transaction matching name + amount.
    Uses transfer_name (from form) as primary key; Facebook name as fallback.
    Returns (txn_row | None, status_emoji, detail_text).
    """
    # Prefer transfer_name if provided, fall back to facebook_name
    match_name = transfer_name if transfer_name else facebook_name
    n_match = normalize(match_name)
    best_amount_match = None

    for txn in bank_rows:
        if txn["txn_id"] and txn["txn_id"] in used_txn_ids:
            continue

        amount_ok = abs(txn["amount"] - expected_fee) < 1.0
        n_sender = normalize(txn["sender"])
        name_ok = bool(n_match and n_sender and (n_match in n_sender or n_sender in n_match))

        if name_ok and amount_ok:
            return txn, "✅", f"ยืนยันแล้ว | โอน {txn['amount']:.0f}฿ | {txn['sender']}"

        if amount_ok and best_amount_match is None:
            best_amount_match = txn

    # Amount matches but name doesn't
    if best_amount_match:
        t = best_amount_match
        label = f"บัญชี: {match_name}" if transfer_name else f"FB: {facebook_name}"
        return t, "⚠️", (
            f"ตรวจสอบชื่อ | ยอดตรง {t['amount']:.0f}฿ | "
            f"ธนาคาร: {t['sender']} | {label}"
        )

    label = f"บัญชี: {match_name}" if transfer_name else f"FB: {facebook_name}"
    return None, "❌", f"ไม่พบในรายงานธนาคาร | {label} | คาดยอด {expected_fee:.0f}฿"


# ── Column Detection (Form) ───────────────────────────────────────────────────

def detect_form_columns(header: list[str]) -> dict[str, int]:
    cols: dict[str, int] = {}
    for col_key, keywords in FORM_COLUMN_KEYWORDS.items():
        for i, h in enumerate(header):
            if any(kw.lower() in h.lower() for kw in keywords):
                cols[col_key] = i
                break
    return cols


# ── Per-Event Processing ──────────────────────────────────────────────────────

def process_event(
    event: dict,
    gc: gspread.Client,
    output_sheet: gspread.Spreadsheet,
    bank_rows: list[dict],
    used_txn_ids: set,
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
    cols = detect_form_columns(header)

    missing = [k for k in FORM_COLUMN_KEYWORDS if k not in cols]
    if missing:
        print(f"  WARNING: Could not detect form columns: {missing}")
        print(f"  Header: {header}")

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
        transfer_name = cell("transfer_name")

        # If form has no In Game Name column, use OpenChat name as game identifier
        if not game_name:
            game_name = openchat_name
        if not game_name:
            continue

        name_display = transfer_name or facebook
        print(f"  [{seq:02d}] {game_name:<25} | ชื่อบัญชี: {name_display}")

        txn, status, detail = find_matching_txn(
            transfer_name, facebook, expected_fee, bank_rows, used_txn_ids
        )

        if status == "✅" and txn and txn["txn_id"]:
            used_txn_ids.add(txn["txn_id"])

        txn_id   = txn["txn_id"] if txn else ""
        txn_date = txn["date"]   if txn else ""
        txn_amt  = f"{txn['amount']:.0f}" if txn else ""
        txn_name = txn["sender"] if txn else ""

        print(f"       {status} {detail}")
        results.append([seq, game_name, openchat_name, facebook, transfer_name,
                        status, detail, txn_amt, txn_name, txn_id, txn_date])

    if results:
        ws.append_rows(results)

    confirmed = sum(1 for r in results if r[4] == "✅")
    print(f"\n  Result: {confirmed}/{len(results)} ยืนยันแล้ว")
    return len(results)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Match form responses against bank CSV")
    parser.add_argument("--csv", required=True, help="Path to bank statement CSV file")
    parser.add_argument("--event", help="Process only this event name (default: all)")
    args = parser.parse_args()

    TMP_DIR.mkdir(exist_ok=True)

    if not CONFIG_PATH.exists():
        print("ERROR: events_config.json not found.")
        print("Copy events_config.example.json → events_config.json and fill in your Sheet IDs.")
        sys.exit(1)

    csv_path = Path(args.csv)
    if not csv_path.exists():
        print(f"ERROR: CSV file not found: {args.csv}")
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

    print(f"Loading bank CSV: {args.csv}")
    bank_rows = load_bank_csv(str(csv_path))
    used_txn_ids = load_used_txns()

    print("Authenticating with Google...")
    creds = get_google_credentials()
    gc = gspread.authorize(creds)
    output_sheet = gc.open_by_key(output_sheet_id)

    total = 0
    for event in events:
        total += process_event(event, gc, output_sheet, bank_rows, used_txn_ids)
        save_used_txns(used_txn_ids)

    print(f"\n{'='*55}")
    print(f"เสร็จสิ้น! ประมวลผลทั้งหมด {total} รายการ")
    print(f"ดูผล: https://docs.google.com/spreadsheets/d/{output_sheet_id}")


if __name__ == "__main__":
    main()
