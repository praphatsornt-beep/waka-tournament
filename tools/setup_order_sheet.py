#!/usr/bin/env python3
"""ตั้ง structure Google Sheet สำหรับระบบออเดอร์การ์ด"""

import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

try:
    import gspread
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
except ImportError:
    print("ERROR: pip install gspread google-auth")
    sys.exit(1)

SHEET_URL = "https://docs.google.com/spreadsheets/d/1aUHbSt3qlQ4uMIzlCGbF-iFm0AqSeqx12nxk5ny1JoY/edit"
SHEET_ID  = "1aUHbSt3qlQ4uMIzlCGbF-iFm0AqSeqx12nxk5ny1JoY"

SCOPES     = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_PATH = Path("token.json")

BRANCHES = ["tonsak", "muangthong", "srinakarin"]

TABS = {
    "_catalog": [
        ["name", "category", "price_box", "price_pack", "active", "image_url"],
    ],
    "_config": [
        ["key", "value"],
        ["group_tonsak",      ""],   # Line Group ID สาขาต้นสัก
        ["group_muangthong",  ""],   # Line Group ID เมืองทอง
        ["group_srinakarin",  ""],   # Line Group ID ศรีนครินทร์
        ["delivery_fee",      "50"],
        ["bank_name",         "ธนาคาร..."],
        ["bank_account",      "xxx-x-xxxxx-x"],
        ["bank_account_name", "ชื่อบัญชี"],
    ],
    "orders": [
        ["order_id", "timestamp", "line_user_id", "display_name",
         "items_json", "total", "branch", "real_name", "phone",
         "slip_status", "slip_amount", "slip_txn_id", "notes"],
    ],
    "stock_tonsak": [
        ["name", "category", "qty_box", "qty_pack"],
    ],
    "stock_muangthong": [
        ["name", "category", "qty_box", "qty_pack"],
    ],
    "stock_srinakarin": [
        ["name", "category", "qty_box", "qty_pack"],
    ],
}


def get_gc():
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print("ERROR: ไม่พบ token.json")
            sys.exit(1)
    return gspread.authorize(creds)


def main():
    print(f"เชื่อมต่อ Sheet: {SHEET_ID}")
    gc  = get_gc()
    sht = gc.open_by_key(SHEET_ID)

    existing = {ws.title for ws in sht.worksheets()}
    print(f"Tab ที่มีอยู่: {existing}")

    for tab_name, rows in TABS.items():
        if tab_name in existing:
            print(f"  ข้าม '{tab_name}' — มีอยู่แล้ว")
            continue
        ws = sht.add_worksheet(title=tab_name, rows=500, cols=20)
        ws.append_rows(rows)
        print(f"  สร้าง '{tab_name}' ✅")

    print("\nเสร็จแล้ว! กรุณากรอก:")
    print("  _config → group_tonsak/muangthong/srinakarin = Line Group ID แต่ละสาขา")
    print("  _config → bank_name, bank_account, bank_account_name = เลขบัญชีธนาคาร")
    print("  _catalog → เพิ่มรายการสินค้า")


if __name__ == "__main__":
    main()
