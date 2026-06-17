#!/usr/bin/env python3
"""
WAKA Tournament — Payment Verification Web App
วิธีรัน: python -m streamlit run tools/verify_app.py
"""

import csv
import io
import json
import os
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

try:
    GMAIL_ADDRESS = st.secrets.get("GMAIL_ADDRESS", "") or os.getenv("GMAIL_ADDRESS", "")
    GMAIL_APP_PWD = st.secrets.get("GMAIL_APP_PASSWORD", "") or os.getenv("GMAIL_APP_PASSWORD", "")
except Exception:
    GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS", "")
    GMAIL_APP_PWD = os.getenv("GMAIL_APP_PASSWORD", "")

try:
    import gspread
    import pandas as pd
    import pdfplumber
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError as e:
    st.error(f"ติดตั้ง packages ก่อน: `pip install -r requirements.txt`\n\nรายละเอียด: {e}")
    st.stop()

# ── Paths & constants ─────────────────────────────────────────────────────────

SCOPES      = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_PATH  = Path("token.json")
CREDS_PATH  = Path("credentials.json")
CONFIG_PATH = Path("events_config.json")

COL_NAME = "ชื่อการแข่งขัน"
COL_URL  = "ลิงค์ Form Responses"
COL_FEE  = "ค่าสมัคร (฿)"
COL_DATE = "วันแข่งขัน"
COL_TIME = "เวลานัด"

FORM_COLUMN_KEYWORDS = {
    "game_name":     ["in game name", "ingame", "แข่งในวงการ", "ชื่อแข่ง", "trainer id", "openchat"],
    "openchat_name": ["openchat", "open chat"],
    "facebook":      ["facebook", "เฟสบุค", "fb"],
    "transfer_name": ["ชื่อบัญชี", "ชื่อที่โอน", "ชื่อโอน", "ใช้โอน", "ใช้ในการโอน",
                      "transfer name", "ชื่อเจ้าของบัญชี", "ชื่อที่ใช้"],
    "slip_url":      ["สลิป", "slip", "หลักฐาน", "การชำระ", "payment", "อัพโหลด"],
    "email":         ["อีเมล", "email", "e-mail", "gmail", "mail"],
}

OUTPUT_HEADER = [
    "#", "ชื่อที่ใช้แข่ง", "ชื่อใน OpenChat", "ชื่อเฟสบุค", "ชื่อบัญชีที่โอน",
    "สถานะ", "รายละเอียด", "ยอดที่พบ", "ชื่อในธนาคาร", "เลขที่รายการ", "วันที่โอน", "ลิงค์สลิป",
    "ตรวจสลิปแล้ว",  # admin กรอกเอง — ไม่โดนลบเมื่อรันใหม่
]

# Pre-filled defaults (ใช้ครั้งแรก ถ้ายังไม่มี events_config.json)
DEFAULT_EVENTS = [
    {COL_NAME: "Lorcana",           COL_URL: "https://docs.google.com/spreadsheets/d/1vEzBdnQ1doPH3KLADSN9XbrC5tViZiGq1kLWET-6LeQ/edit#gid=118500649",  COL_FEE: "350",     COL_DATE: "", COL_TIME: ""},
    {COL_NAME: "Riftbound",         COL_URL: "https://docs.google.com/spreadsheets/d/1EIqD91sJGBpwN0AeU9w8ior67elMIQS1vyY2mwShLZU/edit#gid=935532850",  COL_FEE: "500,900", COL_DATE: "", COL_TIME: ""},
    {COL_NAME: "Pokemon Champions", COL_URL: "https://docs.google.com/spreadsheets/d/1HLZ_tM3WbCiPP6BtuHSKmBo2H5KkYizsh71lJ4_moJA/edit#gid=2026372639", COL_FEE: "200",     COL_DATE: "", COL_TIME: ""},
    {COL_NAME: "Pokemon TCG",       COL_URL: "https://docs.google.com/spreadsheets/d/1eHQQ-8ANDJfMcUyVwtrtumfgrtnTaiZr--wzUFExMpE/edit#gid=759768958",  COL_FEE: "250",     COL_DATE: "", COL_TIME: ""},
    {COL_NAME: "BOT",               COL_URL: "https://docs.google.com/spreadsheets/d/1K7AFCxOOPzw-kp0kzUf07d-96dCsKnHjkYbIemzxeDM/edit#gid=2054056416", COL_FEE: "250",     COL_DATE: "", COL_TIME: ""},
]
DEFAULT_OUTPUT_URL = "https://docs.google.com/spreadsheets/d/1WSfd9sKHl2H5O7Ai1DvqVaL7Tuwni-nfBc-hTX8HilA"

# ── Config helpers ─────────────────────────────────────────────────────────────

def load_config() -> dict:
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"events": DEFAULT_EVENTS, "output_sheet_url": DEFAULT_OUTPUT_URL}

def save_config(events_records: list, output_url: str):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump({"events": events_records, "output_sheet_url": output_url},
                  f, ensure_ascii=False, indent=2)

def parse_sheet_url(url: str) -> tuple[str, int | None]:
    """แปลง Google Sheets URL → (sheet_id, gid หรือ None)"""
    url = url.strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if m:
        sheet_id = m.group(1)
        gid_m    = re.search(r"[#?&]gid=(\d+)", url)
        return sheet_id, (int(gid_m.group(1)) if gid_m else None)
    if re.match(r"^[a-zA-Z0-9_-]{20,}$", url):
        return url, None
    raise ValueError("ไม่ใช่ Google Sheets URL ที่ถูกต้อง")

def parse_fees(fee_str: str) -> list[float]:
    """'350' → [350.0]  |  '500,900' → [500.0, 900.0]"""
    parts = [p.strip() for p in str(fee_str).split(",") if p.strip()]
    fees  = []
    for p in parts:
        try:
            fees.append(float(p))
        except ValueError:
            raise ValueError(f"ค่าสมัครไม่ถูกต้อง: '{p}' (ใส่ตัวเลข เช่น 350 หรือ 500,900)")
    if not fees:
        raise ValueError("ไม่ได้ระบุค่าสมัคร")
    return fees

# ── Google Auth ────────────────────────────────────────────────────────────────

@st.cache_resource
def get_gc():
    # บน Streamlit Cloud — อ่านจาก Secrets
    try:
        has_token = "GOOGLE_TOKEN" in st.secrets
    except Exception:
        has_token = False
    if has_token:
        token_data = json.loads(st.secrets["GOOGLE_TOKEN"])
        creds = Credentials(
            token=token_data.get("token"),
            refresh_token=token_data.get("refresh_token"),
            token_uri=token_data.get("token_uri"),
            client_id=token_data.get("client_id"),
            client_secret=token_data.get("client_secret"),
            scopes=token_data.get("scopes"),
        )
        if not creds.valid:
            creds.refresh(Request())
        return gspread.authorize(creds)

    # Local — อ่านจากไฟล์
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDS_PATH.exists():
                raise FileNotFoundError(
                    "ไม่พบ credentials.json — วางไฟล์ไว้ใน root folder ของโปรเจค"
                )
            flow  = InstalledAppFlow.from_client_secrets_file(str(CREDS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return gspread.authorize(creds)

# ── Matching logic ─────────────────────────────────────────────────────────────

def normalize(name):
    return re.sub(r"[\s.\-_]", "", name).lower()

def name_similarity(a, b):
    n_a, n_b = normalize(a), normalize(b)
    if not n_a or not n_b:
        return 0.0
    if n_a in n_b or n_b in n_a:
        return 1.0
    common = sum(1 for c in n_a if c in n_b)
    return common / max(len(n_a), len(n_b))

def detect_columns(header, keywords):
    cols = {}
    for key, kws in keywords.items():
        for i, h in enumerate(header):
            if any(kw.lower() in h.lower() for kw in kws):
                cols[key] = i
                break
    return cols

def load_bank_pdf(content_bytes: bytes) -> list[dict]:
    txn_re  = re.compile(r"(T30-\d{4}-\d{4}-\d{4}-\d+)\s+(.+?)\s+([\d,]+\.\d{2})\s+([A-Z]{2,8})", re.UNICODE)
    date_re = re.compile(r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})")
    txns = []
    with pdfplumber.open(io.BytesIO(content_bytes)) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").split("\n"):
                m = txn_re.search(line)
                if not m:
                    continue
                dm = date_re.search(line[:m.start()])
                try:
                    amount = float(m.group(3).replace(",", ""))
                except ValueError:
                    continue
                txns.append({
                    "txn_id": m.group(1),
                    "date":   dm.group(1) if dm else "",
                    "sender": m.group(2).strip(),
                    "amount": amount,
                    "bank":   m.group(4),
                })
    if not txns:
        raise ValueError("ไม่พบ transactions ใน PDF — ตรวจสอบว่าเป็น PDF แบบ text ไม่ใช่รูปภาพ")
    return txns

def _detect_cols_by_pattern(rows: list[list[str]]) -> dict[str, int] | None:
    """ตรวจหา column โดยดูจากรูปแบบข้อมูล (T30-, วันที่, ตัวเลข, ชื่อธนาคาร)"""
    BANK_NAMES = {"scb", "ktb", "kbank", "bbl", "tmb", "ttb", "bay", "gsb", "uob", "lhbank"}
    txn_re  = re.compile(r"T30-\d{4}-\d{4}-\d{4}-\d+")
    date_re = re.compile(r"\d{2}/\d{2}/\d{4}")

    votes: dict[int, dict[str, int]] = {}
    for row in rows[:30]:
        for i, cell in enumerate(row):
            c = cell.strip()
            if txn_re.match(c):
                votes.setdefault(i, {})["txn_id"] = votes[i].get("txn_id", 0) + 1
            if date_re.match(c):
                votes.setdefault(i, {})["date"] = votes[i].get("date", 0) + 1
            if c.lower() in BANK_NAMES:
                votes.setdefault(i, {})["bank"] = votes[i].get("bank", 0) + 1
            try:
                v = float(c.replace(",", ""))
                if v > 0:
                    votes.setdefault(i, {})["amount"] = votes[i].get("amount", 0) + 1
            except ValueError:
                pass

    cols: dict[str, int] = {}
    for typ in ("txn_id", "date", "amount", "bank"):
        best = max((i for i in votes if typ in votes[i]),
                   key=lambda i: votes[i].get(typ, 0), default=None)
        if best is not None:
            cols[typ] = best

    # sender อยู่ถัดจาก txn_id
    if "txn_id" in cols:
        cols["sender"] = cols["txn_id"] + 1

    return cols if len(cols) >= 3 else None


def load_bank_csv(content_bytes: bytes) -> list[dict]:
    CSV_KWS = {
        "txn_id": ["เลขที่รายการ", "transaction", "ref no", "refno"],
        "date":   ["วันทำรายการ", "date", "วันที่"],
        "sender": ["รับเงินจาก", "sender", "จากบัญชี", "from"],
        "amount": ["จำนวนเงิน", "amount", "ยอด"],
        "bank":   ["source of fund", "source", "ธนาคารต้นทาง"],
    }
    for enc in ("utf-8-sig", "utf-8", "cp874", "tis-620"):
        try:
            text = content_bytes.decode(enc)
        except UnicodeDecodeError:
            continue

        for delimiter in ("|", ",", "\t"):
            try:
                rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
            except Exception:
                continue
            if not rows or max(len(r) for r in rows[:5]) < 3:
                continue

            # ลอง keyword detection ก่อน
            header_idx, cols = None, {}
            for i, row in enumerate(rows):
                rl = [c.lower().strip() for c in row]
                hits = sum(1 for kws in CSV_KWS.values()
                           if any(kw in cell for kw in kws for cell in rl))
                if hits >= 1:
                    header_idx = i
                    for key, kws in CSV_KWS.items():
                        for j, cell in enumerate(rl):
                            if any(kw in cell for kw in kws):
                                cols[key] = j
                                break
                    break

            # ถ้า keyword ไม่ได้ผล ใช้ pattern detection
            if header_idx is None or len(cols) < 3:
                cols = _detect_cols_by_pattern(rows) or {}
                header_idx = 0 if cols else None

            if header_idx is None:
                continue

            txns = []
            for row in rows[header_idx + 1:]:
                if not row or not any(row):
                    continue
                def cell(k, _r=row):
                    i = cols.get(k)
                    return _r[i].strip() if i is not None and i < len(_r) else ""
                try:
                    amt = float(cell("amount").replace(",", ""))
                except ValueError:
                    continue
                t = {"txn_id": cell("txn_id"), "date": cell("date"),
                     "sender": cell("sender"), "amount": amt, "bank": cell("bank")}
                if t["sender"] or t["txn_id"]:
                    txns.append(t)
            if txns:
                return txns

    raise ValueError("ไม่สามารถอ่าน CSV ได้ — ลองส่งออกใหม่จากแอป SCB แม่มณี")

def find_matching_txn(transfer_name, facebook_name, expected_fees, bank_rows, used_txn_ids):
    if not isinstance(expected_fees, list):
        expected_fees = [float(expected_fees)]
    match_name = transfer_name if transfer_name else facebook_name
    n_match = normalize(match_name)
    for txn in bank_rows:
        if txn["txn_id"] and txn["txn_id"] in used_txn_ids:
            continue
        amount_ok = any(abs(txn["amount"] - f) < 1.0 for f in expected_fees)
        n_sender  = normalize(txn["sender"])
        name_ok   = bool(n_match and n_sender and (n_match in n_sender or n_sender in n_match))
        if name_ok and amount_ok:
            return txn, "✅", f"ยืนยันแล้ว | โอน {txn['amount']:.0f}฿ | {txn['sender']}"
    return None, "❌", ""

def process_event(event, gc, bank_rows, used_txn_ids, output_sheet=None):
    name          = event["name"]
    expected_fees = event["fee"] if isinstance(event["fee"], list) else [float(event["fee"])]

    src_sh = gc.open_by_key(event["source_sheet_id"])
    src_ws = src_sh.get_worksheet_by_id(event["gid"]) if "gid" in event else src_sh.sheet1
    rows   = src_ws.get_all_values()

    if len(rows) < 2:
        return [], 0, 0, 0, 0

    header, data_rows = rows[0], rows[1:]
    cols = detect_columns(header, FORM_COLUMN_KEYWORDS)

    # อ่านค่า "ตรวจสลิปแล้ว" ที่ admin กรอกไว้ก่อน clear
    manual_ok = {}  # game_name → note
    if output_sheet:
        try:
            ws_prev = output_sheet.worksheet(name)
            prev    = ws_prev.get_all_values()
            if len(prev) > 1:
                hdr   = prev[0]
                oc_i  = next((i for i, h in enumerate(hdr) if h == "ตรวจสลิปแล้ว"), None)
                gn_i  = next((i for i, h in enumerate(hdr) if h == "ชื่อที่ใช้แข่ง"), 1)
                if oc_i is not None:
                    for r in prev[1:]:
                        val = r[oc_i].strip() if oc_i < len(r) else ""
                        gn  = r[gn_i]        if gn_i  < len(r) else ""
                        if val and gn:
                            manual_ok[gn] = val
        except Exception:
            pass

    parsed = []
    for seq, row in enumerate(data_rows, start=1):
        def cell(key, _row=row):
            i = cols.get(key)
            return _row[i].strip() if i is not None and i < len(_row) else ""
        game_name = cell("game_name") or cell("openchat_name")
        if game_name:
            parsed.append((seq, game_name, cell("openchat_name"), cell("facebook"),
                           cell("transfer_name"), cell("slip_url"), cell("email")))

    # Pass 1: lock ✅ (ชื่อ + ยอดตรง)
    exact_matches, txn_to_owner = {}, {}
    for seq, game_name, oc, fb, tr, slip, email_addr in parsed:
        txn, status, detail = find_matching_txn(tr, fb, expected_fees, bank_rows, used_txn_ids)
        if status == "✅" and txn:
            used_txn_ids.add(txn["txn_id"])
            txn_to_owner[txn["txn_id"]] = game_name
            exact_matches[seq] = (txn, status, detail)

    # Pass 2: ⚠️ โดยเรียงความคล้ายชื่อ
    candidates = []
    for seq, game_name, oc, fb, tr, slip, email_addr in parsed:
        if seq in exact_matches:
            continue
        match_name = tr if tr else fb
        for txn in bank_rows:
            if txn["txn_id"] and txn["txn_id"] in used_txn_ids:
                continue
            if any(abs(txn["amount"] - f) < 1.0 for f in expected_fees):
                sim = name_similarity(match_name, txn["sender"])
                candidates.append((sim, seq, game_name, match_name, txn))

    candidates.sort(key=lambda x: -x[0])
    warn_matches = {}
    for sim, seq, gname, match_name, txn in candidates:
        if seq in warn_matches or (txn["txn_id"] and txn["txn_id"] in used_txn_ids):
            continue
        detail = f"ตรวจสอบชื่อ | ยอดตรง {txn['amount']:.0f}฿"
        warn_matches[seq] = (txn, "⚠️", detail)
        used_txn_ids.add(txn["txn_id"])
        txn_to_owner[txn["txn_id"]] = gname

    # Compile
    results = []
    emails  = {}  # game_name → email address
    for seq, game_name, oc, fb, tr, slip, email_addr in parsed:
        if seq in exact_matches:
            txn, status, detail = exact_matches[seq]
        elif seq in warn_matches:
            txn, status, detail = warn_matches[seq]
        else:
            match_name = tr if tr else fb
            fees_str   = "/".join(f"{f:.0f}" for f in expected_fees)
            best_txn, best_sim = None, 0.0
            for t in bank_rows:
                if any(abs(t["amount"] - f) < 1.0 for f in expected_fees):
                    s = name_similarity(match_name, t["sender"])
                    if s > best_sim:
                        best_sim, best_txn = s, t
            if best_txn and best_sim == 1.0 and best_txn["txn_id"] in txn_to_owner:
                txn    = best_txn
                status = "🚫"
                detail = f"สลิปซ้ำ | ใช้ไปแล้วโดย: {txn_to_owner[best_txn['txn_id']]} | T30: {best_txn['txn_id']}"
            else:
                txn    = None
                status = "❌"
                detail = f"ไม่พบในรายงานธนาคาร | บัญชี: {match_name} | คาดยอด {fees_str}฿"

        # ถ้า admin ตรวจสลิปแล้ว → เปลี่ยนสถานะเป็น ✅
        override = manual_ok.get(game_name, "")
        if override and status == "⚠️":
            status = "✅"
            detail = f"ตรวจสลิปแล้ว | {detail}"

        display_name = (
            txn["sender"] if status in ("✅", "🚫") and txn
            else (tr or fb)
        )
        if email_addr:
            emails[game_name] = email_addr
        results.append([
            seq, game_name, oc, fb, tr, status, detail,
            f"{txn['amount']:.0f}" if txn else "",
            display_name,
            txn["txn_id"] if txn else "",
            txn["date"]   if txn else "",
            slip,
            override,
        ])

    # Write to Google Sheet (ถ้ามี)
    if output_sheet:
        try:
            ws = output_sheet.worksheet(name)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = output_sheet.add_worksheet(title=name, rows=500, cols=len(OUTPUT_HEADER))
        ws.append_row(OUTPUT_HEADER)
        if results:
            ws.append_rows(results)
            # ใส่สูตร dropdown ใน column ตรวจสลิปแล้ว (col M = 13)
            # ไม่ต้องทำ — admin กรอกเองได้เลย

    confirmed = sum(1 for r in results if r[5] == "✅")
    warned    = sum(1 for r in results if r[5] == "⚠️")
    duped     = sum(1 for r in results if r[5] == "🚫")
    failed    = len(results) - confirmed - warned - duped
    return results, confirmed, warned, duped, failed, emails

# ── Email sending ─────────────────────────────────────────────────────────────

def send_confirmation_email(
    to_email: str, ev_name: str, game_name: str, order_num: int,
    event_date: str = "", event_time: str = "",
) -> None:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[WAKA] ยืนยันการเข้าแข่งขัน {ev_name}"
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = to_email
    date_row = f'<tr><td style="padding:4px 20px 4px 0;color:#555">วันแข่งขัน</td><td><strong>{event_date}</strong></td></tr>' if event_date else ""
    time_row = f'<tr><td style="padding:4px 20px 4px 0;color:#555">เวลานัด</td><td><strong>{event_time}</strong></td></tr>' if event_time else ""
    body = f"""
<div style="font-family:sans-serif;max-width:520px;padding:16px;color:#222">
  <h2 style="color:#1a73e8">🎮 ยืนยันการเข้าแข่งขัน {ev_name}</h2>
  <p>สวัสดีคุณ <strong>{game_name}</strong>,</p>
  <p>การสมัครแข่งขันของคุณได้รับการยืนยันเรียบร้อยแล้ว ✅</p>
  <table style="border-collapse:collapse;margin:12px 0;font-size:15px">
    <tr><td style="padding:4px 20px 4px 0;color:#555">ลำดับที่</td><td><strong>#{order_num}</strong></td></tr>
    <tr><td style="padding:4px 20px 4px 0;color:#555">ชื่อที่ใช้แข่ง</td><td><strong>{game_name}</strong></td></tr>
    <tr><td style="padding:4px 20px 4px 0;color:#555">การแข่งขัน</td><td><strong>{ev_name}</strong></td></tr>
    {date_row}
    {time_row}
  </table>
  <p>โปรดมาถึงสถานที่ก่อนเวลาแข่งขันครับ</p>
  <p style="margin-top:24px;color:#888;font-size:13px">— ทีม WAKA Tournament</p>
</div>
"""
    msg.attach(MIMEText(body, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PWD)
        server.sendmail(GMAIL_ADDRESS, to_email, msg.as_string())

# ── Row color styling ─────────────────────────────────────────────────────────

STATUS_COLORS = {
    "✅": "background-color: #d4edda; color: #155724",
    "⚠️": "background-color: #fff3cd; color: #856404",
    "🚫": "background-color: #f8d7da; color: #721c24",
    "❌": "background-color: #f8d7da; color: #721c24",
}

def style_row(row):
    color = STATUS_COLORS.get(row["สถานะ"], "")
    return [color] * len(row)

# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="WAKA Tournament", page_icon="🎮", layout="wide")
st.title("🎮 WAKA Tournament — ตรวจสอบการชำระเงิน")

config = load_config()

# ─── ส่วน 1: ตารางการแข่งขัน ──────────────────────────────────────────────────

st.subheader("📋 การแข่งขัน")
st.caption("เพิ่ม/แก้ไขแถว ได้เลย | ค่าสมัครหลายค่าคั่นด้วยคอมมา เช่น `500,900`")

saved_events = config.get("events") or DEFAULT_EVENTS
events_df = pd.DataFrame(saved_events)
for col in [COL_NAME, COL_URL, COL_FEE, COL_DATE, COL_TIME]:
    if col not in events_df.columns:
        events_df[col] = ""

edited_df = st.data_editor(
    events_df[[COL_NAME, COL_URL, COL_FEE, COL_DATE, COL_TIME]],
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        COL_NAME: st.column_config.TextColumn(COL_NAME, width="medium"),
        COL_URL:  st.column_config.TextColumn(COL_URL,  width="large"),
        COL_FEE:  st.column_config.TextColumn(COL_FEE,  width="small"),
        COL_DATE: st.column_config.TextColumn(COL_DATE, width="small", help="เช่น 28 มิ.ย. 69"),
        COL_TIME: st.column_config.TextColumn(COL_TIME, width="small", help="เช่น 10:00 น."),
    },
)

# ─── ส่วน 2: Output ───────────────────────────────────────────────────────────

st.subheader("📤 Output")
output_url = st.text_input(
    "Google Sheet URL สำหรับบันทึกผล (ไม่บังคับ — ถ้าว่างจะแสดงเฉพาะในหน้านี้)",
    value=config.get("output_sheet_url", ""),
    placeholder="https://docs.google.com/spreadsheets/d/...",
)

if st.button("💾 บันทึกการตั้งค่า"):
    save_config(edited_df.to_dict("records"), output_url)
    st.success("บันทึกแล้ว — การตั้งค่าจะถูกโหลดอัตโนมัติครั้งถัดไป")

st.divider()

# ─── ส่วน 3: ไฟล์ธนาคาร + รัน ────────────────────────────────────────────────

st.subheader("📁 ไฟล์รายงานธนาคาร")
bank_file = st.file_uploader(
    "อัปโหลด PDF หรือ CSV จาก SCB แม่มณี",
    type=["pdf", "csv"],
)

# เก็บเนื้อหาไฟล์ใน session_state ป้องกัน reset หลังกดปุ่ม
if bank_file is not None:
    st.session_state["bank_content"] = bank_file.read()
    st.session_state["bank_name"]    = bank_file.name

has_file    = "bank_content" in st.session_state
run_clicked = st.button("🚀 เริ่มตรวจสอบ", type="primary", disabled=not has_file)

if run_clicked:
    # Parse events from editable table
    events = []
    for _, row in edited_df.iterrows():
        name    = str(row.get(COL_NAME, "")).strip()
        url     = str(row.get(COL_URL,  "")).strip()
        fee_raw = str(row.get(COL_FEE,  "")).strip()
        if not name or not url:
            continue
        try:
            sheet_id, gid = parse_sheet_url(url)
            fees          = parse_fees(fee_raw)
        except ValueError as e:
            st.warning(f"⚠️ ข้าม '{name}': {e}")
            continue
        ev = {"name": name, "source_sheet_id": sheet_id, "fee": fees}
        if gid is not None:
            ev["gid"] = gid
        events.append(ev)

    if not events:
        st.error("ไม่มีการแข่งขันที่ตั้งค่าถูกต้อง — ตรวจสอบตารางด้านบน")
        st.stop()

    # Load bank file
    with st.spinner("กำลังอ่านไฟล์ธนาคาร..."):
        content   = st.session_state["bank_content"]
        bank_name = st.session_state["bank_name"]
        try:
            if bank_name.lower().endswith(".pdf") or content[:4] == b"%PDF":
                bank_rows = load_bank_pdf(content)
            else:
                bank_rows = load_bank_csv(content)
            st.info(f"โหลด **{len(bank_rows)} transactions** จากธนาคาร")
        except Exception as e:
            st.error(f"อ่านไฟล์ไม่ได้: {e}")
            st.stop()

    # Connect Google Sheets
    with st.spinner("กำลังเชื่อมต่อ Google Sheets..."):
        try:
            gc = get_gc()
        except FileNotFoundError as e:
            st.error(str(e))
            st.stop()
        except Exception as e:
            st.error(f"เชื่อมต่อ Google ไม่ได้: {e}")
            st.stop()

    # Parse output sheet (optional)
    output_sheet = None
    out_sheet_id = None
    if output_url.strip():
        try:
            out_sheet_id, _ = parse_sheet_url(output_url)
            output_sheet     = gc.open_by_key(out_sheet_id)
        except Exception as e:
            st.warning(f"เปิด Output Sheet ไม่ได้: {e} — จะแสดงเฉพาะในหน้านี้")

    # Process all events
    used_txn_ids = set()
    summary      = []
    all_results  = {}
    all_emails   = {}
    progress_bar = st.progress(0, text="กำลังประมวลผล...")

    for idx, event in enumerate(events):
        progress_bar.progress(idx / len(events), text=f"กำลังตรวจ {event['name']}...")
        try:
            results, ok, warn, dup, fail, emails = process_event(
                event, gc, bank_rows, used_txn_ids, output_sheet
            )
            all_results[event["name"]] = results
            all_emails[event["name"]]  = emails
            summary.append({
                "การแข่งขัน":   event["name"],
                "✅ ยืนยัน":    ok,
                "⚠️ ตรวจสอบ":  warn,
                "🚫 ซ้ำ":       dup,
                "❌ ไม่พบ":     fail,
                "ทั้งหมด":      len(results),
            })
        except Exception as e:
            st.warning(f"**{event['name']}** เกิดข้อผิดพลาด: {e}")
            summary.append({
                "การแข่งขัน": event["name"],
                "✅ ยืนยัน": "-", "⚠️ ตรวจสอบ": "-",
                "🚫 ซ้ำ": "-", "❌ ไม่พบ": "-", "ทั้งหมด": "-",
            })

    progress_bar.progress(1.0, text="เสร็จสิ้น!")

    # Store in session state so results persist across reruns
    st.session_state["all_results"]      = all_results
    st.session_state["all_emails"]       = all_emails
    st.session_state["summary_data"]     = summary
    st.session_state["out_sheet_id_run"] = out_sheet_id
    st.session_state["events_meta"]      = {
        str(row.get(COL_NAME, "")).strip(): {
            "date": str(row.get(COL_DATE, "")).strip(),
            "time": str(row.get(COL_TIME, "")).strip(),
        }
        for _, row in edited_df.iterrows()
        if str(row.get(COL_NAME, "")).strip()
    }
    st.session_state["run_count"]        = st.session_state.get("run_count", 0) + 1
    st.session_state["save_count"]       = 0

# ── Display results (runs whenever session state has data) ────────────────────
if "all_results" in st.session_state:
    all_results  = st.session_state["all_results"]
    out_sheet_id = st.session_state.get("out_sheet_id_run")
    run_count    = st.session_state.get("run_count", 0)
    save_count   = st.session_state.get("save_count", 0)

    # ── Summary ───────────────────────────────────────────────────────────────
    st.subheader("📊 สรุปผลทุกการแข่งขัน")
    st.dataframe(pd.DataFrame(st.session_state["summary_data"]), hide_index=True, use_container_width=True)

    if out_sheet_id:
        st.success(
            f"✅ บันทึกผลไปที่ "
            f"[Google Sheet](https://docs.google.com/spreadsheets/d/{out_sheet_id}) แล้ว"
        )

    # ── Detailed results per event ─────────────────────────────────────────────
    st.subheader("📋 รายละเอียดแต่ละการแข่งขัน")
    for ev_name, results in all_results.items():
        if not results:
            st.info(f"**{ev_name}** — ไม่มีข้อมูล")
            continue
        df        = pd.DataFrame(results, columns=OUTPUT_HEADER)
        confirmed = sum(1 for r in results if r[5] == "✅")
        warned    = (df["สถานะ"] == "⚠️").sum()

        confirmed_df = df[df["สถานะ"] == "✅"].sort_values("#").reset_index(drop=True)

        with st.expander(
            f"**{ev_name}** — {confirmed}/{len(results)} ยืนยัน"
            + (f" | ⚠️ {warned}" if warned else ""),
            expanded=True,
        ):
            slip_label    = f"⚠️ ตรวจสลิป ({warned})" if warned else "⚠️ ตรวจสลิป"
            tab_r, tab_s, tab_a, tab_e = st.tabs(
                ["📊 ผลการตรวจ", slip_label, "📢 ประกาศ", "📧 อีเมล"]
            )

            # ── Tab 1: ผลการตรวจ ──────────────────────────────────────────────
            with tab_r:
                st.dataframe(
                    df.style.apply(style_row, axis=1),
                    use_container_width=True,
                    hide_index=True,
                )

            # ── Tab 2: ตรวจสลิป ───────────────────────────────────────────────
            with tab_s:
                warn_df = df.loc[df["สถานะ"] == "⚠️",
                                 ["#", "ชื่อที่ใช้แข่ง", "ชื่อบัญชีที่โอน", "รายละเอียด",
                                  "ลิงค์สลิป", "ตรวจสลิปแล้ว"]].copy()
                if warn_df.empty:
                    st.success("ไม่มีแถวที่ต้องตรวจสลิปเพิ่มเติม")
                else:
                    st.caption(f"{len(warn_df)} แถว — กรอกช่อง **'ตรวจสลิปแล้ว'** แล้วกด บันทึก")
                    edited_warn = st.data_editor(
                        warn_df,
                        key=f"warn_{ev_name}_{run_count}_{save_count}",
                        disabled=["#", "ชื่อที่ใช้แข่ง", "ชื่อบัญชีที่โอน", "รายละเอียด", "ลิงค์สลิป"],
                        hide_index=True,
                        use_container_width=True,
                        column_config={
                            "ลิงค์สลิป": st.column_config.LinkColumn("ลิงค์สลิป", display_text="เปิดสลิป"),
                            "ตรวจสลิปแล้ว": st.column_config.TextColumn(
                                "ตรวจสลิปแล้ว ✏️",
                                help="กรอกเพื่อยืนยัน เช่น ✅ หรือชื่อ admin",
                            ),
                        },
                    )
                    if st.button("💾 บันทึก", key=f"save_{ev_name}_{run_count}_{save_count}"):
                        updated_df = df.copy()
                        for _, edit_row in edited_warn.iterrows():
                            seq      = edit_row["#"]
                            override = str(edit_row.get("ตรวจสลิปแล้ว", "")).strip()
                            mask     = updated_df["#"] == seq
                            if not mask.any():
                                continue
                            updated_df.loc[mask, "ตรวจสลิปแล้ว"] = override
                            if override and updated_df.loc[mask, "สถานะ"].values[0] == "⚠️":
                                updated_df.loc[mask, "สถานะ"]      = "✅"
                                updated_df.loc[mask, "รายละเอียด"] = (
                                    "ตรวจสลิปแล้ว | " + updated_df.loc[mask, "รายละเอียด"].values[0]
                                )
                        updated_df = updated_df.fillna("")
                        st.session_state["all_results"][ev_name] = updated_df.values.tolist()
                        st.session_state["save_count"] = save_count + 1
                        if out_sheet_id:
                            try:
                                _gc    = get_gc()
                                _sheet = _gc.open_by_key(out_sheet_id)
                                ws     = _sheet.worksheet(ev_name)
                                ws.clear()
                                ws.append_row(OUTPUT_HEADER)
                                ws.append_rows(updated_df.values.tolist())
                                st.success("✅ บันทึกแล้ว")
                            except Exception as e:
                                st.error(f"บันทึกไม่ได้: {e}")
                        else:
                            st.success("✅ บันทึกในหน้านี้แล้ว")
                        st.rerun()

            # ── Tab 3: ประกาศ ─────────────────────────────────────────────────
            with tab_a:
                if confirmed_df.empty:
                    st.info("ยังไม่มีผู้ผ่านการยืนยัน")
                else:
                    col_fb, col_reg = st.columns([3, 1])
                    with col_fb:
                        post_lines = [
                            f"🎮 ประกาศรายชื่อผู้เข้าแข่งขัน {ev_name}",
                            f"มีผู้ผ่านการยืนยัน {len(confirmed_df)} คน\n",
                        ]
                        for i, (_, row) in enumerate(confirmed_df.iterrows(), 1):
                            post_lines.append(f"{i}. {row['ชื่อที่ใช้แข่ง']}")
                        post_lines.append(f"\n#WAKA #{ev_name.replace(' ', '')}")
                        st.text_area(
                            "📢 โพสต์ Facebook (คัดลอกได้เลย):",
                            value="\n".join(post_lines),
                            height=280,
                            key=f"fb_{ev_name}_{run_count}_{save_count}",
                        )
                    with col_reg:
                        st.write("**📋 ใบลงทะเบียน**")
                        st.caption(f"{len(confirmed_df)} คน | เรียงตามลำดับสมัคร")
                        if st.button("สร้างใน Google Sheet", key=f"reg_{ev_name}_{run_count}_{save_count}"):
                            if out_sheet_id:
                                try:
                                    _gc      = get_gc()
                                    _sheet   = _gc.open_by_key(out_sheet_id)
                                    ws_title = f"ลงทะเบียน — {ev_name}"
                                    try:
                                        ws_reg = _sheet.worksheet(ws_title)
                                        ws_reg.clear()
                                    except gspread.WorksheetNotFound:
                                        ws_reg = _sheet.add_worksheet(title=ws_title, rows=200, cols=4)
                                    ws_reg.append_row(["ลำดับ", "ชื่อที่ใช้แข่ง", "เช็คอิน ✓", "หมายเหตุ"])
                                    ws_reg.append_rows([
                                        [i, row["ชื่อที่ใช้แข่ง"], "", ""]
                                        for i, (_, row) in enumerate(confirmed_df.iterrows(), 1)
                                    ])
                                    st.success(f"✅ สร้างแล้ว — tab '{ws_title}'")
                                except Exception as e:
                                    st.error(f"สร้างไม่ได้: {e}")
                            else:
                                st.warning("ต้องระบุ Output Sheet URL ก่อน")

            # ── Tab 4: อีเมล ──────────────────────────────────────────────────
            with tab_e:
                if confirmed_df.empty:
                    st.info("ยังไม่มีผู้ผ่านการยืนยัน")
                else:
                    ev_emails = (st.session_state.get("all_emails") or {}).get(ev_name, {})
                    sent_key  = f"sent_{ev_name}"
                    sent_set  = st.session_state.get(sent_key, set())
                    recipients = [
                        (int(row["#"]), row["ชื่อที่ใช้แข่ง"],
                         ev_emails.get(row["ชื่อที่ใช้แข่ง"], ""))
                        for _, row in confirmed_df.iterrows()
                    ]
                    has_emails = any(e for _, _, e in recipients)

                    if not has_emails:
                        st.info("ไม่พบช่องอีเมลในฟอร์ม — เพิ่มคำถาม 'อีเมล' ในฟอร์มแล้วรันใหม่")
                    elif not GMAIL_ADDRESS or not GMAIL_APP_PWD:
                        st.warning(
                            "ตั้งค่าใน `.env` (local) หรือ Streamlit Secrets (cloud) ก่อน:\n"
                            "```\nGMAIL_ADDRESS=xxx@gmail.com\nGMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx\n```"
                        )
                    else:
                        ev_meta = (st.session_state.get("events_meta") or {}).get(ev_name, {})
                        col_d, col_t = st.columns(2)
                        event_date = col_d.text_input(
                            "วันแข่งขัน", value=ev_meta.get("date", ""),
                            placeholder="เช่น 28 มิ.ย. 69",
                            key=f"edate_{ev_name}_{run_count}",
                        )
                        event_time = col_t.text_input(
                            "เวลานัด", value=ev_meta.get("time", ""),
                            placeholder="เช่น 10:00 น.",
                            key=f"etime_{ev_name}_{run_count}",
                        )
                        st.divider()
                        with_email = [(n, name, em) for n, name, em in recipients if em]
                        already    = sum(1 for _, name, _ in with_email if name in sent_set)
                        check_all  = st.checkbox(
                            "เลือกทั้งหมด",
                            value=True,
                            key=f"chkall_{ev_name}_{run_count}_{save_count}",
                        )
                        recipient_df = pd.DataFrame([
                            {
                                "ส่ง":            check_all if name not in sent_set else False,
                                "#":              n,
                                "ชื่อที่ใช้แข่ง": name,
                                "อีเมล":          em,
                                "":               "ส่งแล้ว ✓" if name in sent_set else "",
                            }
                            for n, name, em in with_email
                        ])
                        edited = st.data_editor(
                            recipient_df,
                            key=f"email_ed_{ev_name}_{run_count}_{save_count}_{check_all}",
                            column_config={
                                "ส่ง":            st.column_config.CheckboxColumn("ส่ง", width="small"),
                                "#":              st.column_config.NumberColumn("#", width="small"),
                                "ชื่อที่ใช้แข่ง": st.column_config.TextColumn("ชื่อที่ใช้แข่ง"),
                                "อีเมล":          st.column_config.TextColumn("อีเมล"),
                                "":               st.column_config.TextColumn("", width="small"),
                            },
                            disabled=["#", "ชื่อที่ใช้แข่ง", "อีเมล", ""],
                            hide_index=True,
                            use_container_width=True,
                        )
                        selected = [
                            (int(row["#"]), row["ชื่อที่ใช้แข่ง"], row["อีเมล"])
                            for _, row in edited.iterrows() if row["ส่ง"]
                        ]
                        st.caption(f"เลือก **{len(selected)}** คน | ส่งแล้ว {already} คน")
                        if selected:
                            if st.button(
                                f"📧 ส่งอีเมล {len(selected)} คน",
                                key=f"email_send_{ev_name}_{run_count}_{save_count}",
                            ):
                                errors = []
                                prog   = st.progress(0, text="กำลังส่ง...")
                                for i, (n, name, em) in enumerate(selected):
                                    try:
                                        send_confirmation_email(em, ev_name, name, n, event_date, event_time)
                                        sent_set.add(name)
                                    except Exception as e:
                                        errors.append(f"{name}: {e}")
                                    prog.progress((i + 1) / len(selected), text=f"ส่งถึง {name}...")
                                st.session_state[sent_key] = sent_set
                                if errors:
                                    st.error("ส่งไม่ได้บางส่วน:\n" + "\n".join(errors))
                                else:
                                    st.success(f"✅ ส่งอีเมลครบ {len(selected)} คนแล้ว")
                                st.rerun()
