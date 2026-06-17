#!/usr/bin/env python3
"""
WAKA Tournament — Payment Verification Web App
วิธีรัน: python -m streamlit run tools/verify_app.py
"""

import csv
import io
import json
import re
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

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

FORM_COLUMN_KEYWORDS = {
    "game_name":     ["in game name", "ingame", "แข่งในวงการ", "ชื่อแข่ง", "trainer id", "openchat"],
    "openchat_name": ["openchat", "open chat"],
    "facebook":      ["facebook", "เฟสบุค", "fb"],
    "transfer_name": ["ชื่อบัญชี", "ชื่อที่โอน", "ชื่อโอน", "ใช้โอน", "ใช้ในการโอน",
                      "transfer name", "ชื่อเจ้าของบัญชี", "ชื่อที่ใช้"],
    "slip_url":      ["สลิป", "slip", "หลักฐาน", "การชำระ", "payment", "อัพโหลด"],
}

OUTPUT_HEADER = [
    "#", "ชื่อที่ใช้แข่ง", "ชื่อใน OpenChat", "ชื่อเฟสบุค", "ชื่อบัญชีที่โอน",
    "สถานะ", "รายละเอียด", "ยอดที่พบ", "ชื่อในธนาคาร", "เลขที่รายการ", "วันที่โอน", "ลิงค์สลิป",
]

# Pre-filled defaults (ใช้ครั้งแรก ถ้ายังไม่มี events_config.json)
DEFAULT_EVENTS = [
    {COL_NAME: "Lorcana",           COL_URL: "https://docs.google.com/spreadsheets/d/1vEzBdnQ1doPH3KLADSN9XbrC5tViZiGq1kLWET-6LeQ/edit#gid=118500649",  COL_FEE: "350"},
    {COL_NAME: "Riftbound",         COL_URL: "https://docs.google.com/spreadsheets/d/1EIqD91sJGBpwN0AeU9w8ior67elMIQS1vyY2mwShLZU/edit#gid=935532850",  COL_FEE: "500,900"},
    {COL_NAME: "Pokemon Champions", COL_URL: "https://docs.google.com/spreadsheets/d/1HLZ_tM3WbCiPP6BtuHSKmBo2H5KkYizsh71lJ4_moJA/edit#gid=2026372639", COL_FEE: "200"},
    {COL_NAME: "Pokemon TCG",       COL_URL: "https://docs.google.com/spreadsheets/d/1eHQQ-8ANDJfMcUyVwtrtumfgrtnTaiZr--wzUFExMpE/edit#gid=759768958",  COL_FEE: "250"},
    {COL_NAME: "BOT",               COL_URL: "https://docs.google.com/spreadsheets/d/1K7AFCxOOPzw-kp0kzUf07d-96dCsKnHjkYbIemzxeDM/edit#gid=2054056416", COL_FEE: "250"},
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
    if "GOOGLE_TOKEN" in st.secrets:
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

    parsed = []
    for seq, row in enumerate(data_rows, start=1):
        def cell(key, _row=row):
            i = cols.get(key)
            return _row[i].strip() if i is not None and i < len(_row) else ""
        game_name = cell("game_name") or cell("openchat_name")
        if game_name:
            parsed.append((seq, game_name, cell("openchat_name"), cell("facebook"),
                           cell("transfer_name"), cell("slip_url")))

    # Pass 1: lock ✅ (ชื่อ + ยอดตรง)
    exact_matches, txn_to_owner = {}, {}
    for seq, game_name, oc, fb, tr, slip in parsed:
        txn, status, detail = find_matching_txn(tr, fb, expected_fees, bank_rows, used_txn_ids)
        if status == "✅" and txn:
            used_txn_ids.add(txn["txn_id"])
            txn_to_owner[txn["txn_id"]] = game_name
            exact_matches[seq] = (txn, status, detail)

    # Pass 2: ⚠️ โดยเรียงความคล้ายชื่อ
    candidates = []
    for seq, game_name, oc, fb, tr, slip in parsed:
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
        detail = f"ตรวจสอบชื่อ | ยอดตรง {txn['amount']:.0f}฿ | ธนาคาร: {txn['sender']} | บัญชี: {match_name}"
        warn_matches[seq] = (txn, "⚠️", detail)
        used_txn_ids.add(txn["txn_id"])
        txn_to_owner[txn["txn_id"]] = gname

    # Compile
    results = []
    for seq, game_name, oc, fb, tr, slip in parsed:
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

        results.append([
            seq, game_name, oc, fb, tr, status, detail,
            f"{txn['amount']:.0f}" if txn else "",
            txn["sender"] if txn else "",
            txn["txn_id"] if txn else "",
            txn["date"]   if txn else "",
            slip,
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

    confirmed = sum(1 for r in results if r[5] == "✅")
    warned    = sum(1 for r in results if r[5] == "⚠️")
    duped     = sum(1 for r in results if r[5] == "🚫")
    failed    = len(results) - confirmed - warned - duped
    return results, confirmed, warned, duped, failed

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
for col in [COL_NAME, COL_URL, COL_FEE]:
    if col not in events_df.columns:
        events_df[col] = ""

edited_df = st.data_editor(
    events_df[[COL_NAME, COL_URL, COL_FEE]],
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        COL_NAME: st.column_config.TextColumn(COL_NAME, width="medium"),
        COL_URL:  st.column_config.TextColumn(COL_URL, width="large"),
        COL_FEE:  st.column_config.TextColumn(COL_FEE, width="small"),
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
    progress_bar = st.progress(0, text="กำลังประมวลผล...")

    for idx, event in enumerate(events):
        progress_bar.progress(idx / len(events), text=f"กำลังตรวจ {event['name']}...")
        try:
            results, ok, warn, dup, fail = process_event(
                event, gc, bank_rows, used_txn_ids, output_sheet
            )
            all_results[event["name"]] = results
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

    # ── Summary ───────────────────────────────────────────────────────────────
    st.subheader("📊 สรุปผลทุกการแข่งขัน")
    st.dataframe(pd.DataFrame(summary), hide_index=True, use_container_width=True)

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
        confirmed = sum(1 for r in results if r[5] == "✅")
        with st.expander(
            f"**{ev_name}** — {confirmed}/{len(results)} ยืนยัน",
            expanded=True,
        ):
            df = pd.DataFrame(results, columns=OUTPUT_HEADER)
            st.dataframe(
                df.style.apply(style_row, axis=1),
                use_container_width=True,
                hide_index=True,
            )
