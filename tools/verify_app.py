#!/usr/bin/env python3
"""
WAKA Tournament — Payment Verification Web App
วิธีรัน: python -m streamlit run tools/verify_app.py
"""

import base64
import csv
from datetime import date as _date_cls
import io
import json
import os
import re
import smtplib
import time
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

try:
    import qrcode as _qr_lib
    HAS_QR = True
except ImportError:
    HAS_QR = False

try:
    from streamlit_qrcode_scanner import qrcode_scanner as _qr_scanner
    HAS_SCANNER = True
except ImportError:
    HAS_SCANNER = False

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

# ── QR helper ─────────────────────────────────────────────────────────────────

def generate_qr_b64(data: str) -> str | None:
    if not HAS_QR:
        return None
    qr = _qr_lib.QRCode(box_size=6, border=2)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

# ── Paths & constants ─────────────────────────────────────────────────────────

SCOPES           = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/forms.body",
]
TOKEN_PATH       = Path("token.json")
CREDS_PATH       = Path("credentials.json")
CONFIG_PATH      = Path("events_config.json")
CONFIG_SHEET_TAB = "_config"

COL_NAME       = "ชื่อการแข่งขัน"
COL_URL        = "ลิงค์ Form Responses"
COL_FEE        = "ค่าสมัคร (฿)"
COL_WALKIN_FEE = "ค่าสมัคร หน้างาน (฿)"
COL_DATE       = "วันแข่งขัน"
COL_TIME       = "เวลานัด"
COL_VENUE      = "สถานที่"

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
    {COL_NAME: "Lorcana",           COL_URL: "https://docs.google.com/spreadsheets/d/1vEzBdnQ1doPH3KLADSN9XbrC5tViZiGq1kLWET-6LeQ/edit#gid=118500649",  COL_FEE: "350",     COL_DATE: "", COL_TIME: "", COL_VENUE: ""},
    {COL_NAME: "Riftbound",         COL_URL: "https://docs.google.com/spreadsheets/d/1EIqD91sJGBpwN0AeU9w8ior67elMIQS1vyY2mwShLZU/edit#gid=935532850",  COL_FEE: "500,900", COL_DATE: "", COL_TIME: "", COL_VENUE: ""},
    {COL_NAME: "Pokemon Champions", COL_URL: "https://docs.google.com/spreadsheets/d/1HLZ_tM3WbCiPP6BtuHSKmBo2H5KkYizsh71lJ4_moJA/edit#gid=2026372639", COL_FEE: "200",     COL_DATE: "", COL_TIME: "", COL_VENUE: ""},
    {COL_NAME: "Pokemon TCG",       COL_URL: "https://docs.google.com/spreadsheets/d/1eHQQ-8ANDJfMcUyVwtrtumfgrtnTaiZr--wzUFExMpE/edit#gid=759768958",  COL_FEE: "250",     COL_DATE: "", COL_TIME: "", COL_VENUE: ""},
    {COL_NAME: "BOT",               COL_URL: "https://docs.google.com/spreadsheets/d/1K7AFCxOOPzw-kp0kzUf07d-96dCsKnHjkYbIemzxeDM/edit#gid=2054056416", COL_FEE: "250",     COL_DATE: "", COL_TIME: "", COL_VENUE: ""},
]
DEFAULT_OUTPUT_URL = "https://docs.google.com/spreadsheets/d/1WSfd9sKHl2H5O7Ai1DvqVaL7Tuwni-nfBc-hTX8HilA"

# ── Config helpers ─────────────────────────────────────────────────────────────

def _bootstrap_output_url() -> str:
    """อ่าน output sheet URL จาก Secrets (cloud) หรือไฟล์ local — ใช้ bootstrap load_config"""
    try:
        url = st.secrets.get("OUTPUT_SHEET_URL", "") or ""
        if url.strip():
            return url.strip()
    except Exception:
        pass
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8")).get("output_sheet_url", "")
        except Exception:
            pass
    return DEFAULT_OUTPUT_URL

def load_config() -> dict:
    """โหลด config จาก Google Sheet (_config tab) ก่อน ถ้าไม่ได้ค่อย fallback ไปไฟล์ local"""
    bootstrap_url = _bootstrap_output_url()
    if bootstrap_url:
        try:
            sid, _ = parse_sheet_url(bootstrap_url)
            ws  = get_gc().open_by_key(sid).worksheet(CONFIG_SHEET_TAB)
            raw = ws.cell(1, 1).value
            if raw:
                return json.loads(raw)
        except Exception:
            pass
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"events": DEFAULT_EVENTS, "output_sheet_url": bootstrap_url or DEFAULT_OUTPUT_URL}

def save_config(events_records: list, output_url: str) -> str | None:
    """บันทึก config ลงทั้ง local file และ Google Sheet (_config tab)
    คืนค่า error message ถ้า Sheet save ล้มเหลว, None ถ้าสำเร็จทั้งคู่"""
    data    = {"events": events_records, "output_sheet_url": output_url}
    payload = json.dumps(data, ensure_ascii=False)
    CONFIG_PATH.write_text(payload, encoding="utf-8")
    try:
        sid, _ = parse_sheet_url(output_url)
        sht = get_gc().open_by_key(sid)
        try:
            ws = sht.worksheet(CONFIG_SHEET_TAB)
        except gspread.WorksheetNotFound:
            ws = sht.add_worksheet(title=CONFIG_SHEET_TAB, rows=5, cols=2)
        ws.update_cell(1, 1, payload)
        return None
    except Exception as e:
        return str(e)

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
def _get_creds() -> "Credentials":
    """คืน raw Google Credentials — ใช้ร่วมกันทั้ง Sheets และ Forms"""
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
        return creds

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
    return creds

@st.cache_resource
def get_gc():
    return gspread.authorize(_get_creds())

@st.cache_resource
def get_forms_service():
    from googleapiclient.discovery import build
    return build("forms", "v1", credentials=_get_creds())

def _has_forms_scope() -> bool:
    """ตรวจ token ปัจจุบันว่ามี forms.body scope หรือเปล่า (non-cached)"""
    try:
        has_secrets = "GOOGLE_TOKEN" in st.secrets
    except Exception:
        has_secrets = False
    if has_secrets:
        try:
            scopes = json.loads(st.secrets["GOOGLE_TOKEN"]).get("scopes", [])
            return any("forms" in s for s in scopes)
        except Exception:
            return False
    if TOKEN_PATH.exists():
        try:
            scopes = json.loads(TOKEN_PATH.read_text()).get("scopes", [])
            return any("forms" in s for s in scopes)
        except Exception:
            return False
    return False

# ── Sheets retry helper ────────────────────────────────────────────────────────

def _retry_429(fn):
    """เรียก fn() อีกครั้งถ้าเจอ 429 — รอ 15/30 วิ แล้ว retry สูงสุด 3 ครั้ง"""
    for attempt in range(3):
        try:
            return fn()
        except Exception as e:
            if "429" in str(e) and attempt < 2:
                time.sleep(15 * (2 ** attempt))  # 15s, 30s
            else:
                raise

# ── Google Forms creation ──────────────────────────────────────────────────────

def create_google_form(event_name: str, fee_str: str) -> dict:
    """สร้าง Google Form สำหรับรับสมัคร คืน {form_id, form_url, edit_url}"""
    svc = get_forms_service()

    form = svc.forms().create(body={
        "info": {
            "title":         f"สมัครแข่งขัน {event_name}",
            "documentTitle": f"สมัครแข่งขัน {event_name}",
        }
    }).execute()
    form_id = form["formId"]

    questions = [
        {
            "title":       "ชื่อที่ใช้แข่ง (In Game Name / Trainer ID)",
            "description": "ชื่อที่จะแสดงในตาราง bracket",
            "required":    True,
        },
        {
            "title":       "ชื่อบัญชีธนาคารที่ใช้โอนเงิน",
            "description": "ชื่อที่ปรากฏในแอปธนาคาร — ใช้ตรวจสอบการชำระ",
            "required":    True,
        },
        {
            "title":       "ชื่อ Facebook",
            "description": "สำหรับติดต่อในกรณีจำเป็น",
            "required":    False,
        },
        {
            "title":       "อีเมล (สำหรับรับ QR Code เข้างาน)",
            "description": "จะใช้ส่ง QR Code ยืนยันการสมัครให้คุณ",
            "required":    True,
        },
        {
            "title":       f"ลิงก์สลิปการโอนเงิน ค่าสมัคร {fee_str}฿",
            "description": (
                "1) โอนเงินให้เรียบร้อย\n"
                "2) ถ่ายรูปหรือ screenshot สลิป แล้วอัปโหลดลง Google Drive\n"
                "3) คลิกขวา → Share → 'Anyone with the link' → Copy link มาวางที่นี่"
            ),
            "required":    True,
        },
    ]

    requests = [
        {
            "createItem": {
                "item": {
                    "title":       q["title"],
                    "description": q["description"],
                    "questionItem": {
                        "question": {
                            "required":     q["required"],
                            "textQuestion": {"paragraph": False},
                        }
                    },
                },
                "location": {"index": i},
            }
        }
        for i, q in enumerate(questions)
    ]

    svc.forms().batchUpdate(
        formId=form_id,
        body={"requests": requests},
    ).execute()

    return {
        "form_id":  form_id,
        "form_url": f"https://docs.google.com/forms/d/{form_id}/viewform",
        "edit_url": f"https://docs.google.com/forms/d/{form_id}/edit",
    }

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


def parse_txn_date(date_str: str) -> "_date_cls | None":
    """DD/MM/YYYY[...] ปี พ.ศ. → CE date  (2569 → 2026)  หรือ None ถ้า parse ไม่ได้"""
    if not date_str:
        return None
    try:
        d = date_str.strip().split()[0]
        day, month, year = d.split("/")
        year_ce = int(year) - 543
        if year_ce < 2000:      # อาจเป็น ค.ศ. อยู่แล้ว เช่น 2026
            year_ce = int(year)
        return _date_cls(year_ce, int(month), int(day))
    except Exception:
        return None

def process_event(event, gc, bank_rows, used_txn_ids, output_sheet=None):
    name          = event["name"]
    expected_fees = event["fee"] if isinstance(event["fee"], list) else [float(event["fee"])]

    src_sh = gc.open_by_key(event["source_sheet_id"])
    src_ws = src_sh.get_worksheet_by_id(event["gid"]) if "gid" in event else src_sh.sheet1
    rows   = _retry_429(src_ws.get_all_values)

    if len(rows) < 2:
        return [], 0, 0, 0, 0

    header, data_rows = rows[0], rows[1:]
    cols = detect_columns(header, FORM_COLUMN_KEYWORDS)

    # Fallback: ถ้าหา slip_url ด้วย keyword ไม่ได้ ให้สแกนหาคอลัมน์ที่มีลิงก์ Drive
    if "slip_url" not in cols and data_rows:
        for col_i in range(len(header)):
            for row in data_rows[:15]:
                val = row[col_i] if col_i < len(row) else ""
                if "drive.google.com" in val or ("docs.google.com" in val and "/file/" in val):
                    cols["slip_url"] = col_i
                    break
            if "slip_url" in cols:
                break

    # อ่านค่า "ตรวจสลิปแล้ว" ที่ admin กรอกไว้ก่อน clear
    manual_ok = {}  # game_name → note
    if output_sheet:
        try:
            ws_prev = output_sheet.worksheet(name)
            prev    = _retry_429(ws_prev.get_all_values)
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

    # กรอง bank_rows เฉพาะ transaction ไม่เก่ากว่า 3 วัน นับจาก transaction ล่าสุดใน statement
    _txn_dates = [parse_txn_date(t["date"]) for t in bank_rows]
    _ref_date  = max((d for d in _txn_dates if d), default=None)
    MAX_TXN_AGE_DAYS = 3
    if _ref_date:
        recent_bank_rows = [
            t for t in bank_rows
            if (lambda d: d is None or (_ref_date - d).days <= MAX_TXN_AGE_DAYS)(parse_txn_date(t["date"]))
        ]
    else:
        recent_bank_rows = bank_rows

    # Pass 1: lock ✅ (ชื่อ + ยอดตรง)
    exact_matches, txn_to_owner = {}, {}
    for seq, game_name, oc, fb, tr, slip, email_addr in parsed:
        txn, status, detail = find_matching_txn(tr, fb, expected_fees, recent_bank_rows, used_txn_ids)
        if status == "✅" and txn:
            used_txn_ids.add(txn["txn_id"])
            txn_to_owner[txn["txn_id"]] = game_name
            exact_matches[seq] = (txn, status, detail)

    # Pass 2: ⚠️ โดยเรียงความคล้ายชื่อ (ต้องคล้ายกัน >= 30% จึง match)
    MIN_PASS2_SIM = 0.3
    candidates = []
    for seq, game_name, oc, fb, tr, slip, email_addr in parsed:
        if seq in exact_matches or game_name in manual_ok:
            continue
        match_name = tr if tr else fb
        for txn in recent_bank_rows:
            if txn["txn_id"] and txn["txn_id"] in used_txn_ids:
                continue
            if any(abs(txn["amount"] - f) < 1.0 for f in expected_fees):
                sim = name_similarity(match_name, txn["sender"])
                if sim >= MIN_PASS2_SIM:
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

    # Pass 3: ⚠️ ยอดไม่ตรง — ชื่อตรงแต่ยอดต่างจากค่าสมัคร (เช่น โอน 1000 แต่ fee 500)
    amount_mismatch = {}
    for seq, game_name, oc, fb, tr, slip, email_addr in parsed:
        if seq in exact_matches or seq in warn_matches or game_name in manual_ok:
            continue
        match_name = tr if tr else fb
        n_match = normalize(match_name)
        if not n_match:
            continue
        for txn in recent_bank_rows:
            if txn["txn_id"] and txn["txn_id"] in used_txn_ids:
                continue
            n_sender = normalize(txn["sender"])
            if n_sender and (n_match in n_sender or n_sender in n_match):
                fees_str = "/".join(f"{f:.0f}" for f in expected_fees)
                detail = f"ยอดไม่ตรง | พบ {txn['amount']:.0f}฿ (คาดยอด {fees_str}฿)"
                amount_mismatch[seq] = (txn, "⚠️", detail)
                used_txn_ids.add(txn["txn_id"])
                txn_to_owner[txn["txn_id"]] = game_name
                break

    # Compile
    results = []
    emails  = {}  # game_name → email address
    for seq, game_name, oc, fb, tr, slip, email_addr in parsed:
        if seq in exact_matches:
            txn, status, detail = exact_matches[seq]
        elif seq in warn_matches:
            txn, status, detail = warn_matches[seq]
        elif seq in amount_mismatch:
            txn, status, detail = amount_mismatch[seq]
        else:
            match_name = tr if tr else fb
            fees_str   = "/".join(f"{f:.0f}" for f in expected_fees)
            best_txn, best_sim = None, 0.0
            for t in recent_bank_rows:
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
        if override:
            status = "✅"
            if seq not in exact_matches:
                # ล้าง bank transaction ที่อาจ match ผิดคน
                detail = f"ตรวจสลิปแล้ว (admin)"
                txn = None
            else:
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
    event_date: str = "", event_time: str = "", event_venue: str = "",
) -> None:
    date_row  = f'<tr><td style="padding:4px 20px 4px 0;color:#555">วันแข่งขัน</td><td><strong>{event_date}</strong></td></tr>'  if event_date  else ""
    time_row  = f'<tr><td style="padding:4px 20px 4px 0;color:#555">เวลานัด</td><td><strong>{event_time}</strong></td></tr>'     if event_time  else ""
    venue_row = f'<tr><td style="padding:4px 20px 4px 0;color:#555">สถานที่</td><td><strong>{event_venue}</strong></td></tr>'    if event_venue else ""
    qr_data  = f"WAKA|{ev_name}|{game_name}|{order_num}"
    qr_b64   = generate_qr_b64(qr_data)
    qr_bytes = base64.b64decode(qr_b64) if qr_b64 else None
    qr_block = (
        '<div style="text-align:center;margin:20px 0">'
        '<p style="color:#555;font-size:13px;margin-bottom:8px">แสดง QR นี้ให้เจ้าหน้าที่สแกนเพื่อเช็คอินเข้างาน</p>'
        '<img src="cid:qrcode" width="160" height="160" style="border:1px solid #ddd;padding:4px"/>'
        '</div>'
    ) if qr_bytes else ""
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
    {venue_row}
  </table>
  {qr_block}
  <p>โปรดมาถึงสถานที่ก่อนเวลาแข่งขันครับ</p>
  <p style="margin-top:24px;color:#888;font-size:13px">— ทีม WAKA Tournament</p>
</div>
"""
    # ใช้ multipart/related เพื่อแนบรูป QR แบบ inline (email client ทุกตัวรองรับ)
    outer = MIMEMultipart("related")
    outer["Subject"] = f"[WAKA] ยืนยันการเข้าแข่งขัน {ev_name}"
    outer["From"]    = GMAIL_ADDRESS
    outer["To"]      = to_email
    outer.attach(MIMEText(body, "html"))
    if qr_bytes:
        img_part = MIMEImage(qr_bytes, "png")
        img_part.add_header("Content-ID", "<qrcode>")
        img_part.add_header("Content-Disposition", "inline", filename="qrcode.png")
        outer.attach(img_part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PWD)
        server.sendmail(GMAIL_ADDRESS, to_email, outer.as_string())

# ── Check-in sheet sync ───────────────────────────────────────────────────────

def _sync_checkin_sheet(out_sheet_id: str, ev_name: str, confirmed_df, ci_state: dict, name_filter=None):
    """Write ci_state to the registration sheet. name_filter=single name for fast single-row update."""
    _gc      = get_gc()
    _sheet   = _gc.open_by_key(out_sheet_id)
    ws_title = f"ลงทะเบียน — {ev_name}"
    try:
        ws_reg = _sheet.worksheet(ws_title)
    except gspread.WorksheetNotFound:
        if confirmed_df is None:
            raise gspread.WorksheetNotFound(f"ไม่พบชีต '{ws_title}' — สร้างจาก tab ประกาศก่อน")
        ws_reg = _sheet.add_worksheet(title=ws_title, rows=200, cols=4)
        ws_reg.append_row(["ลำดับ", "ชื่อที่ใช้แข่ง", "เช็คอิน ✓", "หมายเหตุ"])
        ws_reg.append_rows([
            [i, r["ชื่อที่ใช้แข่ง"], "", ""]
            for i, (_, r) in enumerate(confirmed_df.iterrows(), 1)
        ])
    reg_rows = _retry_429(ws_reg.get_all_values)
    if len(reg_rows) < 2:
        return
    hdr    = reg_rows[0]
    n_idx  = next((i for i, h in enumerate(hdr) if h == "ชื่อที่ใช้แข่ง"), 1)
    ci_idx = next((i for i, h in enumerate(hdr) if "เช็คอิน" in h), 2)
    for r_idx, r in enumerate(reg_rows[1:], start=2):
        pname = r[n_idx] if n_idx < len(r) else ""
        if name_filter and pname != name_filter:
            continue
        if pname in ci_state:
            ws_reg.update_cell(r_idx, ci_idx + 1, ci_state[pname])

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


def _load_from_sheet(edited_df, output_url: str, run_count: int) -> str | None:
    """โหลดผลจาก Google Sheet → session_state  คืน None=สำเร็จ, str=error message"""
    try:
        _gc  = get_gc()
        _oid, _ = parse_sheet_url(output_url)
        _sht = _gc.open_by_key(_oid)
        _loaded_results: dict = {}
        _loaded_emails:  dict = {}
        _loaded_summary: list = []
        for _, _erow in edited_df.iterrows():
            _en = str(_erow.get(COL_NAME, "")).strip()
            _eu = str(_erow.get(COL_URL,  "")).strip()
            if not _en:
                continue
            try:
                _ws   = _sht.worksheet(_en)
                _rows = _retry_429(_ws.get_all_values)
                _loaded_results[_en] = _rows[1:] if len(_rows) > 1 else []
            except gspread.WorksheetNotFound:
                _loaded_results[_en] = []
            _em: dict[str, str] = {}
            if _eu:
                try:
                    _sid, _gid = parse_sheet_url(_eu)
                    _src = _gc.open_by_key(_sid)
                    _sw  = _src.get_worksheet_by_id(_gid) if _gid is not None else _src.sheet1
                    _sr  = _retry_429(_sw.get_all_values)
                    if len(_sr) > 1:
                        _sc = detect_columns(_sr[0], FORM_COLUMN_KEYWORDS)
                        for _r in _sr[1:]:
                            def _cv(k, _row=_r, _cols=_sc):
                                _i = _cols.get(k)
                                return _row[_i].strip() if _i is not None and _i < len(_row) else ""
                            _gn = _cv("game_name") or _cv("openchat_name")
                            _ea = _cv("email")
                            if _gn and _ea:
                                _em[_gn] = _ea
                except Exception:
                    pass
            _loaded_emails[_en] = _em
            _d  = _loaded_results[_en]
            _ok = sum(1 for _r in _d if len(_r) > 5 and _r[5] == "✅")
            _wn = sum(1 for _r in _d if len(_r) > 5 and _r[5] == "⚠️")
            _dp = sum(1 for _r in _d if len(_r) > 5 and _r[5] == "🚫")
            _fl = sum(1 for _r in _d if len(_r) > 5 and _r[5] == "❌")
            _loaded_summary.append({
                "การแข่งขัน": _en,
                "✅ ยืนยัน": _ok, "⚠️ ตรวจสอบ": _wn,
                "🚫 ซ้ำ": _dp, "❌ ไม่พบ": _fl, "ทั้งหมด": len(_d),
            })
        st.session_state["all_results"]      = _loaded_results
        st.session_state["all_emails"]       = _loaded_emails
        st.session_state["summary_data"]     = _loaded_summary
        st.session_state["out_sheet_id_run"] = _oid
        st.session_state["events_meta"]      = {
            str(_r.get(COL_NAME, "")).strip(): {
                "date":  str(_r.get(COL_DATE,  "")).strip(),
                "time":  str(_r.get(COL_TIME,  "")).strip(),
                "venue": str(_r.get(COL_VENUE, "")).strip(),
            }
            for _, _r in edited_df.iterrows()
            if str(_r.get(COL_NAME, "")).strip()
        }
        st.session_state["run_count"]  = run_count + 1
        st.session_state["save_count"] = 0
        return None
    except Exception as _e:
        return str(_e)


# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="WAKA Tournament", page_icon="🎮", layout="wide")
st.title("🎮 WAKA Tournament — ตรวจสอบการชำระเงิน")

config = load_config()

# เตรียม DataFrame ก่อน tabs (ใช้ร่วมกันทั้งสอง tab)
saved_events = config.get("events") or DEFAULT_EVENTS
events_df    = pd.DataFrame(saved_events)
for col in [COL_NAME, COL_URL, COL_FEE, COL_WALKIN_FEE, COL_DATE, COL_TIME, COL_VENUE]:
    if col not in events_df.columns:
        events_df[col] = ""

tab_settings, tab_verify, tab_list = st.tabs(["⚙️ ตั้งค่า", "🔍 ตรวจสลิป", "📋 รายชื่อ"])

# ─── Tab: ตั้งค่า ─────────────────────────────────────────────────────────────
with tab_settings:
    st.subheader("📋 การแข่งขัน")
    st.caption("เพิ่ม/แก้ไขแถว ได้เลย | ค่าสมัครหลายค่าคั่นด้วยคอมมา เช่น `500,900`")
    edited_df = st.data_editor(
        events_df[[COL_NAME, COL_URL, COL_FEE, COL_WALKIN_FEE, COL_DATE, COL_TIME, COL_VENUE]],
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            COL_NAME:       st.column_config.TextColumn(COL_NAME,       width="medium"),
            COL_URL:        st.column_config.TextColumn(COL_URL,        width="large"),
            COL_FEE:        st.column_config.TextColumn(COL_FEE,        width="small", help="ล่วงหน้า — หลายค่าคั่นด้วย , เช่น 500,900"),
            COL_WALKIN_FEE: st.column_config.TextColumn(COL_WALKIN_FEE, width="small", help="ค่าสมัครวันงาน (ถ้าต่างจากล่วงหน้า)"),
            COL_DATE:       st.column_config.TextColumn(COL_DATE,       width="small", help="เช่น 28 มิ.ย. 69"),
            COL_TIME:       st.column_config.TextColumn(COL_TIME,       width="small", help="เช่น 10:00 น."),
            COL_VENUE:      st.column_config.TextColumn(COL_VENUE,      width="medium"),
        },
    )
    st.divider()
    st.subheader("📤 Output Sheet")
    output_url = st.text_input(
        "Google Sheet URL สำหรับบันทึกผล (ไม่บังคับ — ถ้าว่างจะแสดงเฉพาะในหน้านี้)",
        value=config.get("output_sheet_url", ""),
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
    with st.expander("ℹ️ วิธีให้ config คงอยู่บน Streamlit Cloud"):
        st.markdown(
            "เพิ่ม Secret ชื่อ `OUTPUT_SHEET_URL` ใน Streamlit Cloud dashboard "
            "(Settings → Secrets) แล้วใส่ URL ของ Output Sheet\n\n"
            "```toml\nOUTPUT_SHEET_URL = \"https://docs.google.com/spreadsheets/d/...\"\n```\n\n"
            "เมื่อตั้งค่าแล้ว กด 💾 บันทึกการตั้งค่า ครั้งเดียว — config จะถูกเก็บใน tab `_config` "
            "ของ Sheet นั้น และโหลดขึ้นมาอัตโนมัติทุกครั้งแม้ deploy ใหม่"
        )
    if st.button("💾 บันทึกการตั้งค่า"):
        err = save_config(edited_df.to_dict("records"), output_url)
        if err:
            st.warning(f"บันทึกลงไฟล์ local แล้ว แต่บันทึกลง Google Sheet ไม่ได้: {err}")
        else:
            st.success("✅ บันทึกแล้ว — config เก็บใน Google Sheet tab `_config` และไฟล์ local")

    st.divider()
    st.subheader("📝 สร้าง Google Form")
    _ev_names_cfg = [
        str(r.get(COL_NAME, "")).strip()
        for _, r in edited_df.iterrows()
        if str(r.get(COL_NAME, "")).strip()
    ]
    if not _ev_names_cfg:
        st.info("เพิ่มชื่อการแข่งขันในตารางด้านบนก่อน")
    else:
        _fc1, _fc2, _fc3 = st.columns([2, 1, 1])
        _sel_ev = _fc1.selectbox("การแข่งขัน", _ev_names_cfg, key="form_create_ev")
        _sel_row_cfg = next(
            (r for _, r in edited_df.iterrows() if str(r.get(COL_NAME, "")).strip() == _sel_ev),
            None,
        )
        _sel_fee_cfg = str(_sel_row_cfg.get(COL_FEE, "")).strip() if _sel_row_cfg is not None else ""
        _fc2.text_input("ค่าสมัคร", value=_sel_fee_cfg, disabled=True, key="form_create_fee_disp")

        if not _has_forms_scope():
            st.warning(
                "⚠️ Token ปัจจุบันไม่มีสิทธิ์สร้าง Form\n\n"
                "วิธีแก้ (ทำบนเครื่องตัวเองครั้งเดียว):\n"
                "1. ลบไฟล์ `token.json`\n"
                "2. รัน `python -m streamlit run tools/verify_app.py`\n"
                "3. เปิด browser แล้ว approve ให้ครบทั้ง Sheets และ Forms\n"
                "4. ถ้าใช้ Streamlit Cloud: copy เนื้อหา token.json ใหม่ไปอัปเดต Secret `GOOGLE_TOKEN`"
            )
        elif _fc3.button("➕ สร้าง Form", key="btn_create_form"):
            try:
                with st.spinner("⏳ กำลังสร้าง Google Form..."):
                    _form_result = create_google_form(_sel_ev, _sel_fee_cfg)
                st.success("✅ สร้าง Form แล้ว!")
                st.markdown(f"**ลิงก์แชร์ให้ผู้สมัคร:**  \n{_form_result['form_url']}")
                st.markdown(f"**ลิงก์แก้ไข Form:**  \n{_form_result['edit_url']}")
                st.info(
                    "📋 ขั้นตอนต่อไป:\n"
                    "1. เปิดลิงก์แก้ไขด้านบน → Responses → **Link to Sheets** → สร้าง Sheet ใหม่\n"
                    "2. copy URL ของ Sheet นั้น → วางในคอลัมน์ **ลิงค์ Form Responses** ด้านบน\n"
                    "3. กด 💾 บันทึกการตั้งค่า"
                )
            except Exception as _fe:
                st.error(f"สร้างไม่ได้: {_fe}")

# ─── Tab: ตรวจสลิป ───────────────────────────────────────────────────────────
with tab_verify:
    st.caption("🗓️ สำหรับผู้ลงทะเบียน**ล่วงหน้า**ผ่าน Google Form เท่านั้น | ผู้ลงทะเบียนหน้างานเพิ่มได้ใน tab 📋 รายชื่อ → ประกาศ")
    st.subheader("📁 ไฟล์รายงานธนาคาร")
    bank_files = st.file_uploader(
        "อัปโหลด PDF หรือ CSV จาก SCB แม่มณี (เลือกได้หลายไฟล์)",
        type=["pdf", "csv"],
        accept_multiple_files=True,
    )

    # เก็บเนื้อหาไฟล์ใน session_state ป้องกัน reset หลังกดปุ่ม
    if bank_files:
        st.session_state["bank_files"] = [(f.name, f.read()) for f in bank_files]

    has_file    = "bank_files" in st.session_state
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

        # Load bank files (รองรับหลายไฟล์)
        if "bank_files" not in st.session_state or not st.session_state["bank_files"]:
            st.error("กรุณาอัปโหลดไฟล์ธนาคารก่อน")
            st.stop()
        with st.spinner("กำลังอ่านไฟล์ธนาคาร..."):
            bank_rows = []
            seen_txn_ids: set[str] = set()
            errors_load = []
            for fname, content in st.session_state["bank_files"]:
                try:
                    if fname.lower().endswith(".pdf") or content[:4] == b"%PDF":
                        rows = load_bank_pdf(content)
                    else:
                        rows = load_bank_csv(content)
                    before = len(bank_rows)
                    for r in rows:
                        tid = r.get("txn_id", "")
                        if tid and tid in seen_txn_ids:
                            continue  # ข้าม transaction ซ้ำระหว่างไฟล์
                        if tid:
                            seen_txn_ids.add(tid)
                        bank_rows.append(r)
                    st.info(f"**{fname}** — {len(bank_rows) - before} transactions")
                except Exception as e:
                    errors_load.append(f"{fname}: {e}")
            if errors_load:
                st.error("อ่านไฟล์ไม่ได้:\n" + "\n".join(errors_load))
            if not bank_rows:
                st.stop()
            st.success(f"รวมทั้งหมด **{len(bank_rows)} transactions** จาก {len(st.session_state['bank_files'])} ไฟล์")

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
                "date":  str(row.get(COL_DATE,  "")).strip(),
                "time":  str(row.get(COL_TIME,  "")).strip(),
                "venue": str(row.get(COL_VENUE, "")).strip(),
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
            need_review = warned + (df["สถานะ"] == "❌").sum()

            confirmed_df = df[df["สถานะ"] == "✅"].sort_values("#").reset_index(drop=True)

            with st.expander(
                f"**{ev_name}** — {confirmed}/{len(results)} ยืนยัน"
                + (f" | ⚠️ {warned}" if warned else ""),
                expanded=True,
            ):
                slip_label    = f"⚠️ ตรวจสลิป ({need_review})" if need_review else "⚠️ ตรวจสลิป"
                tab_r, tab_s = st.tabs(["📊 ผลการตรวจ", slip_label])

                # ── Tab 1: ผลการตรวจ ──────────────────────────────────────────────
                with tab_r:
                    st.dataframe(
                        df.style.apply(style_row, axis=1),
                        use_container_width=True,
                        hide_index=True,
                    )

                # ── Tab 2: ตรวจสลิป ───────────────────────────────────────────────
                with tab_s:
                    # แสดง ⚠️ ทุกแถว + ❌ ทุกแถว (admin ตัดสินใจเองว่าจะ approve หรือไม่)
                    _needs_review = (df["สถานะ"] == "⚠️") | (df["สถานะ"] == "❌")
                    warn_df = df.loc[_needs_review,
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

# ─── Tab: รายชื่อ ─────────────────────────────────────────────────────────────
with tab_list:
    _list_results = st.session_state.get("all_results")
    _list_emails  = st.session_state.get("all_emails", {})
    _out_id_list  = st.session_state.get("out_sheet_id_run")
    _run_count    = st.session_state.get("run_count", 0)
    _save_count   = st.session_state.get("save_count", 0)

    # ── Auto-load ครั้งแรก + ปุ่มโหลดใหม่ ───────────────────────────────────────
    _col_reload, _col_pad = st.columns([1, 5])
    if _col_reload.button("🔄 โหลดใหม่", key="btn_reload_list"):
        for _k in ("all_results", "all_emails", "summary_data", "_load_retry_after"):
            st.session_state.pop(_k, None)
        st.rerun()

    if not _list_results:
        if not output_url.strip():
            st.info("ระบุ Output Sheet URL ในแท็บ ⚙️ ตั้งค่า แล้วกด 💾 บันทึกการตั้งค่า ก่อน")
        else:
            _retry_after = st.session_state.get("_load_retry_after", 0)
            _now = time.time()
            if _now < _retry_after:
                _wait = int(_retry_after - _now)
                st.warning(f"Google Sheets rate limit — รอ {_wait} วินาทีแล้วกด 🔄 โหลดใหม่")
            else:
                with st.spinner("⏳ กำลังโหลดรายชื่อจาก Sheet..."):
                    _load_err = _load_from_sheet(edited_df, output_url, _run_count)
                if _load_err:
                    if "429" in _load_err:
                        st.session_state["_load_retry_after"] = _now + 60
                        st.error("Google Sheets ถูกเรียกถี่เกินไป — รอ 1 นาทีแล้วกด 🔄 โหลดใหม่")
                    else:
                        st.error(f"โหลดไม่ได้: {_load_err}")
                else:
                    _list_results = st.session_state.get("all_results")
                    _list_emails  = st.session_state.get("all_emails", {})
                    _out_id_list  = st.session_state.get("out_sheet_id_run")
                    _run_count    = st.session_state.get("run_count", 0)
                    _save_count   = st.session_state.get("save_count", 0)

    # ── แสดงรายชื่อต่อเมื่อมีข้อมูล ──────────────────────────────────────────
    if _list_results:
        out_sheet_id = _out_id_list
        run_count    = _run_count
        save_count   = _save_count

        for ev_name, results in _list_results.items():
            if not results:
                st.info(f"**{ev_name}** — ยังไม่มีข้อมูล")
                continue
            df           = pd.DataFrame(results, columns=OUTPUT_HEADER)
            confirmed_df = df[df["สถานะ"] == "✅"].sort_values("#").reset_index(drop=True)
            confirmed    = len(confirmed_df)

            with st.expander(f"**{ev_name}** — {confirmed} คนยืนยัน", expanded=True):
                tab_a, tab_e, tab_c = st.tabs(["📢 ประกาศ", "📧 อีเมล", "🎫 เช็คอิน"])

                # ── ประกาศ ──────────────────────────────────────────────────────
                with tab_a:
                    _ev_row_a    = edited_df[edited_df[COL_NAME] == ev_name]
                    _pre_fee_a   = str(_ev_row_a[COL_FEE].values[0]).strip()        if not _ev_row_a.empty and COL_FEE        in _ev_row_a.columns else ""
                    _walkin_fee_a= str(_ev_row_a[COL_WALKIN_FEE].values[0]).strip() if not _ev_row_a.empty and COL_WALKIN_FEE in _ev_row_a.columns else ""
                    if _pre_fee_a or _walkin_fee_a:
                        _fee_parts = []
                        if _pre_fee_a:    _fee_parts.append(f"ล่วงหน้า: **{_pre_fee_a}฿**")
                        if _walkin_fee_a: _fee_parts.append(f"หน้างาน: **{_walkin_fee_a}฿**")
                        st.caption("🏷️ " + " | ".join(_fee_parts))

                    if confirmed_df.empty:
                        st.info("ยังไม่มีผู้ผ่านการยืนยัน (ล่วงหน้า)")
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
                                        # invalidate ci cache
                                        _cik = f"checkin_{ev_name}"
                                        st.session_state.pop(_cik, None)
                                        st.session_state.pop(f"{_cik}_order", None)
                                        st.success(f"✅ สร้างแล้ว — tab '{ws_title}'")
                                    except Exception as e:
                                        st.error(f"สร้างไม่ได้: {e}")
                                else:
                                    st.warning("ต้องระบุ Output Sheet URL ก่อน")

                    # ── ลงทะเบียนหน้างาน ────────────────────────────────────────
                    st.divider()
                    st.write("**➕ ลงทะเบียนหน้างาน**")
                    _walkin_name = st.text_input(
                        "ชื่อผู้แข่ง", placeholder="พิมพ์ชื่อ…",
                        key=f"walkin_{ev_name}_{run_count}_{save_count}",
                    )
                    if st.button("➕ เพิ่มหน้างาน", key=f"walkin_btn_{ev_name}_{run_count}_{save_count}"):
                        if not _walkin_name.strip():
                            st.warning("กรอกชื่อก่อน")
                        elif not out_sheet_id:
                            st.warning("ต้องระบุ Output Sheet URL ก่อน")
                        else:
                            try:
                                _gc_w    = get_gc()
                                _sht_w   = _gc_w.open_by_key(out_sheet_id)
                                _ws_t    = f"ลงทะเบียน — {ev_name}"
                                try:
                                    _ws_w    = _sht_w.worksheet(_ws_t)
                                    _exist   = _retry_429(_ws_w.get_all_values)
                                    _next_n  = len(_exist)  # header + n rows → next = n+1 = len
                                except gspread.WorksheetNotFound:
                                    _ws_w = _sht_w.add_worksheet(title=_ws_t, rows=200, cols=4)
                                    _ws_w.append_row(["ลำดับ", "ชื่อที่ใช้แข่ง", "เช็คอิน ✓", "หมายเหตุ"])
                                    _next_n = 1
                                _ws_w.append_row([_next_n, _walkin_name.strip(), "", "หน้างาน"])
                                _cik = f"checkin_{ev_name}"
                                st.session_state.pop(_cik, None)
                                st.session_state.pop(f"{_cik}_order", None)
                                st.success(f"✅ เพิ่ม **{_walkin_name.strip()}** ลำดับที่ {_next_n} (หน้างาน) แล้ว")
                                st.rerun()
                            except Exception as e:
                                st.error(f"เพิ่มไม่ได้: {e}")

                # ── อีเมล ────────────────────────────────────────────────────────
                with tab_e:
                    if confirmed_df.empty:
                        st.info("ยังไม่มีผู้ผ่านการยืนยัน")
                    else:
                        ev_emails = (_list_emails or {}).get(ev_name, {})
                        sent_key  = f"sent_{ev_name}"
                        sent_set  = st.session_state.get(sent_key, set())
                        recipients = [
                            (int(row["#"]), row["ชื่อที่ใช้แข่ง"],
                             ev_emails.get(row["ชื่อที่ใช้แข่ง"], ""))
                            for _, row in confirmed_df.iterrows()
                        ]
                        has_emails = any(e for _, _, e in recipients)

                        if not has_emails:
                            st.info("ไม่พบอีเมล — ลอง 📋 โหลดรายชื่อจาก Sheet อีกครั้ง")
                        elif not GMAIL_ADDRESS or not GMAIL_APP_PWD:
                            st.warning(
                                "ตั้งค่าใน `.env` (local) หรือ Streamlit Secrets (cloud) ก่อน:\n"
                                "```\nGMAIL_ADDRESS=xxx@gmail.com\nGMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx\n```"
                            )
                        else:
                            ev_meta   = (st.session_state.get("events_meta") or {}).get(ev_name, {})
                            _ev_row   = edited_df[edited_df[COL_NAME] == ev_name]
                            def _cfg(col):
                                v = _ev_row[col].values[0] if not _ev_row.empty and col in _ev_row.columns else ""
                                return str(v).strip()
                            col_d, col_t = st.columns(2)
                            event_date = col_d.text_input(
                                "วันแข่งขัน", value=ev_meta.get("date", _cfg(COL_DATE)),
                                placeholder="เช่น 28 มิ.ย. 69",
                                key=f"edate_{ev_name}_{run_count}",
                            )
                            event_time = col_t.text_input(
                                "เวลานัด", value=ev_meta.get("time", _cfg(COL_TIME)),
                                placeholder="เช่น 10:00 น.",
                                key=f"etime_{ev_name}_{run_count}",
                            )
                            event_venue = st.text_input(
                                "สถานที่", value=ev_meta.get("venue", _cfg(COL_VENUE)),
                                placeholder="เช่น ร้าน WAKA Game Shop",
                                key=f"evenue_{ev_name}_{run_count}",
                            )
                            if st.button("💾 บันทึกรายละเอียด", key=f"emeta_save_{ev_name}_{run_count}"):
                                _meta = st.session_state.get("events_meta") or {}
                                _meta[ev_name] = {"date": event_date, "time": event_time, "venue": event_venue}
                                st.session_state["events_meta"] = _meta
                                st.success("✅ บันทึกรายละเอียดแล้ว")
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
                            edited_r = st.data_editor(
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
                                for _, row in edited_r.iterrows() if row["ส่ง"]
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
                                            send_confirmation_email(em, ev_name, name, n, event_date, event_time, event_venue)
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

                # ── เช็คอิน ──────────────────────────────────────────────────────
                with tab_c:
                    if confirmed_df.empty:
                        st.info("ยังไม่มีผู้ผ่านการยืนยัน")
                    else:
                        ci_key        = f"checkin_{ev_name}"
                        last_scan_key = f"last_scan_{ev_name}"
                        scan_msg_key  = f"scan_msg_{ev_name}"

                        if ci_key not in st.session_state:
                            ci_state: dict[str, str] = {}
                            _reg_order: list[tuple]  = []  # (num, name)
                            if out_sheet_id:
                                try:
                                    _gc    = get_gc()
                                    _sheet = _gc.open_by_key(out_sheet_id)
                                    ws_reg = _sheet.worksheet(f"ลงทะเบียน — {ev_name}")
                                    reg_rows = _retry_429(ws_reg.get_all_values)
                                    if len(reg_rows) > 1:
                                        hdr     = reg_rows[0]
                                        n_idx   = next((i for i, h in enumerate(hdr) if h == "ชื่อที่ใช้แข่ง"), 1)
                                        ci_idx  = next((i for i, h in enumerate(hdr) if "เช็คอิน" in h), 2)
                                        num_idx = next((i for i, h in enumerate(hdr) if h == "ลำดับ"), 0)
                                        for r in reg_rows[1:]:
                                            pname = r[n_idx]   if n_idx   < len(r) else ""
                                            ci_v  = r[ci_idx]  if ci_idx  < len(r) else ""
                                            num   = r[num_idx] if num_idx < len(r) else ""
                                            if pname:
                                                ci_state[pname] = ci_v
                                                _reg_order.append((num, pname))
                                except Exception:
                                    pass
                            st.session_state[ci_key]            = ci_state
                            st.session_state[f"{ci_key}_order"] = _reg_order

                        ci_state    = st.session_state[ci_key]
                        _reg_order  = st.session_state.get(f"{ci_key}_order", [])
                        # valid names รวม walk-in ด้วย (ทุกคนใน ลงทะเบียน sheet)
                        valid_names = set(ci_state.keys()) or set(confirmed_df["ชื่อที่ใช้แข่ง"].tolist())

                        if HAS_SCANNER:
                            st.write("**📷 สแกน QR จากอีเมลผู้แข่ง**")
                            scanned = _qr_scanner(key=f"qr_{ev_name}_{run_count}")

                            if scanned and scanned != st.session_state.get(last_scan_key, ""):
                                st.session_state[last_scan_key] = scanned
                                if scanned.startswith("WAKA|"):
                                    parts        = scanned.split("|")
                                    scanned_name = parts[2] if len(parts) >= 3 else ""
                                    if scanned_name in valid_names:
                                        if ci_state.get(scanned_name):
                                            st.session_state[scan_msg_key] = ("warning", f"⚠️ {scanned_name} เช็คอินไปแล้ว")
                                        else:
                                            ci_state[scanned_name] = "✓"
                                            st.session_state[ci_key] = ci_state
                                            if out_sheet_id:
                                                try:
                                                    _sync_checkin_sheet(out_sheet_id, ev_name, confirmed_df, ci_state, name_filter=scanned_name)
                                                except Exception:
                                                    pass
                                            st.session_state[scan_msg_key] = ("success", f"✅ เช็คอิน **{scanned_name}** แล้ว!")
                                    else:
                                        st.session_state[scan_msg_key] = ("error", f"ไม่พบ '{scanned_name}' ในรายการ {ev_name}")
                                else:
                                    st.session_state[scan_msg_key] = ("error", "QR ไม่ถูกต้อง — ใช้ QR จากอีเมลยืนยันเท่านั้น")

                            if scan_msg_key in st.session_state:
                                lvl, msg = st.session_state[scan_msg_key]
                                if lvl == "success":
                                    st.success(msg)
                                elif lvl == "warning":
                                    st.warning(msg)
                                else:
                                    st.error(msg)

                            st.divider()
                        else:
                            st.info("ติดตั้ง `streamlit-qrcode-scanner` เพื่อใช้งาน QR scanner")

                        if _reg_order:
                            all_ci_rows = [
                                {"เช็คอิน ✓": bool(ci_state.get(name, "")), "#": num, "ชื่อที่ใช้แข่ง": name}
                                for num, name in _reg_order
                            ]
                        else:
                            all_ci_rows = [
                                {"เช็คอิน ✓": bool(ci_state.get(row["ชื่อที่ใช้แข่ง"], "")), "#": int(row["#"]), "ชื่อที่ใช้แข่ง": row["ชื่อที่ใช้แข่ง"]}
                                for _, row in confirmed_df.iterrows()
                            ]
                        checked_in_n = sum(1 for r in all_ci_rows if r["เช็คอิน ✓"])
                        st.caption(f"เช็คอินแล้ว **{checked_in_n} / {len(all_ci_rows)}** คน")

                        search = st.text_input(
                            "🔍 ค้นหา", placeholder="พิมพ์ชื่อเพื่อกรอง…",
                            key=f"ci_search_{ev_name}_{run_count}",
                        )
                        ci_df = pd.DataFrame(all_ci_rows)
                        if search:
                            ci_df = ci_df[ci_df["ชื่อที่ใช้แข่ง"].str.contains(search, case=False, na=False)]

                        edited_ci = st.data_editor(
                            ci_df,
                            key=f"ci_ed_{ev_name}_{run_count}_{save_count}",
                            column_config={
                                "เช็คอิน ✓":      st.column_config.CheckboxColumn("เช็คอิน ✓", width="small"),
                                "#":               st.column_config.NumberColumn("#", width="small"),
                                "ชื่อที่ใช้แข่ง": st.column_config.TextColumn("ชื่อที่ใช้แข่ง"),
                            },
                            disabled=["#", "ชื่อที่ใช้แข่ง"],
                            hide_index=True,
                            use_container_width=True,
                        )

                        if st.button("💾 บันทึกเช็คอิน", key=f"ci_save_{ev_name}_{run_count}_{save_count}"):
                            for _, row in edited_ci.iterrows():
                                ci_state[row["ชื่อที่ใช้แข่ง"]] = "✓" if row["เช็คอิน ✓"] else ""
                            st.session_state[ci_key] = ci_state
                            if out_sheet_id:
                                try:
                                    _sync_checkin_sheet(out_sheet_id, ev_name, confirmed_df, ci_state)
                                    st.success("✅ บันทึกเช็คอินแล้ว")
                                except Exception as e:
                                    st.error(f"บันทึกไม่ได้: {e}")
                            else:
                                st.success("✅ บันทึกในหน้านี้แล้ว")
                            st.rerun()
