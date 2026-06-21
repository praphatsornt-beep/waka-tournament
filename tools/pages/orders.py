#!/usr/bin/env python3
"""Card Game Order Dashboard — admin view"""

import json
import re
from pathlib import Path
from datetime import date, timedelta

import streamlit as st
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

try:
    import gspread
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
except ImportError as e:
    st.error(f"ติดตั้ง packages ก่อน: `pip install -r requirements.txt`\n\n{e}")
    st.stop()

SCOPES     = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_PATH = Path("token.json")
SHEET_ID   = "1aUHbSt3qlQ4uMIzlCGbF-iFm0AqSeqx12nxk5ny1JoY"

BRANCHES   = ["ต้นสัก", "เมืองทอง", "ศรีนครินทร์", "จัดส่ง"]
ALL_STATUS = ["รอตรวจ", "รอตรวจเพิ่ม", "ยืนยัน", "ยอดไม่ตรง", "สลิปซ้ำ", "บัญชีไม่ตรง", "สงสัยปลอม", "ยกเลิก", "ไม่มีสลิป"]

# ── Auth ──────────────────────────────────────────────────────────────────────
@st.cache_resource
def get_gc():
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
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("ไม่พบ token.json")
    return gspread.authorize(creds)


@st.cache_data(ttl=60)
def load_orders() -> pd.DataFrame:
    try:
        ws   = get_gc().open_by_key(SHEET_ID).worksheet("orders")
        rows = ws.get_all_values()
        if len(rows) < 2:
            return pd.DataFrame()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df["total"]        = pd.to_numeric(df.get("total", 0), errors="coerce").fillna(0)
        df["slip_amount"]  = pd.to_numeric(df.get("slip_amount", 0), errors="coerce").fillna(0)
        df["timestamp_dt"] = pd.to_datetime(df.get("timestamp", ""), errors="coerce", utc=True)
        df["date"]         = df["timestamp_dt"].dt.tz_convert("Asia/Bangkok").dt.date
        df["row_num"]      = range(2, len(df) + 2)  # แถวจริงใน Sheet (1-indexed + header)
        return df
    except Exception as e:
        st.error(f"โหลด orders ไม่ได้: {e}")
        return pd.DataFrame()


def update_slip_status(row_num: int, status: str, amount: str = "", note: str = ""):
    ws = get_gc().open_by_key(SHEET_ID).worksheet("orders")
    rows = ws.get_all_values()
    hdr  = rows[0]
    status_col = hdr.index("slip_status") + 1 if "slip_status" in hdr else None
    amount_col = hdr.index("slip_amount") + 1 if "slip_amount" in hdr else None
    notes_col  = hdr.index("notes")       + 1 if "notes"       in hdr else None
    if status_col:
        ws.update_cell(row_num, status_col, status)
    if amount_col and amount:
        ws.update_cell(row_num, amount_col, amount)
    if notes_col and note:
        ws.update_cell(row_num, notes_col, note)


def update_fulfillment(row_num: int, status: str):
    from datetime import datetime
    ws = get_gc().open_by_key(SHEET_ID).worksheet("orders")
    hdr = ws.row_values(1)
    ff_col = hdr.index("fulfillment") + 1 if "fulfillment" in hdr else None
    at_col = hdr.index("fulfilled_at") + 1 if "fulfilled_at" in hdr else None
    if ff_col:
        ws.update_cell(row_num, ff_col, status)
    if at_col:
        ws.update_cell(row_num, at_col, datetime.now().strftime("%Y-%m-%d %H:%M"))


def staff_confirm_handover(row_num: int):
    from datetime import datetime
    ws = get_gc().open_by_key(SHEET_ID).worksheet("orders")
    hdr = ws.row_values(1)
    col = hdr.index("staff_confirmed_at") + 1 if "staff_confirmed_at" in hdr else None
    ff_col = hdr.index("fulfillment") + 1 if "fulfillment" in hdr else None
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    if col:
        ws.update_cell(row_num, col, now)
    if ff_col:
        ws.update_cell(row_num, ff_col, "สาขายืนยัน")
    return now


GAS_URL = "https://script.google.com/macros/s/AKfycbz52wvADM7O1zMjqKlT2G4HPkq8gwAon_fUCuKgbmUMkDPQkaYKUWnv598U3EkFN1AByQ/exec"

def send_line_notify(order_id: str, line_user_id: str, msg_type: str):
    import requests
    confirm_url = GAS_URL + "?action=confirm&order=" + order_id
    requests.post(GAS_URL, json={
        "_action": "sendConfirmLink",
        "lineUserId": line_user_id,
        "orderId": order_id,
        "confirmUrl": confirm_url,
        "msgType": msg_type,
    }, timeout=15)


def parse_items(items_json: str) -> list:
    try:
        return json.loads(items_json) if items_json else []
    except Exception:
        return []


def items_text(items: list) -> str:
    lines = []
    for i in items:
        unit = "กล่อง" if i.get("type") == "box" else "ซอง"
        lines.append(f"• {i.get('name','')} ({unit}) ×{i.get('qty',1)} = {i.get('price',0)*i.get('qty',1):,} บาท")
    return "\n".join(lines)


def status_color(s: str) -> str:
    if "ยืนยัน" in s:     return "🟢"
    if "รอตรวจ" in s:     return "🟡"
    if "ยอดไม่ตรง" in s:  return "🟠"
    if "สลิปซ้ำ" in s:    return "🔴"
    if "บัญชีไม่ตรง" in s: return "🔴"
    if "สงสัยปลอม" in s:  return "🔴"
    if "ยกเลิก" in s:     return "🔴"
    return "⚪"

def needs_attention(s: str) -> bool:
    return s in ("รอตรวจ", "ยอดไม่ตรง", "สลิปซ้ำ", "บัญชีไม่ตรง", "สงสัยปลอม")

def fulfill_icon(s: str) -> str:
    if s == "รับแล้ว":            return "✅"
    if s == "สาขายืนยัน":        return "🤝"
    if s == "จัดส่งแล้ว":         return "📦"
    if s == "พร้อมรับ":           return "📍"
    if s == "กำลังจัดส่งไปสาขา":  return "🚚"
    return "⏳"


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Orders", page_icon="🛒", layout="wide")
st.title("🛒 ออเดอร์การ์ดเกม")

# ── Sidebar filters ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("🔍 กรอง")
    if st.button("🔄 โหลดใหม่"):
        st.cache_data.clear()
        st.rerun()

    date_range = st.date_input(
        "ช่วงวันที่",
        value=(date.today() - timedelta(days=7), date.today()),
    )
    branch_filter = st.multiselect("สาขา / จัดส่ง", BRANCHES, default=BRANCHES)
    status_filter = st.multiselect("สถานะสลิป", ALL_STATUS, default=ALL_STATUS)
    search = st.text_input("ค้นหาชื่อ / เบอร์ / เลขออเดอร์", "")

# ── Load ──────────────────────────────────────────────────────────────────────
df = load_orders()
if df.empty:
    st.info("ยังไม่มีออเดอร์")
    st.stop()

# ── Filter ────────────────────────────────────────────────────────────────────
filtered = df.copy()
if len(date_range) == 2:
    filtered = filtered[(filtered["date"] >= date_range[0]) & (filtered["date"] <= date_range[1])]
if branch_filter:
    filtered = filtered[filtered["branch"].isin(branch_filter)]
if status_filter:
    filtered = filtered[filtered["slip_status"].isin(status_filter)]
if search:
    s = search.lower()
    mask = (
        filtered.get("real_name", pd.Series(dtype=str)).str.lower().str.contains(s, na=False) |
        filtered.get("phone",     pd.Series(dtype=str)).str.lower().str.contains(s, na=False) |
        filtered.get("order_id",  pd.Series(dtype=str)).str.lower().str.contains(s, na=False) |
        filtered.get("display_name", pd.Series(dtype=str)).str.lower().str.contains(s, na=False)
    )
    filtered = filtered[mask]

filtered = filtered.sort_values("timestamp_dt", ascending=False).reset_index(drop=True)

# ── KPI ───────────────────────────────────────────────────────────────────────
confirmed = filtered[filtered["slip_status"] == "ยืนยัน"]
pending   = filtered[filtered["slip_status"] == "รอตรวจ"]
problems  = filtered[filtered["slip_status"].isin(["ยอดไม่ตรง", "สลิปซ้ำ", "บัญชีไม่ตรง"])]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("ทั้งหมด", len(filtered))
c2.metric("ยืนยัน 🟢", len(confirmed))
c3.metric("รอตรวจ 🟡", len(pending))
c4.metric("ปัญหา 🔴", len(problems))
c5.metric("ยอดรวม (ยืนยัน)", f"฿{confirmed['total'].sum():,.0f}")

st.divider()

# ── Order cards ───────────────────────────────────────────────────────────────
if filtered.empty:
    st.info("ไม่มีออเดอร์ตามเงื่อนไขที่เลือก")
    st.stop()

for _, row in filtered.iterrows():
    items   = parse_items(row.get("items_json", ""))
    is_del  = row.get("branch", "") == "จัดส่ง"
    s_icon  = status_color(row.get("slip_status", ""))
    label   = f"{s_icon} #{row.get('order_id','')}  |  {row.get('real_name','?')}  |  {row.get('branch','')}  |  ฿{int(row.get('total',0)):,}  |  {row.get('date','')}"

    ff_status = row.get("fulfillment", "") or "รอเตรียม"
    ff_time   = row.get("fulfilled_at", "")
    ff_icon   = fulfill_icon(ff_status)

    with st.expander(label, expanded=needs_attention(row.get("slip_status", ""))):
        # ── บรรทัดที่ 1: ลูกค้า + สาขา
        cur_status = row.get("slip_status", "รอตรวจ")
        info_parts = [
            f"👤 **{row.get('real_name','—')}** ({row.get('display_name','—')})",
            f"📱 {row.get('phone','—')}",
            f"{'🚚 จัดส่ง' if is_del else '📦 ' + row.get('branch','—')}",
            f"{s_icon} {cur_status}",
        ]
        st.markdown(" · ".join(info_parts))
        if is_del and row.get("address"):
            st.caption(f"ที่อยู่: {row.get('address')}")

        # ── บรรทัดที่ 2: สินค้า (inline)
        if items:
            items_str = " | ".join([f"{i.get('name','')} ({'กล่อง' if i.get('type')=='box' else 'ซอง'}) ×{i.get('qty',1)} = ฿{i.get('price',0)*i.get('qty',1):,}" for i in items])
            st.markdown(f"🎴 {items_str} → **฿{int(row.get('total',0)):,}**")

        # ── บรรทัดที่ 3: สลิป + หมายเหตุ
        if row.get("slip_amount", 0):
            st.caption(f"ยอดในสลิป: ฿{int(row.get('slip_amount',0)):,}")
        if row.get("notes"):
            st.caption(f"📝 {row.get('notes')}")

        # ── รูปสลิป + เปลี่ยนสถานะ (2 คอลัมน์เล็ก)
        col_slip, col_act = st.columns([1, 2])
        with col_slip:
            slip_url = row.get("slip_url", "")
            if slip_url and slip_url.startswith("http"):
                st.image(slip_url, width=150)
        with col_act:
            new_status = st.selectbox(
                "เปลี่ยนสถานะ",
                ALL_STATUS,
                index=ALL_STATUS.index(cur_status) if cur_status in ALL_STATUS else 0,
                key=f"status_{row.get('order_id')}",
            )
            new_note = st.text_input("หมายเหตุ", key=f"note_{row.get('order_id')}", placeholder="เช่น โอนไม่ครบ")
            if st.button("💾 บันทึก", key=f"save_{row.get('order_id')}"):
                try:
                    update_slip_status(int(row["row_num"]), new_status, "", new_note)
                    st.success("บันทึกแล้ว")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"บันทึกไม่ได้: {e}")

            st.caption(f"📦 จัดส่ง: {ff_icon} {ff_status}")
