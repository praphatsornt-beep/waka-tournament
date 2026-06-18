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

# ── Constants ─────────────────────────────────────────────────────────────────
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_PATH = Path("token.json")
CREDS_PATH = Path("credentials.json")

BRANCHES   = ["ต้นสัก", "เมืองทอง", "ศรีนครินทร์", "จัดส่งพัสดุ"]
STATUS_OK  = "✅ ยืนยันแล้ว"
COL_SLIP   = "slip_status"

# ── Auth (ใช้ get_gc เดียวกับ verify_app) ────────────────────────────────────
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
            raise RuntimeError("ไม่พบ token.json หรือ GOOGLE_TOKEN ใน Secrets")
    return gspread.authorize(creds)


def load_orders(sheet_id: str) -> pd.DataFrame:
    try:
        ws   = get_gc().open_by_key(sheet_id).worksheet("orders")
        rows = ws.get_all_values()
        if len(rows) < 2:
            return pd.DataFrame()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0)
        df["slip_amount"] = pd.to_numeric(df["slip_amount"], errors="coerce").fillna(0)
        df["timestamp_dt"] = pd.to_datetime(df["timestamp"], errors="coerce")
        df["date"] = df["timestamp_dt"].dt.date
        return df
    except Exception as e:
        st.error(f"โหลด orders ไม่ได้: {e}")
        return pd.DataFrame()


def parse_items(items_json: str) -> list:
    try:
        return json.loads(items_json) if items_json else []
    except Exception:
        return []


# ── Page ──────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Orders Dashboard", page_icon="🛒", layout="wide")
st.title("🛒 ออเดอร์การ์ดเกม")

# ── Sheet ID input ────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ ตั้งค่า")
    sheet_url = st.text_input(
        "Orders Sheet URL",
        value=st.session_state.get("order_sheet_url", ""),
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
    if sheet_url:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", sheet_url)
        if m:
            st.session_state["order_sheet_url"] = sheet_url
            sheet_id = m.group(1)
        else:
            st.error("URL ไม่ถูกต้อง")
            st.stop()
    else:
        st.info("ใส่ URL ของ Google Sheet ที่เก็บ orders")
        st.stop()

    if st.button("🔄 โหลดใหม่"):
        st.cache_data.clear()
        st.rerun()

# ── Load ──────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=120)
def _load(sid):
    return load_orders(sid)

df = _load(sheet_id)
if df.empty:
    st.info("ยังไม่มีออเดอร์")
    st.stop()

# ── Filters ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.subheader("🔍 กรอง")
    date_range = st.date_input(
        "ช่วงวันที่",
        value=(date.today() - timedelta(days=7), date.today()),
    )
    branch_filter = st.multiselect("สาขา", BRANCHES, default=BRANCHES)
    status_filter = st.multiselect(
        "สถานะสลิป",
        df[COL_SLIP].unique().tolist(),
        default=df[COL_SLIP].unique().tolist(),
    )

# Apply filters
filtered = df.copy()
if len(date_range) == 2:
    filtered = filtered[
        (filtered["date"] >= date_range[0]) &
        (filtered["date"] <= date_range[1])
    ]
for branch in BRANCHES:
    pass
filtered = filtered[filtered["pickup"].str.contains("|".join(branch_filter), na=False, regex=False) if branch_filter else True]
filtered = filtered[filtered[COL_SLIP].isin(status_filter)] if status_filter else filtered

# ── KPI ───────────────────────────────────────────────────────────────────────
confirmed = filtered[filtered[COL_SLIP].str.contains("✅", na=False, regex=False)]
pending   = filtered[~filtered[COL_SLIP].str.contains("✅", na=False, regex=False)]

c1, c2, c3, c4 = st.columns(4)
c1.metric("ออเดอร์ทั้งหมด", len(filtered))
c2.metric("ยืนยันแล้ว ✅", len(confirmed))
c3.metric("รอตรวจ ⏳", len(pending))
c4.metric("ยอดรวม (ยืนยันแล้ว)", f"฿{confirmed['total'].sum():,.0f}")

st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_orders, tab_branch, tab_report = st.tabs(["📋 รายการออเดอร์", "🏪 สรุปต่อสาขา", "📊 รายงาน"])

# ── Tab: รายการออเดอร์ ────────────────────────────────────────────────────────
with tab_orders:
    show_cols = ["order_id", "date", "display_name", "nickname", "real_name",
                 "pickup", "total", COL_SLIP, "slip_amount", "notes"]
    show_cols = [c for c in show_cols if c in filtered.columns]
    display_df = filtered[show_cols].sort_values("date", ascending=False).reset_index(drop=True)
    st.dataframe(
        display_df,
        column_config={
            "order_id":    st.column_config.TextColumn("เลขออเดอร์", width="medium"),
            "date":        st.column_config.DateColumn("วันที่"),
            "display_name":st.column_config.TextColumn("Line Name"),
            "nickname":    st.column_config.TextColumn("ชื่อเล่น"),
            "real_name":   st.column_config.TextColumn("ชื่อจริง"),
            "pickup":      st.column_config.TextColumn("สาขา/จัดส่ง"),
            "total":       st.column_config.NumberColumn("ยอด (฿)", format="฿%.0f"),
            COL_SLIP:      st.column_config.TextColumn("สถานะสลิป"),
            "slip_amount": st.column_config.NumberColumn("ยอดในสลิป (฿)", format="฿%.0f"),
            "notes":       st.column_config.TextColumn("หมายเหตุ"),
        },
        hide_index=True,
        use_container_width=True,
    )

    if st.button("⬇️ Export CSV"):
        csv = display_df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("ดาวน์โหลด", csv, "orders.csv", "text/csv")

# ── Tab: สรุปต่อสาขา ─────────────────────────────────────────────────────────
with tab_branch:
    for branch in BRANCHES:
        bdf = filtered[filtered["pickup"].str.contains(branch, na=False, regex=False)]
        if bdf.empty:
            continue
        b_confirmed = bdf[bdf[COL_SLIP].str.contains("✅", na=False, regex=False)]
        with st.expander(f"🏪 {branch} — {len(bdf)} ออเดอร์ (ยืนยัน {len(b_confirmed)})", expanded=True):
            cols = ["order_id", "date", "display_name", "nickname", "real_name", "total", COL_SLIP]
            cols = [c for c in cols if c in bdf.columns]
            st.dataframe(
                bdf[cols].sort_values("date", ascending=False).reset_index(drop=True),
                hide_index=True,
                use_container_width=True,
            )
            st.caption(f"ยอดรวมยืนยัน: ฿{b_confirmed['total'].sum():,.0f}")

# ── Tab: รายงาน ───────────────────────────────────────────────────────────────
with tab_report:
    st.subheader("ยอดขายต่อวัน")
    daily = (
        confirmed.groupby("date")["total"].sum()
        .reset_index()
        .rename(columns={"total": "ยอด (฿)"})
        .sort_values("date")
    )
    if not daily.empty:
        st.bar_chart(daily.set_index("date"))

    st.subheader("ยอดขายต่อสาขา")
    branch_sum = (
        confirmed.groupby("pickup")["total"].agg(["sum", "count"])
        .reset_index()
        .rename(columns={"pickup": "สาขา", "sum": "ยอดรวม (฿)", "count": "จำนวนออเดอร์"})
        .sort_values("ยอดรวม (฿)", ascending=False)
    )
    st.dataframe(branch_sum, hide_index=True, use_container_width=True)

    st.subheader("สินค้าขายดี")
    item_rows = []
    for _, row in confirmed.iterrows():
        for item in parse_items(row.get("items_json", "")):
            item_rows.append({
                "สินค้า": item.get("name", ""),
                "จำนวน": item.get("qty", 0),
                "ยอด (฿)": item.get("qty", 0) * item.get("price", 0),
            })
    if item_rows:
        item_df = (
            pd.DataFrame(item_rows)
            .groupby("สินค้า")[["จำนวน", "ยอด (฿)"]].sum()
            .reset_index()
            .sort_values("ยอด (฿)", ascending=False)
        )
        st.dataframe(item_df, hide_index=True, use_container_width=True)
