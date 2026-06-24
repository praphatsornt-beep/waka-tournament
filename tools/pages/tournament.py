#!/usr/bin/env python3
"""Tournament Management Dashboard"""

import json
from pathlib import Path
from datetime import date, timedelta, timezone, datetime

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
GAS_URL    = "https://script.google.com/macros/s/AKfycbz52wvADM7O1zMjqKlT2G4HPkq8gwAon_fUCuKgbmUMkDPQkaYKUWnv598U3EkFN1AByQ/exec"

TH_TZ = timezone(timedelta(hours=7))

# ── Auth ─────────────────────────────────────────────────────────────────────
def _build_creds():
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
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                raise RuntimeError("ไม่พบ token.json")
        return creds
    raise RuntimeError("ไม่พบ GOOGLE_TOKEN หรือ token.json")


_gc_client = None

def get_gc():
    global _gc_client
    if _gc_client is not None:
        try:
            _gc_client.open_by_key(SHEET_ID)
            return _gc_client
        except Exception:
            _gc_client = None
    _gc_client = gspread.authorize(_build_creds())
    return _gc_client


def _now_th():
    return datetime.now(TH_TZ).strftime("%Y-%m-%d")


# ── Data loading ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=30)
def load_registrations() -> pd.DataFrame:
    try:
        ws = get_gc().open_by_key(SHEET_ID).worksheet("tournament_reg")
        rows = ws.get_all_values()
        if len(rows) < 2:
            return pd.DataFrame()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df["row_num"] = range(2, len(df) + 2)
        return df
    except gspread.exceptions.WorksheetNotFound:
        return pd.DataFrame()
    except Exception as e:
        st.error(f"โหลด tournament_reg ไม่ได้: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30)
def load_player_stats() -> pd.DataFrame:
    try:
        ws = get_gc().open_by_key(SHEET_ID).worksheet("player_stats")
        rows = ws.get_all_values()
        if len(rows) < 2:
            return pd.DataFrame()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        for col in ["total_plays", "accumulation_count", "cards_received", "boxes_earned", "boxes_given"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        df["row_num"] = range(2, len(df) + 2)
        return df
    except gspread.exceptions.WorksheetNotFound:
        return pd.DataFrame()
    except Exception as e:
        st.error(f"โหลด player_stats ไม่ได้: {e}")
        return pd.DataFrame()


def update_reg_field(reg_id: str, field: str, value: str):
    import requests
    requests.get(
        GAS_URL,
        params={"action": "api", "do": "tournament_update_reg", "reg_id": reg_id, "field": field, "value": value},
        timeout=15,
    )


def give_box(line_user_id: str):
    import requests
    requests.get(
        GAS_URL,
        params={"action": "api", "do": "tournament_give_box", "line_user_id": line_user_id},
        timeout=15,
    )


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Tournament", page_icon="🏆", layout="wide")

t1, t2 = st.columns([3, 1])
with t1:
    st.markdown("### 🏆 Tournament")
with t2:
    if st.button("🔄 โหลดใหม่", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

tab_today, tab_stats, tab_slips = st.tabs(["📋 วันนี้", "📦 สะสม", "🧾 ตรวจสลิป"])

# ── Tab 1: Today ─────────────────────────────────────────────────────────────
with tab_today:
    df_reg = load_registrations()
    today_str = _now_th()

    with st.sidebar:
        st.header("กรอง")
        selected_date = st.date_input("วันที่", value=date.today())
        today_str = selected_date.strftime("%Y-%m-%d")

    if df_reg.empty:
        st.info("ยังไม่มีข้อมูลลงทะเบียน")
    else:
        today_df = df_reg[df_reg["event_date"] == today_str].copy()

        if today_df.empty:
            st.info(f"ไม่มีผู้ลงทะเบียนวันที่ {today_str}")
        else:
            verified = today_df[today_df["slip_status"] == "verified"]
            pending = today_df[today_df["slip_status"] == "pending"]
            cards_count = today_df[today_df["choice"] == "cards"]
            accum_count = today_df[today_df["choice"] == "accumulate"]

            kpi = (
                f"📋 **{len(today_df)}** ลงทะเบียน &nbsp;|&nbsp; "
                f"✅ **{len(verified)}** ตรวจแล้ว &nbsp;|&nbsp; "
                f"🟡 **{len(pending)}** รอตรวจ &nbsp;|&nbsp; "
                f"🎴 **{len(cards_count)}** รับการ์ด &nbsp;|&nbsp; "
                f"📦 **{len(accum_count)}** สะสม"
            )
            st.markdown(kpi)

            names_list = []
            for _, r in today_df.iterrows():
                name = r.get("real_name") or r.get("display_name") or "—"
                names_list.append(name)

            if st.button("📋 คัดลอกรายชื่อ (สำหรับจับคู่แข่ง)"):
                names_text = "\n".join(f"{i+1}. {n}" for i, n in enumerate(names_list))
                st.code(names_text, language=None)

            st.markdown("---")

            for _, row in today_df.iterrows():
                rid = row.get("reg_id", "")
                name = row.get("real_name") or row.get("display_name") or "—"
                slip_st = row.get("slip_status", "pending")
                choice = row.get("choice", "cards")
                cards_given = row.get("cards_given", "")

                s_icon = "✅" if slip_st == "verified" else ("❌" if slip_st == "rejected" else "🟡")
                c_icon = "📦 สะสม" if choice == "accumulate" else "🎴 รับการ์ด"
                given_icon = " ✅แจกแล้ว" if cards_given == "TRUE" else ""

                label = f"{s_icon} **#{rid}** · {name} · {c_icon}{given_icon}"

                with st.expander(label, expanded=(slip_st == "pending")):
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        slip_url = row.get("slip_url", "")
                        if slip_url and slip_url.startswith("http"):
                            st.image(slip_url, width=150)
                        else:
                            st.caption("ไม่มีสลิป")
                    with col2:
                        st.markdown(
                            f"👤 **{name}** ({row.get('display_name', '—')})\n\n"
                            f"📱 {row.get('phone', '—')} · {c_icon}"
                        )
                        if row.get("note"):
                            st.caption(f"📝 {row.get('note')}")

                        c1, c2, c3 = st.columns(3)
                        with c1:
                            if slip_st != "verified":
                                if st.button("✅ ยืนยันสลิป", key=f"verify_{rid}"):
                                    update_reg_field(rid, "slip_status", "verified")
                                    st.success("ยืนยันแล้ว")
                                    st.cache_data.clear()
                                    st.rerun()
                        with c2:
                            if slip_st != "rejected":
                                if st.button("❌ ปฏิเสธ", key=f"reject_{rid}"):
                                    update_reg_field(rid, "slip_status", "rejected")
                                    st.warning("ปฏิเสธแล้ว")
                                    st.cache_data.clear()
                                    st.rerun()
                        with c3:
                            if choice == "cards" and cards_given != "TRUE":
                                if st.button("🎴 แจกการ์ดแล้ว", key=f"give_{rid}"):
                                    update_reg_field(rid, "cards_given", "TRUE")
                                    st.success("บันทึกแล้ว")
                                    st.cache_data.clear()
                                    st.rerun()

# ── Tab 2: Player Stats ─────────────────────────────────────────────────────
with tab_stats:
    df_stats = load_player_stats()

    if df_stats.empty:
        st.info("ยังไม่มีข้อมูลผู้เล่น")
    else:
        search = st.text_input("🔍 ค้นหาผู้เล่น", placeholder="ชื่อ / display name", key="stats_search")
        if search:
            s = search.lower()
            mask = (
                df_stats.get("display_name", pd.Series(dtype=str)).str.lower().str.contains(s, na=False) |
                df_stats.get("real_name", pd.Series(dtype=str)).str.lower().str.contains(s, na=False)
            )
            df_stats = df_stats[mask]

        for _, row in df_stats.iterrows():
            name = row.get("real_name") or row.get("display_name") or "—"
            plays = int(row.get("total_plays", 0))
            acc = int(row.get("accumulation_count", 0))
            cards = int(row.get("cards_received", 0))
            boxes_e = int(row.get("boxes_earned", 0))
            boxes_g = int(row.get("boxes_given", 0))
            uid = row.get("line_user_id", "")

            pending_box = boxes_e - boxes_g
            badge = f" 🎁 **Box ค้าง {pending_box}**" if pending_box > 0 else ""

            with st.expander(f"👤 **{name}** · เล่น {plays} ครั้ง · สะสม {acc}/10{badge}"):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("เข้าร่วม", plays)
                c2.metric("สะสม", f"{acc}/10")
                c3.metric("รับการ์ด", f"{cards} ครั้ง")
                c4.metric("Box ได้/แจก", f"{boxes_e}/{boxes_g}")

                pct = min(acc * 10, 100)
                st.progress(pct / 100, text=f"สะสม {acc}/10")

                if pending_box > 0:
                    if st.button(f"🎁 ให้ Box แล้ว ({pending_box} ค้าง)", key=f"box_{uid}"):
                        give_box(uid)
                        st.success("บันทึกแล้ว")
                        st.cache_data.clear()
                        st.rerun()

# ── Tab 3: Slip Verification ─────────────────────────────────────────────────
with tab_slips:
    df_reg2 = load_registrations()

    if df_reg2.empty:
        st.info("ยังไม่มีข้อมูล")
    else:
        pending_df = df_reg2[df_reg2["slip_status"] == "pending"].copy()
        pending_df = pending_df.iloc[::-1].reset_index(drop=True)

        if pending_df.empty:
            st.success("ไม่มีสลิปรอตรวจ")
        else:
            st.markdown(f"**{len(pending_df)}** สลิปรอตรวจ")
            for _, row in pending_df.iterrows():
                rid = row.get("reg_id", "")
                name = row.get("real_name") or row.get("display_name") or "—"
                ev_date = row.get("event_date", "")

                with st.expander(f"🟡 #{rid} · {name} · {ev_date}"):
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        slip_url = row.get("slip_url", "")
                        if slip_url and slip_url.startswith("http"):
                            st.image(slip_url, width=200)
                        else:
                            st.caption("ไม่มีรูปสลิป")
                    with col2:
                        st.markdown(f"👤 **{name}**\n\n📅 {ev_date}")
                        c1, c2 = st.columns(2)
                        with c1:
                            if st.button("✅ ยืนยัน", key=f"sv_{rid}"):
                                update_reg_field(rid, "slip_status", "verified")
                                st.success("ยืนยันแล้ว")
                                st.cache_data.clear()
                                st.rerun()
                        with c2:
                            if st.button("❌ ปฏิเสธ", key=f"sr_{rid}"):
                                update_reg_field(rid, "slip_status", "rejected")
                                st.warning("ปฏิเสธแล้ว")
                                st.cache_data.clear()
                                st.rerun()
