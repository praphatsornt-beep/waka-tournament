# Verify Tournament Registrations

## Objective
ตรวจสอบสลิปโอนเงินและจับคู่กับรายชื่อผู้สมัครจาก Google Form แล้วสร้างรายชื่อยืนยันใน Google Sheet

## Inputs Required
- `events_config.json` — Sheet IDs และค่าสมัครแต่ละรายการ
- `credentials.json` — Google OAuth (ดู `workflows/setup_google_auth.md`)
- `.env` — มี `ANTHROPIC_API_KEY`

## Output
Google Sheet (output_sheet_id ใน config) โดยมี 1 tab ต่อ 1 รายการแข่ง

แต่ละแถวมี:
| คอลัมน์ | ความหมาย |
|---------|----------|
| ชื่อที่ใช้แข่ง | game name ในวงการ |
| ชื่อใน OpenChat | ชื่อ Line OpenChat |
| ชื่อเฟสบุค | Facebook name |
| ชื่อที่โอนเงิน | ตามที่กรอกในฟอร์ม |
| สถานะ | ✅ ⚠️ ❌ 🔍 |
| รายละเอียด | AI อ่านจากสลิป |
| วันที่สลิป | วันที่โอน |
| ธนาคาร | ธนาคารที่โอน |

## Status Legend
| สถานะ | ความหมาย |
|-------|---------|
| ✅ ยืนยันแล้ว | ชื่อ + ยอดตรง — ยอมรับได้เลย |
| ⚠️ ยอดไม่ตรง | ชื่อตรงแต่ยอดต่าง — แอดมินตรวจ |
| ⚠️ ชื่อไม่ตรง | ยอดตรงแต่ชื่อต่าง — แอดมินตรวจ |
| ❌ ไม่ตรงทั้งคู่ | ทั้งชื่อและยอดไม่ตรง — ติดต่อลูกค้า |
| 🔍 อ่านสลิปไม่ได้ | รูปไม่ชัดหรือไม่ใช่สลิป — ขอสลิปใหม่ |
| 🚫 สลิปซ้ำ | สลิปนี้ถูกใช้ไปแล้ว — ติดต่อลูกค้าขอสลิปใหม่ |

---

## ขั้นตอนการใช้งาน

### ครั้งแรก — Setup
1. ติดตั้ง dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. ตั้งค่า Google credentials ตาม `workflows/setup_google_auth.md`
3. Copy `.env.example` → `.env` แล้วใส่ `ANTHROPIC_API_KEY`
4. Copy `events_config.example.json` → `events_config.json` แล้วกรอก:
   - `output_sheet_id` — Sheet ที่จะเขียนผล (เปิด Sheet → เอา ID จาก URL)
   - แต่ละ event: `source_sheet_id` (Form Responses Sheet) และ `fee`

### วิธีหา Sheet ID
จาก URL: `https://docs.google.com/spreadsheets/d/**SHEET_ID**/edit`

### รันทุก event
```bash
uv run tools/process_registrations.py
```

### รัน event เดียว
```bash
uv run tools/process_registrations.py --event "Standard"
```

---

## Edge Cases

| ปัญหา | สาเหตุ | วิธีแก้ |
|-------|--------|--------|
| "Cannot download file" | สลิปใน Drive ไม่ได้ share | ตรวจสอบสิทธิ์ Drive ของ Form |
| "Could not detect columns" | ชื่อหัวคอลัมน์ไม่ตรง keyword | แก้ `COLUMN_KEYWORDS` ใน script |
| "readable: false" | รูปเบลอ/ไม่ใช่สลิป | ติดต่อลูกค้าขอสลิปใหม่ |
| ชื่อใน สลิปเป็นภาษาอังกฤษ | ชื่อบัญชีอาจเป็น EN | ระบบ normalize แล้ว แต่ตรวจสอบ ⚠️ |

## Notes
- สลิปที่ส่งใน Form จะเก็บใน Google Drive ของเจ้าของ Form อัตโนมัติ
- Script ใช้ Claude Vision อ่านสลิป — แต่ละรูปใช้ ~1-2 วินาที
- รัน script ซ้ำได้ — จะ clear และเขียน Sheet ใหม่ทุกครั้ง

---

## วิธีใช้งานแบบฟรี — จับคู่กับรายงานธนาคาร

แทนที่จะให้ AI อ่านรูปสลิป สามารถ export รายงานธนาคารเป็น CSV แล้วจับคู่โดยตรงได้เลย

### เงื่อนไข
- Form ต้องมีช่องให้ลูกค้ากรอก **ชื่อบัญชีที่โอน** (ชื่อเจ้าของบัญชีธนาคาร)
- Export รายงานธนาคารเป็น CSV จากแอปหรือ internet banking

### รัน
```bash
uv run tools/match_bank_csv.py --csv path/to/bank_statement.csv
uv run tools/match_bank_csv.py --csv statement.csv --event "Standard"
```

### Status Legend เพิ่มเติม
| สถานะ | ความหมาย |
|-------|---------|
| ✅ ยืนยันแล้ว | ชื่อบัญชี + ยอดตรงกับรายงานธนาคาร |
| ⚠️ ตรวจสอบชื่อ | ยอดตรงแต่ชื่อต่าง — แอดมินตรวจ |
| ❌ ไม่พบ | ไม่มี transaction ที่ตรงในรายงาน — ติดต่อลูกค้า |
| 🚫 transaction ซ้ำ | เลขที่รายการนี้ถูกใช้ไปแล้ว |

### ข้อดี
- ฟรี 100% ไม่ใช้ AI
- แม่นกว่า เพราะข้อมูลจากธนาคารโดยตรง
- แต่ละ transaction ใช้ได้ครั้งเดียว (ป้องกันการแอบใช้ซ้ำ)
