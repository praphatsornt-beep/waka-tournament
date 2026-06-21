#!/usr/bin/env python3
"""ดาวน์โหลดข้อมูลที่อยู่ไทย แล้วสร้าง JSON สำหรับ LIFF dropdown"""

import json
import requests
from pathlib import Path

URL = "https://raw.githubusercontent.com/earthchie/jquery.Thailand.js/master/jquery.Thailand.js/database/raw_database/raw_database.json"
OUT = Path(__file__).parent.parent / "liff" / "thai-address.json"


def main():
    print("กำลังดาวน์โหลดข้อมูลที่อยู่ไทย...")
    res = requests.get(URL, timeout=30)
    res.raise_for_status()
    raw = res.json()
    print(f"ได้ {len(raw)} รายการ")

    tree = {}
    for r in raw:
        prov = r.get("province", "")
        amp  = r.get("amphoe", "")
        dist = r.get("district", "")
        zc   = str(r.get("zipcode", ""))
        if not prov or not amp or not dist:
            continue
        if prov not in tree:
            tree[prov] = {}
        if amp not in tree[prov]:
            tree[prov][amp] = {}
        if dist not in tree[prov][amp]:
            tree[prov][amp][dist] = zc

    provinces = sorted(tree.keys())
    result = []
    for p in provinces:
        amphoes = sorted(tree[p].keys())
        amp_list = []
        for a in amphoes:
            districts = sorted(tree[p][a].keys())
            dist_list = [{"n": d, "z": tree[p][a][d]} for d in districts]
            amp_list.append({"n": a, "d": dist_list})
        result.append({"n": p, "a": amp_list})

    OUT.write_text(json.dumps(result, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    size_kb = OUT.stat().st_size / 1024
    print(f"บันทึกที่ {OUT} ({size_kb:.0f} KB)")
    print(f"จังหวัด: {len(result)}, อำเภอ: {sum(len(p['a']) for p in result)}, ตำบล: {sum(len(a['d']) for p in result for a in p['a'])}")


if __name__ == "__main__":
    main()
