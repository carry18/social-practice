#!/usr/bin/env python3
"""企业台账 companies.csv → enterprise_index.json（面板用，无联系人/职务）。

只保留具名单位；去掉“部门/招聘服务”这类非公司记录。
"""
import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT.parent / "企业台账" / "companies.csv"
ALIASES = ROOT.parent / "企业台账" / "enterprise_aliases.csv"
OUT = ROOT / "enterprise_index.json"

DEFENSE = {
    "中核", "航天科技", "航天科工", "航空工业", "中国航发", "中国船舶",
    "兵器工业", "兵器装备", "中国电科", "中国电子", "中国工程物理研究院",
}
TROOPS = {"部队"}

def formalize_name(s):
    """公文口径：只去掉间隔号，编号保留阿拉伯数字（14所、719所）。"""
    return str(s or "").replace("·", "")


def main():
    rows = list(csv.DictReader(open(SRC, encoding="utf-8-sig", newline="")))
    alias_rows = list(csv.DictReader(open(ALIASES, encoding="utf-8-sig", newline="")))
    alias_map = {(a["单位名称"], a["集团"]): a for a in alias_rows}
    companies = []
    for r in rows:
        if r["是否具名单位"] != "是":
            continue
        if r["层级"] in ("部门", "服务商"):
            continue
        group = r["集团"]
        a = alias_map.get((r["单位名称"], group), {})
        legal = a.get("法定全称", "")
        canonical_raw = a.get("规范名", "") or r["展示名称"]
        canonical = formalize_name(canonical_raw)
        short = formalize_name(a.get("简称", "") or canonical_raw)
        aliases = []
        for x in (a.get("别名", "") or "").split("|"):
            for v in (x, formalize_name(x)):
                if v and v not in aliases:
                    aliases.append(v)
        name_status = a.get("状态", "")
        write_name = legal or canonical
        companies.append({
            "id": "c_%04d" % (len(companies) + 1),
            "name": canonical,
            "canonical_name": canonical,
            "short_name": short,
            "legal_name": legal,
            "aliases": aliases,
            "name_status": name_status,
            "raw": r["单位名称"],
            "write_name": write_name,
            "group": group,
            "level": r["层级"],
            "province": r["省份"],
            "city": r["城市"],
            "source": r["地点来源"],
            "defense": group in DEFENSE,
            "troops": group in TROOPS,
        })

    group_count = Counter(c["group"] for c in companies)
    city_count = Counter(c["city"] for c in companies if c["city"])
    province_of_city = {}
    for c in companies:
        if c["city"] and c["province"]:
            province_of_city.setdefault(c["city"], c["province"])
    for c in companies:
        c["group_count"] = group_count[c["group"]]

    cities = {city: {"province": province_of_city.get(city, ""), "count": n}
              for city, n in city_count.most_common()}
    groups = {g: {"count": n, "defense": g in DEFENSE, "troops": g in TROOPS}
              for g, n in group_count.most_common()}

    out = {
        "meta": {
            "version": "2026-11-07",
            "source": "联系人导出 + 企业台账（无联系人/职务）",
            "count": len(companies),
            "city_count": len(cities),
            "group_count": len(groups),
        },
        "cities": cities,
        "groups": groups,
        "companies": companies,
    }
    OUT.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"companies={len(companies)} cities={len(cities)} groups={len(groups)} bytes={OUT.stat().st_size}")


if __name__ == "__main__":
    main()
