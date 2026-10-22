#!/usr/bin/env python3
"""决算端预算查询：按支队关键词从预算表读出决算审核的封顶基准。

用法：
    python3 budget_lookup.py 预算表.xlsx 关键词 [--sheet 表名] [--start-row 7]

关键词匹配 B 列实践地点（可给多个词，如 "合肥 上海" 命中"合肥、上海"行）。
输出该行决算比对所需的全部上限：交通逐段依据(AD)、单人交通费(T)、住宿口径、保险上限、总额(R)。
「审核报销」skill §七 预决算比对调用本脚本取数。
"""
import argparse, re, sys
from datetime import date

import openpyxl


def parse_days(text, today=None):
    m = re.search(r"(\d+)\.(\d+)\s*-\s*(?:(\d+)\.)?(\d+)", str(text or ""))
    if not m:
        return None
    today = today or date.today()
    m1, d1 = int(m.group(1)), int(m.group(2))
    m2, d2 = int(m.group(3) or m.group(1)), int(m.group(4))
    for yr in (today.year, today.year + 1):
        try:
            s, e = date(yr, m1, d1), date(yr, m2, d2)
        except ValueError:
            continue
        if e >= today:
            return (e - s).days + 1
    return None


def eval_t(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").lstrip("=")
    if re.fullmatch(r"[0-9+\-.\s]+", s):
        return float(eval(s))
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("keywords", nargs="+", help="实践地关键词（全部命中才匹配）")
    ap.add_argument("--sheet", default="各支队预算按实际人数校准")
    ap.add_argument("--start-row", type=int, default=7)
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.input, data_only=True)
    wsf = openpyxl.load_workbook(args.input)  # 公式版兜底（T/AD 等）
    ws = wb[args.sheet] if args.sheet in wb.sheetnames else wb.active
    wsf = wsf[ws.title]

    hits = []
    r = args.start_row
    while ws[f"A{r}"].value is not None or ws[f"B{r}"].value is not None:
        b = re.sub(r"\s+", "", str(ws[f"B{r}"].value or ""))
        if all(k in b for k in args.keywords):
            hits.append(r)
        r += 1
    if not hits:
        sys.exit(f"未找到匹配 {'+'.join(args.keywords)} 的支队行")

    for r in hits:
        t_formula = wsf[f"T{r}"].value
        t_val = ws[f"T{r}"].value if isinstance(ws[f"T{r}"].value, (int, float)) else eval_t(t_formula)
        days = ws[f"J{r}"].value if isinstance(ws[f"J{r}"].value, (int, float)) \
            else parse_days(wsf[f"E{r}"].value)
        people = ws[f"K{r}"].value if isinstance(ws[f"K{r}"].value, (int, float)) \
            else wsf[f"F{r}"].value
        v_price = wsf[f"V{r}"].value or 300
        z_price = wsf[f"Z{r}"].value or 30
        ac = ws[f"AC{r}"].value
        if not isinstance(ac, (int, float)):  # 无缓存时按 D 列口径推
            d_txt = str(wsf[f"D{r}"].value or "")
            ac = 1 if ("就协" in d_txt or "就业协会" in d_txt) else 2 / 3
        ac_txt = f"{ac:.4g}"
        s_txt = ws[f"S{r}"].value or ("就协全部" if ac == 1 else "院系1/3，剩余就协")
        print(f"══ 行{r}  支队: {ws[f'B{r}'].value} ══")
        print(f"  参访日期: {wsf[f'E{r}'].value}   天数: {days}（{days - 1 if days else '?'}晚）   "
              f"人数: {people}   报销比例: {ac_txt}")
        print(f"  交通费上限: 单人 {t_val} 元（T列: {t_formula}）")
        print(f"    取价依据: {wsf[f'AD{r}'].value or '（无）'}")
        print(f"  住宿上限: {v_price / 2:g} 元/人·晚 × {days - 1 if days else '?'} 晚"
              f"{'（实践地承担，V=0）' if v_price == 0 else ''}")
        weeks = -(-(days or 1) // 7)
        print(f"  保险上限: {z_price} 元/人/周 × {weeks}周")
        r_total = ws[f"R{r}"].value
        if r_total is None and all(isinstance(x, (int, float)) for x in (t_val, people)) and days:
            import math
            r_total = (t_val * people * ac
                       + (v_price * math.ceil(people / 2) * (days - 1) * ac if v_price else 0)
                       + z_price * people * weeks * ac)
            r_total = f"{r_total:g}（按口径重算）"
        print(f"  预算总额: {r_total if r_total is not None else '(R列未缓存，需在Excel打开保存)'}")
        print(f"  备注: {s_txt}")
        print()


if __name__ == "__main__":
    main()
