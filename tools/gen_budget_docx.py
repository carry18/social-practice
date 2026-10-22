#!/usr/bin/env python3
"""由预算 xlsx（rebuild_budget.py 重建或人工调整后）生成《实践预算》Word 公文。

只生成"企业实习实践支队"部分（我们就协自己的部分）；基层实习实践支队（基研会部分）不生成，
统筹文字中相应位置留【】占位。

用法：
    python3 gen_budget_docx.py 预算表.xlsx [-o 输出.docx] \
        [--title '2026年暑假“清华学子地域行”实践预算'] [--season 2026年暑假]

调整入口：在 Excel 里改完预算表（人数/票价/住宿单价/比例等均可改）并保存后，
直接重跑本脚本即可同步生成新 Word——呈现列优先读 Excel 公式缓存值（以 Excel 实际数字为准）；
未在 Excel 打开保存过（无缓存）时按既定口径从 B/D/E/F/T/V/Z 输入格重算
（交通=单人票价×人数×比例；住宿=V标间×ceil(人数/2)×(天数-1)×比例；保险=Z×人数×ceil(天数/7)×比例，
 即 ≤30元/人/周；比例：D 含"就协"/"就业协会" → 1，否则 2/3）。
"""
import argparse, math, re
from datetime import date
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import openpyxl

def parse_days(text, year=2026):
    m = re.search(r"(\d+)\.(\d+)\s*-\s*(?:(\d+)\.)?(\d+)", str(text or ""))
    if not m:
        return None
    m1, d1 = int(m.group(1)), int(m.group(2))
    m2, d2 = int(m.group(3) or m.group(1)), int(m.group(4))
    return (date(year, m2, d2) - date(year, m1, d1)).days + 1

def parse_fare(v):
    """T 列：数字或 '=672+667' 式公式 → 数值。"""
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and v.startswith("="):
        return sum(float(x) for x in re.findall(r"[\d.]+", v))
    return None

def ratio_of(d):
    """D 列含"就协"或"就业协会"（联合承办也算）→ 1，否则 2/3。"""
    d = re.sub(r"\s+", "", str(d or ""))
    return 1 if ("就协" in d or "就业协会" in d) else 2 / 3

def ratio_str(r):
    return "x1" if r == 1 else "x 2/3"

def money(x):
    return str(int(round(x)))

def read_teams(path, sheet, start_row):
    """读取支队行。呈现列(L/M/N/O/P/Q/R/S)与辅助值(K/AC/V/Z)优先取 Excel 公式缓存值
    （用户在 Excel 里调整保存后，以 Excel 实际数字为准）；无缓存时按既定口径重算。
    保险统一口径：≤30 元/人/周（就协规则库），即 Z × 人数 × ceil(天数/7) × 比例。"""
    wb = openpyxl.load_workbook(path)
    wbv = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet in wb.sheetnames else wb.active
    wsv = wbv[ws.title]
    teams, r = [], start_row
    while ws[f"A{r}"].value is not None or ws[f"B{r}"].value is not None:
        B, D, E, F = (ws[f"{c}{r}"].value for c in "BDEF")
        cv = {c: wsv[f"{c}{r}"].value for c in
              ("K", "L", "M", "N", "O", "P", "Q", "R", "S", "AC")}
        T = parse_fare(ws[f"T{r}"].value)
        days = parse_days(E)
        people = int(cv["K"]) if isinstance(cv["K"], (int, float)) else int(F or 12)
        ac_cell = ws[f"AC{r}"].value  # AC 一般为公式；手写调整时可能是字面量
        ratio = (float(cv["AC"]) if isinstance(cv["AC"], (int, float))
                 else float(ac_cell) if isinstance(ac_cell, (int, float))
                 else ratio_of(D))
        # V/Z 优先读单元格（rebuild 已写入；实践地承担住宿时 V=0）
        v_cell = ws[f"V{r}"].value
        z_cell = ws[f"Z{r}"].value
        nights = (days - 1) if days else 0
        if isinstance(v_cell, (int, float)):
            per_lodge = v_cell / 2          # V 是标间价，折半为人·晚
        else:
            lodging_free = "承担住宿" in "".join(
                str(ws[f"{c}{r}"].value or "") for c in "ABCDEFGH")
            per_lodge = 0 if lodging_free else 150
        ins = z_cell if isinstance(z_cell, (int, float)) else 30

        t_sum = (T or 0) * people * ratio
        l_sum = per_lodge * 2 * math.ceil(people / 2) * nights * ratio if per_lodge else 0
        weeks = math.ceil(days / 7) if days else 1
        i_sum = ins * people * weeks * ratio

        def txt(key, fallback):
            v = cv[key]
            return v.strip() if isinstance(v, str) and v.strip() else fallback

        def num(key, calc):
            v = cv[key]
            return float(v) if isinstance(v, (int, float)) else calc

        ratio_s = "" if ratio == 1 else ratio_str(ratio)
        t_desc = txt("L", f"{money(T or 0)}/人{ratio_s}" if T is not None else "待查")
        l_desc = txt("N", ("实践地承担" if per_lodge == 0 else
                           (f"{money(per_lodge)}/人·天*{nights}晚" if ratio == 1 else
                            f"{money(per_lodge)}/人·天*2/3*{nights}晚")))
        i_desc = txt("P", (f"{money(ins)}/人·周*{weeks}周{ratio_s}" if weeks > 1 else
                           f"{money(ins)}/人{ratio_s}"))
        teams.append({
            "dest": re.sub(r"\s+", "", str(B or "")),
            "dept": str(D or ""),
            "people": people,
            "t_desc": t_desc,
            "t_sum": num("M", t_sum),
            "l_desc": l_desc,
            "l_sum": num("O", l_sum),
            "i_desc": i_desc,
            "i_sum": num("Q", i_sum),
            "total": num("R", t_sum + l_sum + i_sum),
            "remark": txt("S", ("就协全部" if ratio_of(D) == 1 else "院系 1/3，剩余就协")
                          + ("；部分经费由地方支持" if per_lodge == 0 else "")),
        })
        r += 1
    return teams

def set_font(run, size=10.5, bold=False, name="仿宋_GB2312"):
    run.font.size = Pt(size); run.font.bold = bold; run.font.name = name
    run.font.element.rPr.rFonts.set(qn("w:eastAsia"), name)

def add_cell_text(cell, text, size=9, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]; p.alignment = align
    set_font(p.add_run(text), size=size, bold=bold, name="宋体")

def build_docx(teams, title, season, out):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(29.7), Cm(21)  # A4 横向
    sec.left_margin = sec.right_margin = Cm(1.5)

    # 标题
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run(title), size=18, bold=True, name="方正大标宋简体")

    # 统筹文字（只写我们部分；基层/基研会部分留占位）
    total = sum(t["total"] for t in teams)
    body = (f"职业发展中心于{season}组织【__】支基层实习实践支队、{len(teams)}支企业主题实践支队。"
            f"其中，基层实习实践支队实践预算为【__】元（基研会部分，另行填写）；"
            f"企业主题实践支队{len(teams)}支，预算共{money(total)}元。"
            f"实践预算总额为【__】元，详情如下：")
    p = doc.add_paragraph()
    set_font(p.add_run(body), size=12)

    # 表格
    tb = doc.add_table(rows=3 + len(teams) + 1, cols=11)
    tb.style = "Table Grid"
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER

    # R0 大标题行
    c = tb.rows[0].cells[0].merge(tb.rows[0].cells[10])
    add_cell_text(c, "企业实习实践支队", size=11, bold=True)
    # R1 组表头
    heads1 = ["序号", "支队目的地", "人数", "交通", None, "住宿", None, "保险", None, "总计", "备注"]
    for i, h in enumerate(heads1):
        if h is not None:
            add_cell_text(tb.rows[1].cells[i], h, bold=True)
    for a, b in ((3, 4), (5, 6), (7, 8)):
        tb.rows[1].cells[a].merge(tb.rows[1].cells[b])
    for i in (0, 1, 2, 9, 10):  # 与 R2 纵向合并
        tb.rows[1].cells[i].merge(tb.rows[2].cells[i])
    # R2 子表头
    add_cell_text(tb.rows[2].cells[3], "单人往返费用", bold=True)
    add_cell_text(tb.rows[2].cells[4], "合计", bold=True)
    add_cell_text(tb.rows[2].cells[5], "单间费用", bold=True)
    add_cell_text(tb.rows[2].cells[6], "合计", bold=True)
    add_cell_text(tb.rows[2].cells[7], "单人费用", bold=True)
    add_cell_text(tb.rows[2].cells[8], "合计", bold=True)

    # 数据行
    for idx, t in enumerate(teams, start=1):
        row = tb.rows[2 + idx].cells
        vals = [str(idx), f'{t["dest"]}\n{t["dept"]}', str(t["people"]),
                t["t_desc"], money(t["t_sum"]), t["l_desc"], money(t["l_sum"]),
                t["i_desc"], money(t["i_sum"]), money(t["total"]), t["remark"]]
        for i, v in enumerate(vals):
            add_cell_text(row[i], v)

    # 合计行
    last = tb.rows[-1]
    c1 = last.cells[0].merge(last.cells[1]).merge(last.cells[2]).merge(last.cells[3])
    add_cell_text(c1, "合计", bold=True)
    c2 = last.cells[4].merge(last.cells[5]).merge(last.cells[6]).merge(last.cells[7]) \
         .merge(last.cells[8]).merge(last.cells[9]).merge(last.cells[10])
    add_cell_text(c2, money(total), bold=True)

    # 列宽
    widths = [1.2, 4.5, 1.3, 3.0, 1.8, 3.6, 1.8, 2.4, 1.6, 1.8, 3.2]
    for row in tb.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)

    # 注
    doc.add_paragraph()
    note1 = ("注：（1）上述预算中交通费用、保险费用为0的支队是表示由实习（实践）当地部门负责承担或同学自行承担交通费用和保险费用。"
             "交通费用一般按照高铁（二等座）来计算，更适合飞机出行的基地（路程远、机票费用低于高铁二等座）则按照机票费用来计算；"
             "保险费用按照平安保险-短期综合意外险费用计算。")
    note2 = ("（2）与院系联合组织支队按照支持上限为“北京往返实践地交通费用+当地住宿费用总和2/3”的标准执行，"
             "不含保险、队旗及其他费用。")
    for n in (note1, note2):
        p = doc.add_paragraph()
        set_font(p.add_run(n), size=10.5)

    doc.save(out)
    return total

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--sheet", default="各支队预算按实际人数校准")
    ap.add_argument("--start-row", type=int, default=7)
    ap.add_argument("--title", default="2026年暑假“清华学子地域行”实践预算")
    ap.add_argument("--season", default="2026年暑期")
    args = ap.parse_args()

    teams = read_teams(args.input, args.sheet, args.start_row)
    out = args.output or args.title + ".docx"
    total = build_docx(teams, args.title, args.season, out)
    print(f"已生成 {out}：{len(teams)} 支企业主题实践支队，预算合计 {money(total)} 元")
    for i, t in enumerate(teams, 1):
        print(f"  {i:>2}. {t['dest']:<12} {t['people']:>2}人 总计 {money(t['total']):>6}  [{t['remark']}]")

if __name__ == "__main__":
    main()
