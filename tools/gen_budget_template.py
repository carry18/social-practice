#!/usr/bin/env python3
"""生成预算表模板.xlsx —— 布局与 2026年暑假预算表0626「各支队预算按实际人数校准」一致。

A-H   院系提供：序号/实践地点/拟主要参访点位/承办院系/拟参访日期/支队人数/(空)/院系已有支持金额
I-S   主要呈现列（公式）：实践地点/参访天数/支队人数/往返交通费+合计/住宿+合计/保险+合计/总计/备注
T-AC  计算辅助列：单人交通费(T 为输入) … 队伍总额/报销比例（校研团不重复支持，已移除）
AD    交通费取价依据：逐段写明票价来源（12306实时车次+查询日期 / 缓存），rebuild_budget.py 自动填写
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUT = "预算表模板.xlsx"
DATA_START, N_ROWS = 7, 20
DATA_END = DATA_START + N_ROWS - 1

wb = Workbook()
ws = wb.active
ws.title = "各支队预算按实际人数校准"

thin = Side(style="thin")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
fill_input = PatternFill("solid", fgColor="FFF2CC")   # A-H 院系提供：淡黄
fill_show = PatternFill("solid", fgColor="D9E1F2")    # I-S 呈现列：淡蓝
fill_aux = PatternFill("solid", fgColor="E2EFDA")     # T-AC 辅助列：淡绿
fill_t_input = PatternFill("solid", fgColor="FFE699") # T 单人交通费：辅助列中唯一输入格
fill_demo = PatternFill("solid", fgColor="F2F2F2")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ---- 第 1 行：区域标题 ----
ws.merge_cells("A1:H1"); ws["A1"] = "院系提供"
ws.merge_cells("I1:S1"); ws["I1"] = "主要呈现列（公式，勿手改）"
ws.merge_cells("T1:AC1"); ws["T1"] = "计算辅助列"
for c, f in (("A1", fill_input), ("I1", fill_show), ("T1", fill_aux)):
    ws[c].font = Font(bold=True, size=12); ws[c].alignment = center; ws[c].fill = f

# ---- 第 2 行：说明性小表头 ----
r2 = {"A": "序号", "B": "实践地点", "C": "拟主要参访点位", "D": "承办单位", "E": "拟参访日期",
      "F": "支队人数", "H": "院系已有支持金额", "I": "实践地点", "J": "拟参访天数", "K": "支队人数",
      "L": "往返交通费", "N": "住宿", "P": "保险", "R": "总计",
      "T": "单人交通费", "U": "交通费总和", "V": "住宿费标间", "W": "房间数", "X": "住宿天数",
      "Y": "住宿费用总和", "Z": "保险费每人", "AA": "保险费总额", "AB": "队伍总额",
      "AD": "交通费取价依据"}
for k, v in r2.items():
    ws[f"{k}2"] = v; ws[f"{k}2"].font = Font(bold=True, size=10); ws[f"{k}2"].alignment = center

# ---- 第 3 行：填写说明 ----
notes = {
    "A": "填写说明",
    "C": "1.国防军工点位在前（填写企业全称），按照一天至少2个企业申报\n2.根据实践部反馈意见修改的，需确保符合意见",
    "D": "承办学院（就协组织填“就协”）",
    "E": "（格式：7.19-7.23）",
    "F": "默认12，可按实际调整",
    "H": "如院系支持1/3，此处等于总计*1/3",
    "L": "可在计算辅助列填写后直接套用下方单元格公式",
    "T": "北京-实践地各段高铁二等座最高价之和（多城市逐段相加）",
    "AD": "逐段：票价[12306实时·车次·查询日期]；缓存价标注来源",
    "V": "每人每天住宿报销上限150（标间300/2人）",
    "Z": "保险费每人每周不超过30元（≤7天即30元/人）",
}
ws.merge_cells("L3:P3")
for k, v in notes.items():
    ws[f"{k}3"] = v; ws[f"{k}3"].alignment = center; ws[f"{k}3"].font = Font(size=9, color="808080")

# ---- 公式模板（r = 数据行号；guarded=True 时空行显示空白） ----
def days_formula(r):
    """由 E 列 '7.19-7.23' 解析参访天数（沿用原表公式）。"""
    e = f"E{r}"
    return (f'=IFERROR(DATE(YEAR(TODAY()),MID({e},FIND("-",{e})+1,FIND(".",{e},FIND("-",{e}))-FIND("-",{e})-1),'
            f'RIGHT({e},LEN({e})-FIND(".",{e},FIND("-",{e}))))'
            f'-DATE(YEAR(TODAY()),LEFT({e},FIND(".",{e})-1),MID({e},FIND(".",{e})+1,FIND("-",{e})-FIND(".",{e})-1))+1,"待填")')

def formulas(r, guarded):
    f = {
        "I": f"=B{r}",
        "J": days_formula(r),
        "K": f"=F{r}",
        "L": f'=T{r}&"/人"&"x"&TEXT(AC{r},"# ?/?")',
        "M": f"=U{r}",
        "N": f'=V{r}/2&"/人·天*"&TRIM(TEXT(AC{r},"# ?/?"))&"*"&X{r}&"晚"',
        "O": f"=Y{r}",
        "P": f'=Z{r}&"/人·周*"&CEILING(J{r}/7,1)&"周"&"x"&TEXT(AC{r},"# ?/?")',
        "Q": f"=AA{r}",
        "R": f"=AB{r}",
        "S": f'=IF(OR(ISNUMBER(SEARCH("就协",D{r})),ISNUMBER(SEARCH("就业协会",D{r}))),'
             f'"就协全部","院系"&TEXT(1-AC{r},"# ?/?")&"，剩余就协")',
        "U": f"=K{r}*T{r}*AC{r}",
        "V": 300,
        "W": f"=CEILING(K{r}/2,1)",
        "X": f'=IF(ISNUMBER(J{r}),J{r}-1,"")',
        "Y": f"=V{r}*W{r}*X{r}*AC{r}",
        "Z": 30,
        "AA": f"=Z{r}*K{r}*CEILING(J{r}/7,1)*AC{r}",
        "AB": f"=U{r}+Y{r}+AA{r}",
        "AC": f'=IF(OR(ISNUMBER(SEARCH("就协",D{r})),ISNUMBER(SEARCH("就业协会",D{r}))),1,2/3)',
    }
    if guarded:  # B 列为空则整行空白（V/Z 常量除外）
        for col in ["I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
                    "U", "W", "X", "Y", "AA", "AB", "AC"]:
            f[col] = f'=IF($B{r}="","",{f[col][1:]})'
    return f

# ---- 第 4 行：示例（重庆，与原表一致） ----
demo = {"A": "例如", "B": "重庆",
        "C": "重庆长安汽车、中冶赛迪、重庆市华翌智能装备有限公司、中电科芯片技术公司、中冶赛迪工程技术股份有限公司",
        "D": "工物系", "E": "12.9-12.13", "F": 10, "T": 1922,
        "AD": "北京→重庆 961元[12306实时·G387·2026-10-22查]；重庆→北京 961元[12306实时·G51·2026-10-22查]"}
for col, v in {**demo, **formulas(4, guarded=False)}.items():
    ws[f"{col}4"] = v
    ws[f"{col}4"].fill = fill_demo
    ws[f"{col}4"].alignment = center

# ---- 第 6 行：正式表头 ----
headers = (["序号", "实践地点", "拟主要参访点位", "承办院系", "参访日期", "支队人数", "", "院系已有支持金额"]
           + ["实践地点", "参访天数", "支队人数", "往返交通费", "合计", "住宿", "合计", "保险", "合计", "总计", "备注"]
           + ["单人交通费", "交通费总和", "住宿费标间", "房间数", "住宿天数", "住宿费用总和",
              "保险费每人", "保险费总额", "队伍总额", "报销比例", "交通费取价依据"])
for i, h in enumerate(headers, start=1):
    cell = ws.cell(row=6, column=i, value=h or None)
    cell.font = Font(bold=True); cell.alignment = center; cell.border = border
    cell.fill = fill_input if i <= 8 else (fill_show if i <= 19 else fill_aux)

# ---- 数据行（空行不显示内容） ----
for r in range(DATA_START, DATA_END + 1):
    for col, v in formulas(r, guarded=True).items():
        ws[f"{col}{r}"] = v
    for c in range(1, 31):
        cell = ws.cell(row=r, column=c)
        cell.border = border
        cell.alignment = Alignment(vertical="center", wrap_text=(c in (3, 19, 30)))
        if c <= 8:
            cell.fill = fill_input
        elif c == 20:  # T 单人交通费
            cell.fill = fill_t_input

# ---- 对账区 ----
chk = DATA_END + 2
ws.cell(row=chk, column=17, value="过会金额").font = Font(bold=True)          # Q 列填金额
ws.cell(row=chk + 1, column=17, value="预算合计").font = Font(bold=True)
ws.cell(row=chk + 1, column=18, value=f"=SUM(R{DATA_START}:R{DATA_END})")     # R 列
ws.cell(row=chk + 2, column=17, value="差额").font = Font(bold=True)
ws.cell(row=chk + 2, column=18, value=f"=R{chk}-R{chk+1}")

# ---- 列宽 ----
widths = {"A": 6, "B": 12, "C": 40, "D": 16, "E": 12, "F": 9, "G": 3, "H": 13,
          "I": 12, "J": 9, "K": 9, "L": 14, "M": 10, "N": 18, "O": 10, "P": 12,
          "Q": 10, "R": 11, "S": 26}
for c in range(20, 30):
    widths[get_column_letter(c)] = 11
widths["AD"] = 70
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = "B7"

wb.save(OUT)
print("saved", OUT)
