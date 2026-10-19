# -*- coding: utf-8 -*-
"""
WorkBuddy提供
把 xls/html 格式的"客户-商品"二维表转成 Excel，并套好格式：
    1. 商品编号列（A列）从商品名称里拆出来，去掉括号，只留编号
    2. 商品名称列（B列）放名称（含装箱信息如 "1箱=12瓶"），列宽可调（默认 100）
    3. 各客户明细数字写成真正的"数值"，可以求和/合计
    4. "合计"列（C列）原样照抄原表文字（如 "14瓶"），不转成数字
    5. 表头行自动换行

列布局：
    A=商品编号   B=商品名称   C=合计   D 往后=各客户明细

用法：
    c.py  [列宽]  <输入文件>

示例：
    c.py 1.xls                  # 读 1.xls，输出 1_结果.xlsx，商品名称列宽 = 100
    c.py 120 1.xls              # 读 1.xls，输出 1_结果.xlsx，商品名称列宽 = 120
    c.py --width 120 1.xls      # 同上（命名参数写法）
    c.py                        # 不给参数时，弹提示手动输入文件名

也可以直接把文件拖到 c.py / c.exe 上，会自动当成输入文件。
"""
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup

# ======================= 默认配置（改这里即可） =======================
DEFAULT_WIDTH = 80.0       # 商品名称列（B列）的默认宽度
OUT_SUFFIX = "_结果"         # 输出文件名后缀：a.xls -> a_结果.xlsx
SHEET_NAME = "Sheet1"       # 工作表名字

CODE_COL_WIDTH = 15.0       # 商品编号列（A列）宽度
DATA_COL_WIDTH = 9.0        # 客户明细列（D列往后）宽度
HEADER_ROW_HEIGHT = 42.75   # 表头行高（配合自动换行才好看）
NUMBER_FORMAT = "0_ "       # 数字单元格格式（"数值"格式）

CODE_HEADER = "商品编号"     # A1 表头文字
NAME_HEADER = "商品名称"     # B1 表头文字
TOTAL_HEADER = "合计"        # C1 表头文字
TOTAL_HEADER2 = "合计_数值"        # C1 表头文字

CODE_COL = 1                # 商品编号列
NAME_COL = 2                # 商品名称列
TOTAL_COL = 3               # 合计列(原样输出)
TOTAL_COL2 = 4               # 合计列2(转为数值输出)
FIRST_CUSTOMER_COL = 5      # 第一个客户列（E列）
# =====================================================================

# 纯数字（含小数）的参数会被当成"列宽"
NUMBER_RE = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)$")

# 括号里的内容（全角/半角括号都认）——用来从商品单元格里挑出商品编号
BRACKET_RE = re.compile(r"[（(]\s*([^（()）]+?)\s*[）)]")


def to_float(text, default):
    """参数文本 -> float，转不了就用默认值。"""
    try:
        return float(str(text).strip().strip('"'))
    except (TypeError, ValueError):
        return default


def parse_args(argv):
    """解析启动参数，返回 (名称列宽度, 输入文件名)。文件名为 None 表示没传。"""
    width = DEFAULT_WIDTH
    filename = None
    i = 0
    while i < len(argv):
        arg = argv[i].strip().strip('"')
        if arg in ("-h", "--help", "-?", "/?", "/h"):
            print_usage()
            sys.exit(0)
        if arg in ("-w", "--width"):
            if i + 1 < len(argv):
                width = to_float(argv[i + 1], DEFAULT_WIDTH)
                i += 2
                continue
        if filename is None:
            if os.path.isfile(arg):          # 真实存在的文件（含拖拽过来的路径）
                filename = arg
            elif NUMBER_RE.match(arg):       # 纯数字 -> 列宽
                width = to_float(arg, DEFAULT_WIDTH)
            else:                            # 其它 -> 文件名
                filename = arg
        i += 1
    return width, filename


def build_output_path(src):
    """输入文件名 + "_结果" + .xlsx，输出到源文件同一个目录。"""
    folder = os.path.dirname(os.path.abspath(src))
    stem = os.path.splitext(os.path.basename(src))[0]
    return os.path.join(folder, stem + OUT_SUFFIX + ".xlsx")


def read_text(path):
    """读文本，utf-8 读不了就依次尝试其它编码。"""
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


def to_number(text):
    """'10瓶' -> 10（真数字）；空 -> None；不是数字就原样返回文本。"""
    if text is None:
        return None
    t = str(text)
    for unit in ("瓶", "箱", "件", "个","条","袋","包","板","杯"):    # 去掉单位汉字
        t = t.replace(unit, "")
    t = t.replace(",", "").strip()           # 去掉千分位逗号和空格
    if t == "":
        return None
    if re.fullmatch(r"[+-]?\d+", t):         # 整数
        return int(t)
    if re.fullmatch(r"[+-]?\d+\.\d+", t):    # 小数
        return float(t)
    return t                                 # 不是纯数字就保留原文，避免丢数据


def split_product(td):
    """把一个商品表头单元格拆成 (商品编号, 商品名称)。

    单元格结构通常是：
        名称(规格)<br/>(商品编号)<div class="tbTips"><span>1箱=12瓶</span></div>
    即：<br> 之前是名称，之后是编号，div.tbTips 里是装箱信息。
    编号去掉括号只留编号本身；装箱信息接回名称后面。
    """
    # 装箱信息（1箱=12瓶），最后要接在名称后面
    tips_div = td.find("div", class_="tbTips")
    tips = tips_div.get_text(strip=True) if tips_div is not None else ""

    before, after = [], []
    passed_br = False
    for node in td.children:
        name = getattr(node, "name", None)
        if name == "br":
            passed_br = True
            continue
        if name == "div":            # div.tbTips 已单独处理过，跳过
            continue
        text = str(node).strip()
        if text:
            (after if passed_br else before).append(text)

    product_name = "".join(before).strip()
    code_block = "".join(after).strip()

    if code_block:
        # 正常情况：<br> 后面那段就是 "(商品编号)"
        m = BRACKET_RE.search(code_block)
        code = (m.group(1) if m else code_block).strip()
    else:
        # 没有 <br> 的模板：整段里挑"最后一段长度>=4的括号内容"当编号
        # （长度限制是为了避开名称里的规格括号，如 (12)、（40））
        matches = [m for m in BRACKET_RE.finditer(product_name)
                   if len(m.group(1)) >= 4]
        code = ""
        if matches:
            last = matches[-1]
            code = last.group(1).strip()
            product_name = (product_name[:last.start()]
                            + product_name[last.end():]).strip()

    return code, product_name


def parse_html(html):
    """解析表格，返回 编号列表 / 名称列表 / 客户名列表 / 各商品合计 / 明细二维列表。"""
    soup = BeautifulSoup(html, "html.parser")
    tr_list = soup.find_all("tr")
    # 表格结构：第 0 行表头，最后两行是"合计""凑整"，中间是客户数据行
    if len(tr_list) < 3:
        raise ValueError("表格行数太少（只有 %d 行 <tr>），不是预期的格式" % len(tr_list))

    header_tr = tr_list[0]
    data_trs = tr_list[1:-2]
    total_tr = tr_list[-2]

    # 商品列：表头行从第 4 个 td 开始（前 3 个是 客户编号/客户名称/小计）
    product_codes = []
    product_names = []
    for td in header_tr.find_all("td")[3:]:
        code, name = split_product(td)
        product_codes.append(code)
        product_names.append(name)

    # 客户名称 + 明细：逐行一起取，保证两边下标对得上
    customer_names = []
    data_rows = []
    for tr in data_trs:
        tds = tr.find_all("td")
        if len(tds) < 2:
            continue
        customer_names.append(tds[1].get_text(strip=True))
        data_rows.append([td.get_text(strip=True) for td in tds[3:]])

    # 各商品合计：合计行从第 2 个 td 开始（第 1 个是"合计"两个字）
    total_values = [td.get_text(strip=True) for td in total_tr.find_all("td")[1:]]

    return product_codes, product_names, customer_names, total_values, data_rows


def build_workbook(product_codes, product_names, customer_names,
                   total_values, data_rows, width):
    """建表 + 填数据 + 套格式，返回 workbook。"""
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = SHEET_NAME

    # 第 1 行表头：A=商品编号，B=商品名称，C=合计，D 列开始依次是客户名称
    sheet.cell(1, CODE_COL).value = CODE_HEADER
    sheet.cell(1, NAME_COL).value = NAME_HEADER
    sheet.cell(1, TOTAL_COL).value = TOTAL_HEADER
    sheet.cell(1, TOTAL_COL2).value = TOTAL_HEADER2
    for i, name in enumerate(customer_names):
        sheet.cell(1, FIRST_CUSTOMER_COL + i).value = name

    # 每个商品占一行（从第 2 行开始）：A 列商品编号，B 列商品名称，C 列该商品合计
    # C 列按用户要求"原样输出"，直接照抄原表合计行的文字（如 "14瓶"），不做数字转换
    for i, name in enumerate(product_names):
        row = i + 2
        if i < len(product_codes):
            sheet.cell(row, CODE_COL).value = product_codes[i]
        sheet.cell(row, NAME_COL).value = name
        if i < len(total_values):
            text = total_values[i]
            if text:
                sheet.cell(row, TOTAL_COL).value = text
        if i < len(total_values):
            value = to_number(total_values[i])
            if value is not None:
                sheet.cell(row, TOTAL_COL2).value = value

    # 明细区：第 i 个商品 × 第 j 个客户，写到 (i+2) 行、(j+4) 列
    for j, row_values in enumerate(data_rows):
        for i, text in enumerate(row_values):
            value = to_number(text)
            if value is None:        # 空单元格不写，保持空白
                continue
            sheet.cell(i + 2, FIRST_CUSTOMER_COL + j).value = value

    # ---- 格式 1：A 列（编号）固定宽，B 列（名称）可调宽，客户列固定宽 ----
    sheet.column_dimensions[get_column_letter(CODE_COL)].width = CODE_COL_WIDTH
    sheet.column_dimensions[get_column_letter(NAME_COL)].width = width
    for col in range(FIRST_CUSTOMER_COL, sheet.max_column + 1):
        sheet.column_dimensions[get_column_letter(col)].width = DATA_COL_WIDTH

    # ---- 格式 2：表头行自动换行 + 加大行高 ----
    for col in range(1, sheet.max_column + 1):
        sheet.cell(1, col).alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[1].height = HEADER_ROW_HEIGHT

    # ---- 格式 3：客户明细列（D列往后）设成数值格式；A/B/C 列只居中 ----
    # （C 列现在是文本，套数值格式没意义，只给它居中）
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1):
        for cell in row:
            if cell.column >= FIRST_CUSTOMER_COL:
                cell.number_format = NUMBER_FORMAT
            if cell.column <= NAME_COL:  # A、B 两列
                cell.alignment = Alignment(horizontal="left", vertical="center",
                                           wrap_text=True)
            else:  # C 列往后
                cell.alignment = Alignment(horizontal="center", vertical="center",
                                           wrap_text=True)

    return wb



def main():
    width, src = parse_args(sys.argv[1:])

    if src is None:
        return 1
    if not src:
        return 1
    if not os.path.isfile(src):
        print("找不到文件：%s" % src)
        return 1

    out_path = build_output_path(src)

    html_content = read_text(src)
    product_codes, product_names, customer_names, total_values, data_rows = \
        parse_html(html_content)
    wb = build_workbook(product_codes, product_names, customer_names,
                        total_values, data_rows, width)
    wb.save(out_path)

    print("已生成：%s" % out_path)
    print("商品数：%d，客户数：%d，商品名称列宽：%s" % (len(product_names),
                                            len(customer_names), width))
    return 0


if __name__ == "__main__":
    code = main()
    sys.exit(code)
