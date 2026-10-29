# -*- coding: utf-8 -*-
"""普通 Markdown → Excel：**只有表格**转成 Excel 表格，其余一律按 Markdown 源码纯文本原样显示。

与 doubao-design-doc-excel 的区别（互不影响）：
  - doubao-design-doc-excel：带固定设计书版式模板（前 6 行头部、列段合并、配色体系）——本脚本不涉及
  - 本脚本：不带任何模板，输出干净的新工作簿；表格以外的内容**不做任何美化**，原文照搬

用户需求约束（不可违反）：
  1. Markdown 管道表格（含 GFM 对齐标记 `:---` / `:---:` / `---:`）→ 真正的 Excel 单元格表格
  2. 除表格以外的一切（标题、段落、列表、引用、代码块、分隔线、空行……）全部以
     『Markdown 源码』原文形式写进单元格，不加样式、不改写文字、不剥掉 # / ** / ` 等标记
  3. 不改动 doubao-design-doc-excel 这个技能

用法：
    python md_to_plain_excel.py --input <源.md> [--output <目标.xlsx>] [options]

参数：
    --input PATH           输入 Markdown（UTF-8，必需）
    --output PATH          输出 xlsx（缺省 ./output/<源文件名>.xlsx）
    --sheet-name NAME      sheet 名（缺省取源文件名，超 31 字符自动截断）
    --font NAME            字体名（缺省 宋体）
    --font-size N          字号（缺省 11）
    --blank-lines MODE     keep（缺省，源里几个空行就留几个空行）| collapse（连续空行并成 1 行）| drop（全部丢弃）
    --excel-table          额外把每张表注册为真正的 Excel 表格对象（ListObject，带筛选/条纹）
    --min-col-width N      表格列宽下限（缺省 5）
    --max-col-width N      表格列宽上限（缺省 60，超出部分靠自动换行折行）
    --merge-layout         固定基础列宽模式：所有列一律等宽（--base-col-width），
                           每张表按自己内容把若干基础列合并成宽单元格，列宽不再被某张表撑大
    --base-col-width N     merge-layout 模式下的基础列宽（缺省 10）
    --verify               生成后自动回读 xlsx 做保真校验（A列纯文本逐行一致 + 表格行数一致）
    --force                输出文件已存在时允许覆盖（缺省拒绝）
"""
import argparse
import math
import re
import sys
import unicodedata
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont

THIN = Side(style="thin", color="FF808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="FFF2F2F2")

LINE_PT = 14.5   # 单行文字高度（11pt）
PAD_PT = 4.0
MIN_ROW_PT = 17.0

DELIM_CELL_RE = re.compile(r"^:?-+:?$")
FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})")
BAD_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")
BOLD_RE = re.compile(r"(\*\*[^*]+\*\*)")
BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
# 行內 token：先認反引號代碼 span（`...`，字面保留），再認 **bold**（剝 ** 加粗）
INLINE_TOKEN_RE = re.compile(r"(`[^`]*`)|(\*\*[^*]+\*\*)")


def render_text_line(line, font_name, font_size, strip_bold=False):
    """渲染一行非表格 Markdown 源码，返回要写入单元格的值。

    不做标题样式：标题行(#)与正文同字体、同字号、不加粗（以 SKILL.md frontmatter 为准）。
    - strip_bold=False（缺省，原文照搬）：行内 **...** 段加粗（** 本身保留）
    - strip_bold=True：**...** 渲染成真粗体并剥掉 **；反引号 `...` 内的 ** 不动
    """
    if strip_bold:
        rt = render_inline(line, font_name, font_size)
        return rt if rt is not None else line
    if "**" not in line:
        return line
    normal = InlineFont(rFont=font_name, sz=font_size)
    bold = InlineFont(rFont=font_name, sz=font_size, b=True)
    blocks = []
    for part in BOLD_RE.split(line):
        if part == "":
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            blocks.append(TextBlock(bold, part))
        else:
            blocks.append(TextBlock(normal, part))
    return CellRichText(blocks)


def render_inline(text, font_name, size):
    """行内渲染：**bold** → 真粗体并剥掉 **；反引号 `...` 内的 ** 原样保留。

    支持 **`code`** 这种「粗体包裹行内代码」的嵌套写法。
    返回 CellRichText；整段没有任何标记时返回 None（由调用方回落纯字符串）。
    """
    normal = InlineFont(rFont=font_name, sz=size)
    bold = InlineFont(rFont=font_name, sz=size, b=True)
    blocks = []
    pos = 0
    for m in INLINE_TOKEN_RE.finditer(str(text)):
        if m.start() > pos:
            blocks.append(TextBlock(normal, str(text)[pos:m.start()]))
        if m.group(1) is not None:
            # 反引号代码 span：原样保留（含里面的 **）
            blocks.append(TextBlock(normal, m.group(1)))
        else:
            # **bold**：剥掉外层 **，内容（可能含反引号）整体加粗
            blocks.append(TextBlock(bold, m.group(2)[2:-2]))
        pos = m.end()
    if pos < len(str(text)):
        blocks.append(TextBlock(normal, str(text)[pos:]))
    if not blocks:
        return None
    return CellRichText(blocks)


def rich_text_to_str(val):
    """把 CellRichText / TextBlock / 普通字符串拍平成纯文本（verify 回读比对用）"""
    if val is None:
        return None
    if isinstance(val, str):
        return val
    if isinstance(val, CellRichText):
        return "".join(rich_text_to_str(b) for b in val)
    if isinstance(val, TextBlock):
        return val.text
    return str(val)


def set_value(cell, value):
    """写入单元格；以 = 开头的文字按纯文本保存，不让 openpyxl 当成公式"""
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"


# ---------------------------------------------------------------- 文本宽度
def disp_width(s) -> int:
    """显示宽度：全角/宽字符算 2，其余算 1（与 Excel 列宽单位对齐）"""
    return sum(2 if unicodedata.east_asian_width(ch) in ("F", "W") else 1 for ch in str(s or ""))


def wrap_lines(text, avail_units) -> int:
    """模拟自动换行后占用的行数（逐字贪婪累加）"""
    total = 0
    for logical in str(text or "").split("\n"):
        if logical == "":
            total += 1
            continue
        avail = max(avail_units, 2)
        cur, lines = 0, 1
        for ch in logical:
            w = 2 if unicodedata.east_asian_width(ch) in ("F", "W") else 1
            if cur + w > avail:
                lines += 1
                cur = w
            else:
                cur += w
        total += lines
    return max(total, 1)


# ---------------------------------------------------------------- Markdown 行解析
def split_row(raw: str):
    """按未转义的 | 切分单元格；\\| 还原成 |；去掉首尾包裹的 |"""
    s = raw.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        k, j = 0, len(s) - 2
        while j >= 0 and s[j] == "\\":
            k += 1
            j -= 1
        if k % 2 == 0:          # 结尾的 | 未被转义 → 视为包裹符
            s = s[:-1]
    out, cur, i = [], [], 0
    while i < len(s):
        ch = s[i]
        if ch == "\\" and i + 1 < len(s) and s[i + 1] == "|":
            cur.append("|")
            i += 2
            continue
        if ch == "|":
            out.append("".join(cur).strip())
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    out.append("".join(cur).strip())
    return out


def has_pipe(raw: str) -> bool:
    """是否含有未转义的 |"""
    i = 0
    while i < len(raw):
        if raw[i] == "\\":
            i += 2
            continue
        if raw[i] == "|":
            return True
        i += 1
    return False


def is_delim_row(raw: str) -> bool:
    """是否为 GFM 表格分隔行，如 |---|:--:|---:|"""
    if "-" not in raw:
        return False
    if not has_pipe(raw):
        return False
    if not re.fullmatch(r"\s*\|?[\s:\-|]+\|?\s*", raw):
        return False
    cells = split_row(raw)
    return bool(cells) and all(DELIM_CELL_RE.match(c) for c in cells)


def col_align(cell: str):
    c = cell.strip()
    left, right = c.startswith(":"), c.endswith(":")
    if left and right:
        return "center"
    if right:
        return "right"
    if left:
        return "left"
    return None


# ---------------------------------------------------------------- 文档扫描
def parse(lines):
    """→ 顺序排列的块列表：('text', 原文行) / ('blank',) / ('table', headers, rows, aligns)"""
    blocks = []
    i, n = 0, len(lines)
    fence = None
    while i < n:
        line = lines[i]

        m = FENCE_RE.match(line)
        if fence is None and m:
            fence = m.group(1)[0]        # ``` 或 ~~~
            blocks.append(("text", line))
            i += 1
            continue
        if fence is not None:
            blocks.append(("text", line))
            if m and m.group(1)[0] == fence:
                fence = None
            i += 1
            continue

        # 表格：当前行含 |，且下一行是分隔行
        if has_pipe(line) and i + 1 < n and is_delim_row(lines[i + 1]):
            header = split_row(line)
            aligns = [col_align(c) for c in split_row(lines[i + 1])]
            i += 2
            body = []
            while i < n and lines[i].strip() != "" and has_pipe(lines[i]):
                body.append(split_row(lines[i]))
                i += 1
            ncol = max([len(header)] + [len(r) for r in body])
            header = (header + [""] * ncol)[:ncol]
            body = [(r + [""] * ncol)[:ncol] for r in body]
            aligns = (aligns + [None] * ncol)[:ncol]
            blocks.append(("table", header, body, aligns))
            continue

        blocks.append(("blank",) if line.strip() == "" else ("text", line))
        i += 1
    return blocks


def cell_value(text, font_name, font_size, strip_bold):
    """表格单元格取值：<br> 换成单元格内换行；strip_bold 时把 **bold** 渲成真粗体并剥掉 **，否则原样字符串"""
    text = BR_RE.sub("\n", text)
    if strip_bold:
        rt = render_inline(text, font_name, font_size)
        return rt if rt is not None else text
    return text


def strip_bold_markers(text):
    """纯文本对照用：把 **bold** 的 ** 剥掉（保留内容；反引号 `...` 内的 ** 不动）"""
    out = []
    pos = 0
    for m in INLINE_TOKEN_RE.finditer(str(text)):
        out.append(str(text)[pos:m.start()])
        if m.group(1) is not None:
            out.append(m.group(1))            # 代码 span 原样
        else:
            out.append(m.group(2)[2:-2])       # 粗体剥 **
        pos = m.end()
    out.append(str(text)[pos:])
    return "".join(out)


# ---------------------------------------------------------------- 渲染
def table_span(header, body, min_w, max_w, base_w):
    """merge-layout：按本表内容给每个 markdown 列分配基础列段数。

    返回 (starts, ks)：starts[j] = 该列起始 Excel 列号(1-based)，
    ks[j] = 该列横跨的基础列数（合并单元格数）。不同表各自独立，互不影响。
    """
    starts, ks = [], []
    s = 1
    for j in range(len(header)):
        vals = [header[j]] + [r[j] for r in body]
        mw = max((disp_width(v) for v in vals), default=0)
        desired = max(min_w, min(max_w, mw + 2))
        k = max(1, math.ceil(desired / base_w))
        starts.append(s)
        ks.append(k)
        s += k
    return starts, ks


def _style_merged(ws, row, c1, c2, value, font, fill, align):
    """写值到 (row,c1)，合并 (c1..c2)，并给合并区每个格子补边框/填充（openpyxl 要求逐格设置）"""
    c = ws.cell(row=row, column=c1)
    set_value(c, value)
    c.font = font
    if fill is not None:
        c.fill = fill
    c.alignment = align
    for cc in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=cc)
        cell.border = BORDER
        if fill is not None:
            cell.fill = fill
    if c2 > c1:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)


def build(blocks, font_name, font_size, blank_mode, excel_table,
          min_w, max_w, sheet_name, merge_layout=False, base_w=10, strip_bold=False):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    plain_font = Font(name=font_name, size=font_size)
    head_font = Font(name=font_name, size=font_size, bold=True)
    body_font = Font(name=font_name, size=font_size)

    # tables: (hrow, lrow, ncol, header, starts, ks)；starts/ks 仅 merge-layout 下有意义
    tables = []
    row = 1
    prev_blank = False
    for blk in blocks:
        if blk[0] == "blank":
            if blank_mode == "drop" or (blank_mode == "collapse" and prev_blank):
                continue
            prev_blank = True
            row += 1                       # 留一个空行，不加任何内容
            continue
        prev_blank = False

        if blk[0] == "text":
            # 原文照搬：不换行、不做标题样式；strip_bold 时 ** 渲染真粗体并剥掉
            c = ws.cell(row=row, column=1)
            set_value(c, render_text_line(blk[1], font_name, font_size, strip_bold))
            c.font = plain_font
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)
            row += 1
            continue

        _, header, body, aligns = blk
        ncol = len(header)
        hrow = row
        if merge_layout:
            starts, ks = table_span(header, body, min_w, max_w, base_w)
            for j in range(ncol):
                a = aligns[j]
                c1, c2 = starts[j], starts[j] + ks[j] - 1
                v = cell_value(header[j], font_name, font_size, strip_bold)
                _style_merged(ws, hrow, c1, c2, v, head_font, HEADER_FILL,
                               Alignment(horizontal=(a or "center"), vertical="center", wrap_text=True))
            row += 1
            for r in body:
                for j in range(ncol):
                    a = aligns[j]
                    c1, c2 = starts[j], starts[j] + ks[j] - 1
                    v = cell_value(r[j], font_name, font_size, strip_bold)
                    _style_merged(ws, row, c1, c2, v, body_font, None,
                                  Alignment(horizontal=(a or "left"), vertical="center", wrap_text=True))
                row += 1
            tables.append((hrow, row - 1, ncol, list(header), starts, ks))
        else:
            for j in range(ncol):
                a = aligns[j]
                v = cell_value(header[j], font_name, font_size, strip_bold)
                c = ws.cell(row=hrow, column=j + 1)
                set_value(c, v)
                c.font = head_font
                c.fill = HEADER_FILL
                c.border = BORDER
                c.alignment = Alignment(horizontal=(a or "center"), vertical="center", wrap_text=True)
            row += 1
            for r in body:
                for j in range(ncol):
                    a = aligns[j]
                    v = cell_value(r[j], font_name, font_size, strip_bold)
                    c = ws.cell(row=row, column=j + 1)
                    set_value(c, v)
                    c.font = body_font
                    c.border = BORDER
                    c.alignment = Alignment(horizontal=(a or "left"), vertical="center", wrap_text=True)
                row += 1
            tables.append((hrow, row - 1, ncol, list(header), None, None))

    if merge_layout:
        # 固定基础列宽：所有用到的列一律 base_w，不再按内容撑列
        max_end = 1
        for hrow, lrow, ncol, header, starts, ks in tables:
            end = max(starts[j] + ks[j] - 1 for j in range(ncol))
            max_end = max(max_end, end)
        for cc in range(1, max_end + 1):
            ws.column_dimensions[get_column_letter(cc)].width = base_w
        # 行高：按各列合并后的实际宽度估算
        for hrow, lrow, ncol, header, starts, ks in tables:
            for r in range(hrow, lrow + 1):
                lines = 1
                for j in range(ncol):
                    eff = ks[j] * base_w - 2
                    lines = max(lines, wrap_lines(ws.cell(row=r, column=starts[j]).value, eff))
                ws.row_dimensions[r].height = max(MIN_ROW_PT, lines * LINE_PT + PAD_PT)
    elif tables:
        # 列宽：同一列被多张表共用，取所有表的最大需求（不超出 [min_w, max_w]）
        need = {}
        for hrow, lrow, ncol, header, starts, ks in tables:
            for j in range(ncol):
                vals = [header[j]]
                for r in range(hrow + 1, lrow + 1):
                    vals.append(ws.cell(row=r, column=j + 1).value)
                need[j] = max(need.get(j, 0), max(disp_width(v) for v in vals) + 2)
        widths = {}
        for j, w in sorted(need.items()):
            widths[j] = max(min_w, min(max_w, w))
            ws.column_dimensions[get_column_letter(j + 1)].width = widths[j]

        # 行高：按最终列宽估算换行行数
        for hrow, lrow, ncol, header, starts, ks in tables:
            for r in range(hrow, lrow + 1):
                lines = 1
                for j in range(ncol):
                    w = widths.get(j, min_w)
                    lines = max(lines, wrap_lines(ws.cell(row=r, column=j + 1).value, w - 2))
                ws.row_dimensions[r].height = max(MIN_ROW_PT, lines * LINE_PT + PAD_PT)

    # 可选：注册为真正的 Excel 表格对象（ListObject，带筛选）
    if excel_table:
        if merge_layout:
            sys.exit("--excel-table 与 --merge-layout 不兼容：Excel 表格对象不允许跨合并单元格，请二选一。")
        idx = 0
        for hrow, lrow, ncol, header, starts, ks in tables:
            if lrow <= hrow:               # 只有表头没有数据行 → 跳过，避免 Excel 修复提示
                continue
            idx += 1
            names, seen = [], {}
            for j in range(ncol):
                h = (header[j] or "").strip().replace("\n", " ") or f"Column{j + 1}"
                if h in seen:
                    seen[h] += 1
                    h = f"{h}_{seen[h]}"
                else:
                    seen[h] = 1
                names.append(h)
            if names != header:            # 表头被规范化时同步回写单元格，保证与表格定义一致
                for j, h in enumerate(names):
                    ws.cell(row=hrow, column=j + 1, value=h)
            name = f"Tbl{idx}"
            ref = f"A{hrow}:{get_column_letter(ncol)}{lrow}"
            t = Table(displayName=name, ref=ref)
            t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False,
                                              showColumnStripes=False, showFirstColumn=False,
                                              showLastColumn=False)
            ws.add_table(t)

    return ws


def make_sheet_name(s: str) -> str:
    name = BAD_SHEET_CHARS.sub("", str(s or "").strip()).replace("：", "_").replace("/", "_")
    return (name or "Sheet1")[:31]


def expected_plain_rows(src_lines, strip_bold, blank_mode):
    """与 build() 同一套规则：源文件非表格行在 A 列应出现的值（空行 → None）"""
    out, prev_blank = [], False
    for blk in parse(src_lines):
        if blk[0] == "table":
            prev_blank = False
            continue
        if blk[0] == "blank":
            if blank_mode == "drop" or (blank_mode == "collapse" and prev_blank):
                continue
            prev_blank = True
            out.append(None)
            continue
        prev_blank = False
        ln = blk[1]
        if ln == "":
            out.append(None)
        else:
            out.append(strip_bold_markers(ln) if strip_bold else ln)
    return out


def verify(xlsx_path, src_lines, strip_bold=False, blank_mode="keep"):
    """回读已保存的 xlsx，校验：
    1. A 列非表格行 == 源文件非表格行（逐行、顺序、内容完全一致；strip_bold 时双方都剥掉 **；
       空行按 blank_mode 与生成时一致）
    2. 表格行数（表头+数据）与 parse() 结果一致
    返回 (ok: bool, detail: str)
    """
    from openpyxl import load_workbook
    wb2 = load_workbook(xlsx_path)
    ws2 = wb2.active
    # A 列纯文本行 = 第 1 列没有左边框的行
    plain_rows = [r for r in range(1, ws2.max_row + 1)
                  if ws2.cell(row=r, column=1).border.left.style is None]
    got = [rich_text_to_str(ws2.cell(row=r, column=1).value) for r in plain_rows]
    exp = expected_plain_rows(src_lines, strip_bold, blank_mode)
    if got != exp:
        diff = next(((k, a, b) for k, (a, b) in enumerate(zip(got, exp)) if a != b), None)
        return False, f"A列文本不一致: len {len(got)} vs {len(exp)}, first diff={diff}"
    # 表格行数 = 有左边框的行数；期望值从 parse() 重算
    tbl_rows = sum(1 for r in range(1, ws2.max_row + 1)
                   if ws2.cell(row=r, column=1).border.left.style == "thin")
    blocks = parse(src_lines)
    expected_tbl_rows = sum(1 + len(b[2]) for b in blocks if b[0] == "table")
    if tbl_rows != expected_tbl_rows:
        return False, f"表格行数不一致: xlsx={tbl_rows}, src(parse)={expected_tbl_rows}"
    return True, f"OK（A列 {len(got)} 行纯文本 + {tbl_rows} 行表格，全部一致）"


def main():
    ap = argparse.ArgumentParser(description="普通 Markdown → Excel（只有表格转成表格，其余原文显示）")
    ap.add_argument("--input", required=True, help="输入 Markdown 路径")
    ap.add_argument("--output", help="输出 xlsx 路径（缺省 ./output/<源文件名>.xlsx）")
    ap.add_argument("--sheet-name", help="sheet 名（缺省取源文件名）")
    ap.add_argument("--font", default="宋体", help="字体名（缺省 宋体）")
    ap.add_argument("--font-size", type=int, default=11, help="字号（缺省 11）")
    ap.add_argument("--blank-lines", choices=["keep", "collapse", "drop"], default="keep",
                    help="空行处理（缺省 keep：源里几个空行就留几个空行）")
    ap.add_argument("--excel-table", action="store_true",
                    help="把每张表注册为真正的 Excel 表格对象（ListObject，带筛选）")
    ap.add_argument("--min-col-width", type=int, default=5)
    ap.add_argument("--max-col-width", type=int, default=60)
    ap.add_argument("--merge-layout", action="store_true",
                    help="固定基础列宽模式：所有列等宽，每张表按内容合并基础列成宽单元格")
    ap.add_argument("--base-col-width", type=int, default=10,
                    help="merge-layout 模式下的基础列宽（缺省 10）")
    ap.add_argument("--strip-bold", action="store_true",
                    help="把 **bold** 渲成真粗体并剥掉 **（反引号内的 ** 不动）；缺省保留 ** 原文")
    ap.add_argument("--verify", action="store_true",
                    help="生成后自动回读 xlsx 做保真校验（A列纯文本 + 表格行数）")
    ap.add_argument("--force", action="store_true",
                    help="输出文件已存在时允许覆盖（缺省拒绝，避免误盖别的技能/工具产出的同名文件）")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.is_file():
        sys.exit(f"输入文件不存在: {src}")
    out = Path(args.output) if args.output else Path.cwd() / "output" / (src.stem + ".xlsx")
    if out.exists() and not args.force:
        sys.exit(f"输出文件已存在，未覆盖：{out}\n"
                 f"确认要覆盖请加 --force，或用 --output 指定其它路径（例如 xxx.plain.xlsx）。")
    out.parent.mkdir(parents=True, exist_ok=True)

    lines = src.read_text(encoding="utf-8").split("\n")
    if lines and lines[-1] == "":          # 末尾换行不算一行内容
        lines.pop()
    blocks = parse(lines)

    n_tables = sum(1 for b in blocks if b[0] == "table")
    n_text = sum(1 for b in blocks if b[0] == "text")
    n_blank = sum(1 for b in blocks if b[0] == "blank")
    n_rows = sum(len(b[2]) for b in blocks if b[0] == "table")

    ws = build(blocks, args.font, args.font_size, args.blank_lines, args.excel_table,
               args.min_col_width, args.max_col_width,
               make_sheet_name(args.sheet_name or src.stem),
               merge_layout=args.merge_layout, base_w=args.base_col_width,
               strip_bold=args.strip_bold)

    wb = ws.parent
    wb.save(out)
    print(f"input : {src}")
    print(f"lines : {len(lines)}  (text {n_text} / blank {n_blank} / table-block {n_tables})")
    print(f"tables: {n_tables} 张，数据行合计 {n_rows} 行")
    print(f"saved : {out}  (sheet={ws.title}, rows={ws.max_row}, cols={ws.max_column})")

    if args.verify:
        ok, detail = verify(out, lines, strip_bold=args.strip_bold, blank_mode=args.blank_lines)
        tag = "PASS" if ok else "FAIL"
        print(f"verify: {tag}  {detail}")
        if not ok:
            sys.exit(1)


if __name__ == "__main__":
    main()
