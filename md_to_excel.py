# -*- coding: utf-8 -*-
#
# ============================================================================
# 版权所有 (Copyright) © 李鹏虎 (Li Penghu). 保留所有权利 (All Rights Reserved).
#
# 本软件仅供个人学习与非商业用途使用。
# 未经作者书面同意，不得用于任何商业目的（包括但不限于销售、集成到商业产品、
# 提供商业服务等）。如需商业授权，请先联系作者征得同意。
#
# 作者邮箱 (Author E-mail): gs_lipenghu@qq.com
#
# THIS SOFTWARE IS FOR PERSONAL AND NON-COMMERCIAL USE ONLY.
# Commercial use of any kind is prohibited without the author's prior written
# consent. For commercial licensing, please contact the author first.
# ============================================================================
"""
Markdown -> Excel 结构化转换工具
================================

版权所有 © 李鹏虎  |  非商业授权  |  商业使用需经作者同意 (gs_lipenghu@qq.com)
--------------------------------------------------------------------------------

功能特性
--------
1. 层级缩进
   - 一级标题(#)从 A 列开始，二级(##)从 B 列开始，以此类推。
   - 列表按嵌套深度缩进；标题下方的正文/引用/表格/代码块也随所属标题层级缩进。
2. 默认格式
   - 所有单元格：列宽 1cm、行高 0.8cm、宋体 12 号、垂直居中。
3. Markdown 表格
   - 表格每一列自动合并 5 个单元格宽度（横向合并 5 格），表头加粗并居中。
4. 代码块
   - 无论内容多少，统一放入一个 “宽 20 列 × 高 12 行” 的合并单元格：
     深灰色背景、白色字体、宋体 10 号、左对齐、上对齐。
5. Mermaid 流程图
   - ```mermaid 代码块会先用 mermaid-cli 渲染成 PNG，再以图片形式嵌入 Excel。
   - 未安装 mermaid-cli 或渲染失败时，自动回退为普通代码块显示原始文本（内容不丢失）。
6. 全内容支持
   - 标题、段落、列表、代码块、表格、引用、图片、任务清单 全部处理。

用法
----
    python md_to_excel.py 输入.md 输出.xlsx
    python md_to_excel.py 输入.md              # 输出同名 .xlsx
    python md_to_excel.py                       # 生成内置示例 demo.md -> demo.xlsx

依赖:
    - openpyxl       (pip install openpyxl)             必需
    - Pillow         (pip install Pillow)               嵌入图片所需
    - mermaid-cli    (npm install -g @mermaid-js/mermaid-cli)  渲染 mermaid 所需(可选)
"""

import re
import sys
import os
import shutil
import tempfile
import subprocess

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from openpyxl.drawing.image import Image as XLImage
    _HAS_IMAGE = True
except Exception:
    _HAS_IMAGE = False


# ---------------------------------------------------------------------------
# 全局格式常量（可按需修改）
# ---------------------------------------------------------------------------
FONT_NAME = "宋体"
FONT_SIZE = 12               # 默认字号
CODE_FONT_SIZE = 10          # 代码块字号

COL_WIDTH_CM = 1.0           # 默认列宽 (cm)
ROW_HEIGHT_CM = 0.8          # 默认行高 (cm)

TABLE_COL_SPAN = 5           # 表格每列合并的单元格数
CODE_WIDTH_COLS = 20         # 代码块宽度(列)
CODE_HEIGHT_ROWS = 12        # 代码块高度(行)

CODE_BG_COLOR = "404040"     # 代码块背景：深灰
CODE_FONT_COLOR = "FFFFFF"   # 代码块字体：白色

HEADER_FILL_COLOR = "D9D9D9"  # 表头背景：浅灰
QUOTE_FILL_COLOR = "F2F2F2"   # 引用背景：更浅灰
LABEL_FILL_COLOR = "E7EEF7"   # 表格首列(标签列)背景：淡蓝
IMG_MAX_WIDTH_PX = 480        # 嵌入图片最大宽度(px)，超出按比例缩小

# Mermaid 图表（需本地安装 mermaid-cli: npm install -g @mermaid-js/mermaid-cli）
MERMAID_BG = "white"          # 生成图片的背景色
MERMAID_SCALE = 2             # 渲染缩放倍数（越大越清晰）
MERMAID_MAX_WIDTH_PX = 720    # 嵌入 Excel 时的最大宽度(px)


# ---------------------------------------------------------------------------
# 单位换算 (基于 96 DPI)
# ---------------------------------------------------------------------------
def cm_to_col_width(cm: float) -> float:
    """厘米 -> openpyxl 列宽(字符数)。近似公式，基于 96DPI / 7px 数字宽。"""
    pixels = cm / 2.54 * 96.0
    return round((pixels - 5) / 7.0, 2)


def cm_to_row_height(cm: float) -> float:
    """厘米 -> 行高(磅 pt)。1 pt = 1/72 inch。"""
    return round(cm / 2.54 * 72.0, 2)


DEFAULT_COL_WIDTH = cm_to_col_width(COL_WIDTH_CM)     # ≈ 4.68
DEFAULT_ROW_HEIGHT = cm_to_row_height(ROW_HEIGHT_CM)  # ≈ 22.68


# ---------------------------------------------------------------------------
# 内联 Markdown 清洗（转纯文本）
# ---------------------------------------------------------------------------
def strip_inline(text: str) -> str:
    """去除内联 Markdown 标记，保留可读文本。"""
    text = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", text)   # 图片 ![alt](url) -> alt
    text = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text)    # 链接 [text](url) -> text
    text = re.sub(r"`([^`]*)`", r"\1", text)                # 行内代码 `code`
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)          # 粗体 **
    text = re.sub(r"__([^_]+)__", r"\1", text)              # 粗体 __
    text = re.sub(r"\*([^*]+)\*", r"\1", text)              # 斜体 *
    text = re.sub(r"(?<!\w)_([^_]+)_(?!\w)", r"\1", text)   # 斜体 _
    text = re.sub(r"~~([^~]+)~~", r"\1", text)              # 删除线 ~~
    return text.strip()


# ---------------------------------------------------------------------------
# Markdown 解析：将文本拆成结构化块
# ---------------------------------------------------------------------------
def parse_markdown(md_text: str):
    """返回块列表。每个块是 dict，包含 type 及相关字段。"""
    lines = md_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # ---- 空行 ----
        if stripped == "":
            i += 1
            continue

        # ---- 围栏代码块 ``` 或 ~~~ ----
        fence = re.match(r"^(\s*)(`{3,}|~{3,})(.*)$", line)
        if fence:
            marker = fence.group(2)[0]
            lang = fence.group(3).strip()
            code_lines = []
            i += 1
            while i < n:
                if re.match(r"^\s*" + re.escape(marker) + r"{3,}\s*$", lines[i]):
                    i += 1
                    break
                code_lines.append(lines[i])
                i += 1
            content = "\n".join(code_lines)
            if lang.lower() == "mermaid":
                blocks.append({"type": "mermaid", "content": content})
            else:
                blocks.append({"type": "code", "lang": lang, "content": content})
            continue

        # ---- 独立图片行 ![alt](src) ----
        img_m = re.match(r"^!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)$", stripped)
        if img_m:
            blocks.append({"type": "image", "alt": img_m.group(1).strip(),
                           "src": img_m.group(2).strip()})
            i += 1
            continue

        # ---- 标题 # ----
        heading = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if heading:
            level = len(heading.group(1))
            blocks.append({"type": "heading", "level": level,
                           "text": strip_inline(heading.group(2))})
            i += 1
            continue

        # ---- 表格（当前行含 | 且下一行为分隔行）----
        if "|" in line and i + 1 < n and re.match(r"^\s*\|?[\s:|-]*-[\s:|-]*\|?\s*$", lines[i + 1]) and "-" in lines[i + 1]:
            table_lines = [line]
            sep_line = lines[i + 1]
            i += 2
            while i < n and "|" in lines[i] and lines[i].strip() != "":
                table_lines.append(lines[i])
                i += 1
            header = _split_table_row(table_lines[0])
            rows = [_split_table_row(r) for r in table_lines[1:]]
            blocks.append({"type": "table", "header": header, "rows": rows})
            continue

        # ---- 引用 > ----
        if re.match(r"^\s*>", line):
            quote_lines = []
            while i < n and re.match(r"^\s*>", lines[i]):
                quote_lines.append(re.sub(r"^\s*>\s?", "", lines[i]))
                i += 1
            text = strip_inline(" ".join(l.strip() for l in quote_lines if l.strip()))
            blocks.append({"type": "quote", "text": text})
            continue

        # ---- 列表项 ----
        list_m = re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$", line)
        if list_m:
            indent = len(list_m.group(1).expandtabs(4))
            depth = indent // 2                     # 每 2 空格算一级缩进
            marker = list_m.group(2)
            ordered = bool(re.match(r"\d+[.)]", marker))
            raw = list_m.group(3)
            # 任务列表 - [ ] / - [x]
            task = re.match(r"^\[([ xX])\]\s+(.*)$", raw)
            if task:
                checked = task.group(1).lower() == "x"
                text = strip_inline(task.group(2))
                blocks.append({"type": "list", "depth": depth, "ordered": False,
                               "task": True, "checked": checked,
                               "bullet": ("☑" if checked else "☐"), "text": text})
            else:
                text = strip_inline(raw)
                bullet = marker if ordered else "•"
                blocks.append({"type": "list", "depth": depth, "ordered": ordered,
                               "task": False, "checked": False,
                               "bullet": bullet, "text": text})
            i += 1
            continue

        # ---- 普通段落（合并连续的普通文本行）----
        para_lines = [stripped]
        i += 1
        while i < n:
            nxt = lines[i]
            nxt_s = nxt.strip()
            if nxt_s == "":
                break
            # 遇到其它块起始，停止
            if (re.match(r"^(#{1,6})\s+", nxt_s) or re.match(r"^\s*>", nxt) or
                    re.match(r"^(\s*)([-*+]|\d+[.)])\s+", nxt) or
                    re.match(r"^(\s*)(`{3,}|~{3,})", nxt) or
                    re.match(r"^!\[[^\]]*\]\([^)\s]+", nxt_s) or
                    ("|" in nxt and i + 1 < n and "-" in (lines[i + 1] if i + 1 < n else ""))):
                break
            para_lines.append(nxt_s)
            i += 1
        blocks.append({"type": "paragraph", "text": strip_inline(" ".join(para_lines))})

    return blocks


def _find_mmdc():
    """定位 mermaid-cli 可执行文件；返回可直接传给 subprocess 的命令前缀列表，找不到返回 None。"""
    for name in ("mmdc", "mmdc.cmd", "mmdc.ps1"):
        p = shutil.which(name)
        if p:
            low = p.lower()
            if low.endswith((".cmd", ".bat")):
                return ["cmd", "/c", p]
            if low.endswith(".ps1"):
                return ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", p]
            return [p]
    # 兜底：常见 npm 全局安装目录
    appdata = os.environ.get("APPDATA")
    if appdata:
        cand = os.path.join(appdata, "npm", "mmdc.cmd")
        if os.path.isfile(cand):
            return ["cmd", "/c", cand]
    return None


def _split_table_row(row: str):
    """拆分表格行为单元格列表。"""
    row = row.strip()
    if row.startswith("|"):
        row = row[1:]
    if row.endswith("|"):
        row = row[:-1]
    return [strip_inline(c.strip()) for c in row.split("|")]


# ---------------------------------------------------------------------------
# 渲染：结构化块 -> Excel
# ---------------------------------------------------------------------------
class ExcelRenderer:
    def __init__(self, base_dir="."):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "内容"
        self.base_dir = base_dir     # 图片相对路径解析基准目录
        self.row = 1                 # 当前写入行(1-based)
        self.section_level = 0       # 最近标题的层级(0 表示尚无标题)
        self.max_col = 1             # 记录用到的最大列
        self.code_cells = set()      # 代码块单元格坐标(白字，勿覆盖默认样式)
        self.mmdc_cmd = _find_mmdc() if _HAS_IMAGE else None  # Mermaid 渲染能力(启动时检测一次)
        self._mermaid_warned = False # 是否已提示过“缺少 Mermaid 支持”

        self.default_align = Alignment(vertical="center", wrap_text=False)
        self.default_font = Font(name=FONT_NAME, size=FONT_SIZE)

    # -- 工具 --
    def _track(self, col):
        if col > self.max_col:
            self.max_col = col

    def _style_range(self, r1, c1, r2, c2, font=None, fill=None, align=None):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.ws.cell(row=r, column=c)
                if font:
                    cell.font = font
                if fill:
                    cell.fill = fill
                if align:
                    cell.alignment = align
        self._track(c2)

    # -- 各类块渲染 --
    def render_heading(self, block):
        level = block["level"]
        col = level                       # H1->1(A), H2->2(B)... 以 1-based 列号
        self.section_level = level
        cell = self.ws.cell(row=self.row, column=col, value=block["text"])
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)  # 各级标题统一宋体12加粗
        cell.alignment = self.default_align
        self._track(col)
        self.row += 1

    def render_paragraph(self, block):
        col = max(self.section_level, 1)  # 正文随所属标题缩进(标题右侧一列)
        col = self.section_level + 1 if self.section_level > 0 else 1
        cell = self.ws.cell(row=self.row, column=col, value=block["text"])
        cell.font = self.default_font
        cell.alignment = self.default_align
        self._track(col)
        self.row += 1

    def render_quote(self, block):
        col = self.section_level + 1 if self.section_level > 0 else 1
        cell = self.ws.cell(row=self.row, column=col, value="“ " + block["text"])
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, italic=True, color="595959")
        cell.fill = PatternFill("solid", fgColor=QUOTE_FILL_COLOR)
        cell.alignment = self.default_align
        self._track(col)
        self.row += 1

    def render_list(self, block):
        base = self.section_level + 1 if self.section_level > 0 else 1
        col = base + block["depth"]
        cell = self.ws.cell(row=self.row, column=col, value=block["bullet"] + " " + block["text"])
        if block.get("task") and block.get("checked"):
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE, strike=True, color="808080")
        else:
            cell.font = self.default_font
        cell.alignment = self.default_align
        self._track(col)
        self.row += 1

    def render_table(self, block):
        base = self.section_level + 1 if self.section_level > 0 else 1
        header = block["header"]
        rows = block["rows"]
        ncols = max([len(header)] + [len(r) for r in rows]) if rows else len(header)

        def write_table_row(data, is_header):
            for ci in range(ncols):
                text = data[ci] if ci < len(data) else ""
                c1 = base + ci * TABLE_COL_SPAN
                c2 = c1 + TABLE_COL_SPAN - 1
                self.ws.merge_cells(start_row=self.row, start_column=c1,
                                    end_row=self.row, end_column=c2)
                cell = self.ws.cell(row=self.row, column=c1, value=text)
                if is_header:
                    cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
                    self._style_range(self.row, c1, self.row, c2,
                                      fill=PatternFill("solid", fgColor=HEADER_FILL_COLOR))
                elif ci == 0:
                    # 首列作为“标签列”：加粗 + 淡蓝底
                    cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
                    self._style_range(self.row, c1, self.row, c2,
                                      fill=PatternFill("solid", fgColor=LABEL_FILL_COLOR))
                else:
                    cell.font = self.default_font
                # 表头居中，其余单元格左对齐；统一不自动换行
                h_align = "center" if is_header else "left"
                cell.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=False)
                self._apply_border(self.row, c1, self.row, c2)
                self._track(c2)
            self.row += 1

        write_table_row(header, True)
        for r in rows:
            write_table_row(r, False)

    def _apply_border(self, r1, c1, r2, c2):
        thin = Side(style="thin", color="BFBFBF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self.ws.cell(row=r, column=c).border = border

    def render_code(self, block):
        base = self.section_level + 1 if self.section_level > 0 else 1
        c1 = base
        c2 = c1 + CODE_WIDTH_COLS - 1
        r1 = self.row
        # 高度：默认 12 行；内容行数超出时自动扩高
        n_lines = block["content"].count("\n") + 1
        height_rows = max(CODE_HEIGHT_ROWS, n_lines)
        r2 = r1 + height_rows - 1
        self.ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cell = self.ws.cell(row=r1, column=c1, value=block["content"])
        code_font = Font(name=FONT_NAME, size=CODE_FONT_SIZE, color=CODE_FONT_COLOR)
        code_fill = PatternFill("solid", fgColor=CODE_BG_COLOR)
        code_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
        self._style_range(r1, c1, r2, c2, font=code_font, fill=code_fill, align=code_align)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self.code_cells.add((r, c))
        self._track(c2)
        self.row = r2 + 1

    def render_image(self, block):
        col = self.section_level + 1 if self.section_level > 0 else 1
        src = block["src"]
        alt = block.get("alt", "")
        path = src if os.path.isabs(src) else os.path.join(self.base_dir, src)
        is_remote = src.lower().startswith(("http://", "https://"))

        if not _HAS_IMAGE or is_remote or not os.path.isfile(path):
            # 无法嵌入(缺依赖/远程/文件不存在)：退化为文本占位
            reason = "远程图片" if is_remote else ("文件不存在" if not is_remote else "不支持")
            cell = self.ws.cell(row=self.row, column=col,
                                value=f"[图片: {alt or src}  ({reason}: {src})]")
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE, italic=True, color="808080")
            cell.alignment = self.default_align
            self._track(col)
            self.row += 1
            return

        try:
            rows_needed = self._embed_image(path, col, IMG_MAX_WIDTH_PX)
            self._track(col)
            self.row += rows_needed
        except Exception as e:
            cell = self.ws.cell(row=self.row, column=col,
                                value=f"[图片嵌入失败: {alt or src} ({e})]")
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE, italic=True, color="808080")
            cell.alignment = self.default_align
            self._track(col)
            self.row += 1

    def _embed_image(self, path, col, max_width):
        """把 PNG/图片文件嵌入当前行 col 列，按 max_width 等比缩放；返回预留行数。"""
        img = XLImage(path)
        if img.width and img.width > max_width:
            ratio = max_width / float(img.width)
            img.width = int(img.width * ratio)
            img.height = int(img.height * ratio)
        anchor = f"{get_column_letter(col)}{self.row}"
        self.ws.add_image(img, anchor)
        # 依据图片像素高度预留行数(行高 0.8cm ≈ 30px @96dpi)
        row_px = ROW_HEIGHT_CM / 2.54 * 96.0
        rows_needed = max(1, int((img.height or row_px) / row_px) + 1)
        return rows_needed

    def _render_mermaid_to_png(self, content):
        """调用 mermaid-cli 把 mermaid 文本渲染成 PNG，返回图片路径。失败抛异常。"""
        mmdc = self.mmdc_cmd
        if not mmdc:
            raise RuntimeError("未找到 mmdc，请先执行: npm install -g @mermaid-js/mermaid-cli")
        tmpdir = tempfile.mkdtemp(prefix="mmd_")
        in_path = os.path.join(tmpdir, "diagram.mmd")
        out_path = os.path.join(tmpdir, "diagram.png")
        with open(in_path, "w", encoding="utf-8") as f:
            f.write(content)
        cmd = mmdc + ["-i", in_path, "-o", out_path,
                      "-b", MERMAID_BG, "-s", str(MERMAID_SCALE)]
        proc = subprocess.run(cmd, capture_output=True, timeout=180)
        if proc.returncode != 0 or not os.path.isfile(out_path):
            err = (proc.stderr or b"").decode("utf-8", "ignore").strip()
            raise RuntimeError(f"mmdc 渲染失败: {err[:300]}")
        return out_path

    def render_mermaid(self, block):
        content = block["content"]

        # 前置判断：若缺少 Mermaid 所需支持（Pillow 或 mermaid-cli），
        # 则不生成图片，直接按普通代码块处理。
        if not self.mmdc_cmd:
            if not self._mermaid_warned:
                if not _HAS_IMAGE:
                    reason = "未安装 Pillow（pip install Pillow）"
                else:
                    reason = "未找到 mermaid-cli（npm install -g @mermaid-js/mermaid-cli）"
                print(f"[提示] {reason}，Mermaid 图表将按代码块显示。")
                self._mermaid_warned = True
            self.render_code({"type": "code", "lang": "mermaid", "content": content})
            return

        # 具备渲染能力：渲染成 PNG 后嵌入；万一运行时出错，仍回退为代码块。
        col = self.section_level + 1 if self.section_level > 0 else 1
        try:
            png = self._render_mermaid_to_png(content)
            rows_needed = self._embed_image(png, col, MERMAID_MAX_WIDTH_PX)
            self._track(col)
            self.row += rows_needed
        except Exception as e:
            print(f"[警告] Mermaid 渲染失败，已按代码块处理：{e}")
            self.render_code({"type": "code", "lang": "mermaid", "content": content})

    # -- 主流程 --
    def render(self, blocks):
        dispatch = {
            "heading": self.render_heading,
            "paragraph": self.render_paragraph,
            "quote": self.render_quote,
            "list": self.render_list,
            "table": self.render_table,
            "code": self.render_code,
            "image": self.render_image,
            "mermaid": self.render_mermaid,
        }
        for block in blocks:
            dispatch[block["type"]](block)
        self._finalize()

    def _finalize(self):
        last_row = self.row - 1
        # 为所有用到的单元格套用默认样式(跳过代码块与已单独设样式的单元格)
        for r in range(1, last_row + 1):
            for c in range(1, self.max_col + 1):
                if (r, c) in self.code_cells:
                    continue
                cell = self.ws.cell(row=r, column=c)
                if cell.font is None or cell.font.name != FONT_NAME:
                    cell.font = self.default_font
                if cell.alignment is None or cell.alignment.vertical is None:
                    cell.alignment = self.default_align

        # 统一列宽
        for c in range(1, self.max_col + 1):
            self.ws.column_dimensions[get_column_letter(c)].width = DEFAULT_COL_WIDTH
        # 统一行高
        for r in range(1, last_row + 1):
            self.ws.row_dimensions[r].height = DEFAULT_ROW_HEIGHT

    def save(self, path):
        self.wb.save(path)


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------
SAMPLE_MD = """# 项目说明文档

这是一个演示段落，用于展示 **段落** 的层级缩进效果与内联样式清洗。

## 功能列表

- 一级列表项
  - 二级列表项
    - 三级列表项
- 支持有序列表
1. 第一步
2. 第二步

## 任务清单

- [x] 已完成的任务
- [ ] 待办任务
  - [ ] 嵌套待办

## 配图

![示例图片](sample.png)

## 数据表格

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | int | 主键 |
| name | str | 名称字段 |

## 代码示例

```python
def hello(name):
    print(f"Hello, {name}!")


def add(a, b):
    return a + b


def main():
    hello("world")
    for i in range(5):
        print(i, add(i, i))


if __name__ == "__main__":
    main()
```

## 流程图

```mermaid
flowchart TD
    A[开始] --> B{是否已登录?}
    B -->|是| C[进入主页]
    B -->|否| D[跳转登录页]
    D --> E[输入账号密码]
    E --> B
    C --> F[结束]
```

## 引用

> 这是一段引用文字，用于展示引用块的样式效果。

### 三级标题下的内容

普通文本也会随所属标题层级缩进。
"""


COPYRIGHT_BANNER = (
    "============================================================\n"
    "  Markdown -> Excel 转换工具\n"
    "  版权所有 © 李鹏虎 (Li Penghu)  保留所有权利\n"
    "  仅供个人学习与非商业用途；商业使用须经作者书面同意。\n"
    "  作者邮箱: gs_lipenghu@qq.com\n"
    "============================================================"
)


def convert(input_path, output_path):
    with open(input_path, "r", encoding="utf-8") as f:
        md_text = f.read()
    blocks = parse_markdown(md_text)
    renderer = ExcelRenderer(base_dir=os.path.dirname(os.path.abspath(input_path)))
    renderer.render(blocks)
    renderer.save(output_path)
    print(f"完成: {input_path} -> {output_path}  (共 {len(blocks)} 个块)")


def main():
    print(COPYRIGHT_BANNER)
    args = sys.argv[1:]
    if not args:
        # 无参数：生成内置示例
        demo_md = "demo.md"
        demo_xlsx = "demo.xlsx"
        with open(demo_md, "w", encoding="utf-8") as f:
            f.write(SAMPLE_MD)
        convert(demo_md, demo_xlsx)
        return

    input_path = args[0]
    if len(args) >= 2:
        output_path = args[1]
    else:
        output_path = os.path.splitext(input_path)[0] + ".xlsx"
    convert(input_path, output_path)


if __name__ == "__main__":
    main()
