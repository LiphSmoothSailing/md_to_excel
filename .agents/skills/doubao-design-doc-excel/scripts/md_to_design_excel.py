# -*- coding: utf-8 -*-
"""MD → 設計書 Excel 生成器（基于 design_template.xlsx 版式）。

两种输入模式：
  1) --input xxx.md     解析 Markdown（保留原文语言）渲染
  2) --blocks xxx.json  直接渲染 blocks JSON（适合 agent 预先翻译/组织内容）

blocks JSON 结构（列表）：
  ["chapter", "标题文本"]                      章节标题（深灰底白字）
  ["section", "标题文本"]                      子节标题（浅灰底）
  ["table", ["列1","列2"], [["a","b"], ...]]   表格（1~3 列；>3 列自动并入第 3 列）
  ["note", "文本"]                             正文段落（可含 \n 多行）
  ["code", "文本"]                             代码/JSON 示例（浅灰底 Consolas）

模板版式铁律（不可违反）：
  - 保留模板前 6 行（值/样式/合并/行高）
  - 列宽与模板完全一致（A=3.25、AG=4.5、AH=3.25，其余默认 12.63），绝不改列宽
  - 长内容一律用合并单元格承载
"""
import argparse
import json
import math
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MMDC = shutil.which('mmdc')

FONT_JA = 'Meiryo'
FONT_MONO = 'Consolas'
LINE_THIN = Side(style='thin', color='FF000000')

# 内容区列段（1-based 列号），基于模板 B..BC 列宽体系
SEG3 = [(2, 5), (6, 29), (30, 55)]   # B:E | F:AC | AD:BC
SEG2 = [(2, 8), (9, 55)]             # B:H | I:BC
SEG1 = [(2, 55)]                     # B:BC
FULL = (2, 55)

DEFAULT_SHEET = '設計書'


def disp_width(s):
    return sum(2 if unicodedata.east_asian_width(ch) in ('F', 'W') else 1
               for ch in str(s or ''))


def est_lines(text, c1, c2):
    text = str(text or '')
    if not text:
        return 1
    avail = max(sum(12.63 for _ in range(c1, c2 + 1)) / 2 - 2, 8)
    return sum(max(1, math.ceil(disp_width(ln) / avail)) for ln in text.split('\n'))


def col_align(values):
    if not values:
        return 'left'
    maxw = max(disp_width(str(v or '')) for v in values)
    return 'center' if maxw <= 10 else 'left'


# ---------------- MD 解析 ----------------
def _split_row(r):
    r = r.strip()
    if r.startswith('|'):
        r = r[1:]
    if r.endswith('|'):
        r = r[:-1]
    return [c.strip() for c in r.split('|')]


def _is_sep_row(r):
    return bool(re.match(r'^\s*\|?[\s:\-|]+\|?\s*$', r)) and '-' in r


def _clean_inline(text):
    """去除行内代码反引号标记（markdown 渲染语义），保留其余原文"""
    return re.sub(r'`([^`]*)`', r'\1', str(text))


def _png_size(path):
    """读取 PNG 像素尺寸（读 IHDR，不依赖 Pillow）"""
    with open(path, 'rb') as f:
        f.read(16)
        w, h = struct.unpack('>II', f.read(8))
    return w, h


def _render_mermaid(code, out_png, width=600):
    """用 mermaid-cli 渲染 mermaid 代码为 PNG；失败返回 False（调用方降级为代码文本）"""
    if not MMDC:
        return False
    mmd = out_png + '.mmd'
    with open(mmd, 'w', encoding='utf-8') as f:
        f.write(code)
    try:
        cmd = f'"{MMDC}" -i "{mmd}" -o "{out_png}" -b white -w {width}'
        r = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        return r.returncode == 0 and os.path.exists(out_png)
    finally:
        if os.path.exists(mmd):
            os.remove(mmd)


def parse_md(text):
    lines = text.split('\n')
    blocks = []
    doc_title = None
    first_level = None
    i, n = 0, len(lines)
    while i < n:
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level, title = len(m.group(1)), _clean_inline(m.group(2).strip())
            if first_level is None:
                first_level = level
                doc_title = title
            elif level == first_level + 1:
                blocks.append(('chapter', title))
            elif level >= first_level + 2:
                blocks.append(('section', title))
            else:  # 同级后续标题按章节处理
                blocks.append(('chapter', title))
            i += 1
            continue
        if line.lstrip().startswith('```'):
            fence = line.strip()[3:].strip().lower()
            i += 1
            code = []
            while i < n and not lines[i].strip().startswith('```'):
                code.append(lines[i])
                i += 1
            i += 1  # 跳过结束 ```
            body = '\n'.join(code).rstrip()
            blocks.append(('mermaid', body) if fence == 'mermaid' else ('code', body))
            continue
        if line.strip().startswith('|'):
            tbl = []
            while i < n and lines[i].strip().startswith('|'):
                tbl.append(lines[i])
                i += 1
            blocks.append(parse_table(tbl))
            continue
        para = []
        while i < n:
            l = lines[i].rstrip()
            if not l.strip() or l.strip().startswith('|') or l.strip().startswith('```') \
                    or re.match(r'^(#{1,6})\s+', l):
                break
            para.append(_clean_inline(l.strip()))
            i += 1
        if para:
            blocks.append(('note', '\n'.join(para)))
    return doc_title, blocks


def parse_table(rows):
    headers = [_clean_inline(h) for h in _split_row(rows[0])]
    data = []
    for r in rows[1:]:
        if _is_sep_row(r):
            continue
        data.append([_clean_inline(c) for c in _split_row(r)])
    maxc = max([len(headers)] + [len(r) for r in data])
    headers = headers + [''] * (maxc - len(headers))
    data = [r + [''] * (maxc - len(r)) for r in data]
    return ('table', headers, data)


def normalize_table(headers, rows):
    """>3 列时并入第 3 列（表头/内容用 ' / ' 连接），返回 (headers, rows)"""
    n = len(headers)
    if n <= 3:
        return headers, rows
    h3 = ' / '.join(headers[2:])
    headers = headers[:2] + [h3]
    out = []
    for r in rows:
        v3 = ' / '.join(str(x) for x in r[2:] if str(x) != '')
        out.append(r[:2] + [v3])
    return headers, out


# ---------------- 渲染 ----------------
def build(template_path, out_path, blocks, doc_title=None, sheet_name=None):
    wb = load_workbook(template_path)
    ws = wb.active
    if sheet_name is None:
        sheet_name = _make_sheet_name(doc_title)
    ws.title = sheet_name

    # 清理模板内容区（第 7 行起），保留前 6 行与列宽
    if ws.max_row > 6:
        ws.delete_rows(7, ws.max_row - 6)
    ws.merged_cells.ranges = [mr for mr in ws.merged_cells.ranges if mr.max_row <= 6]
    for row in ws.iter_rows(min_row=1, max_row=6):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('=#REF!'):
                cell.value = None
    ws.sheet_state = 'visible'
    for s in list(wb.sheetnames):
        if s != sheet_name:
            del wb[s]
    wb.active = 0

    if doc_title:
        ws['A1'] = doc_title

    header_fill = PatternFill('solid', fgColor='FF666666')
    section_fill = PatternFill('solid', fgColor='FFD9D9D9')
    code_fill = PatternFill('solid', fgColor='FFF2F2F2')
    img_dir = tempfile.mkdtemp(prefix='design_mermaid_')

    def set_merged(r, c1, c2, value, fill, font, align, bordered):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if fill is not None:
                cell.fill = fill
            cell.font = font
            cell.alignment = align
            cell.border = (Border(left=LINE_THIN, right=LINE_THIN,
                                  top=LINE_THIN, bottom=LINE_THIN) if bordered else Border())
        if c2 > c1:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        ws.cell(row=r, column=c1).value = value

    row = 7
    for block in blocks:
        kind = block[0]
        if kind == 'chapter':
            set_merged(row, FULL[0], FULL[1], block[1], header_fill,
                       Font(name=FONT_JA, size=11, bold=True, color='FFFFFFFF'),
                       Alignment(horizontal='left', vertical='center'), bordered=False)
            ws.row_dimensions[row].height = 24
        elif kind == 'section':
            set_merged(row, FULL[0], FULL[1], block[1], section_fill,
                       Font(name=FONT_JA, size=10, bold=True, color='FF000000'),
                       Alignment(horizontal='left', vertical='center'), bordered=False)
            ws.row_dimensions[row].height = 22
        elif kind == 'note':
            set_merged(row, FULL[0], FULL[1], block[1], None,
                       Font(name=FONT_JA, size=10, color='FF000000'),
                       Alignment(horizontal='left', vertical='center', wrap_text=True),
                       bordered=False)
            ws.row_dimensions[row].height = max(18, est_lines(block[1], *FULL) * 14 + 8)
        elif kind == 'code':
            set_merged(row, FULL[0], FULL[1], block[1], code_fill,
                       Font(name=FONT_MONO, size=10, color='FF333333'),
                       Alignment(horizontal='left', vertical='center', wrap_text=True),
                       bordered=False)
            ws.row_dimensions[row].height = est_lines(block[1], *FULL) * 13 + 8
        elif kind == 'mermaid':
            png = os.path.join(img_dir, f'mermaid_{row}.png')
            if _render_mermaid(block[1], png):
                w, h = _png_size(png)
                max_h = 520  # Excel 行高上限约 409.5pt ≈ 546px，保守取 520
                scale = min(1.0, max_h / h)
                img = XLImage(png)
                img.width = int(w * scale)
                img.height = int(h * scale)
                ws.add_image(img, f'B{row}')
                ws.row_dimensions[row].height = img.height * 0.75 + 4
            else:  # mmdc 不可用或渲染失败 → 降级为代码文本
                set_merged(row, FULL[0], FULL[1], block[1], code_fill,
                           Font(name=FONT_MONO, size=10, color='FF333333'),
                           Alignment(horizontal='left', vertical='center', wrap_text=True),
                           bordered=False)
                ws.row_dimensions[row].height = est_lines(block[1], *FULL) * 13 + 8
        elif kind == 'table':
            headers, rows = normalize_table(block[1], block[2])
            n = len(headers)
            segs = SEG1 if n == 1 else SEG2 if n == 2 else SEG3
            for j, h in enumerate(headers):
                c1, c2 = segs[j]
                set_merged(row, c1, c2, h, header_fill,
                           Font(name=FONT_JA, size=10, bold=True, color='FFFFFFFF'),
                           Alignment(horizontal='center', vertical='center', wrap_text=True),
                           bordered=True)
            ws.row_dimensions[row].height = 22
            row += 1
            for r in rows:
                for j, v in enumerate(r):
                    c1, c2 = segs[j]
                    set_merged(row, c1, c2, v, None,
                               Font(name=FONT_JA, size=9, color='FF000000'),
                               Alignment(horizontal=col_align([rr[j] for rr in rows]),
                                         vertical='center', wrap_text=True),
                               bordered=True)
                    ws.row_dimensions[row].height = max(
                        ws.row_dimensions[row].height or 0, est_lines(v, c1, c2) * 14 + 6)
                row += 1
        row += 1  # 块间空行

    wb.save(out_path)
    shutil.rmtree(img_dir, ignore_errors=True)
    return ws.max_row


def _make_sheet_name(title):
    name = (title or DEFAULT_SHEET).strip()
    name = re.sub(r'[\\/*?:\[\]]', '', name).replace('：', '_').replace('/', '_')
    return (name or DEFAULT_SHEET)[:31]


def load_blocks_json(path):
    data = json.loads(Path(path).read_text(encoding='utf-8'))
    out = []
    for item in data:
        kind = item[0]
        if kind == 'table':
            out.append((kind, list(item[1]), [list(r) for r in item[2]]))
        else:
            out.append((kind, item[1]))
    return out


def main():
    ap = argparse.ArgumentParser(description='MD → 設計書 Excel（design_template 版式）')
    ap.add_argument('--input', help='输入 Markdown 文件路径')
    ap.add_argument('--blocks', help='blocks JSON 文件路径（与 --input 二选一）')
    ap.add_argument('--output', help='输出 xlsx 路径（默认：与输入同目录同名 .xlsx）')
    ap.add_argument('--template', help='模板 xlsx 路径（默认：本 skill assets/design_template.xlsx）')
    ap.add_argument('--sheet-name', help='sheet 名（默认取文档标题）')
    ap.add_argument('--title', help='A1 文档标题（默认：MD 模式取首个 # 标题，blocks 模式保留模板标题）')
    args = ap.parse_args()

    if not args.input and not args.blocks:
        ap.error('必须提供 --input 或 --blocks')
    if args.template:
        tpl = Path(args.template)
    else:
        tpl = Path(__file__).resolve().parent.parent / 'assets' / 'design_template.xlsx'
    if not tpl.exists():
        sys.exit(f'模板不存在: {tpl}')

    if args.blocks:
        blocks = load_blocks_json(args.blocks)
        out = Path(args.output) if args.output else Path(args.blocks).with_suffix('.xlsx')
        doc_title = args.title  # None 则保留模板 A1 标题
    else:
        src = Path(args.input)
        md_title, blocks = parse_md(src.read_text(encoding='utf-8'))
        out = Path(args.output) if args.output else src.with_suffix('.xlsx')
        doc_title = args.title or md_title
        if doc_title is None and blocks and blocks[0][0] == 'chapter':
            doc_title = blocks[0][1]

    rows = build(str(tpl), str(out), blocks, doc_title=doc_title, sheet_name=args.sheet_name)
    print(f'saved: {out} ({rows} rows, sheet={Path(out).stem})')


if __name__ == '__main__':
    main()
