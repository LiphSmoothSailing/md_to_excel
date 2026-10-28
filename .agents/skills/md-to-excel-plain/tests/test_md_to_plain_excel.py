# -*- coding: utf-8 -*-
r"""md-to-excel-plain 的验收测试：解析边界情况 + 「源码保真」回读 + 真实文档冒烟。

运行（在本项目根目录下）：
    python .agents\skills\md-to-excel-plain\tests\test_md_to_plain_excel.py [真实md路径]
退出码 0 = 全部通过。中间产物写在项目 `temp/` 下。
第 4 节是可选的：给了参数就用该文件，没给就从下面的 REAL_DIR 里自动挑一个，
都没有就 SKIP（便于技能被复制到别的项目）。
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve()
SKILL_DIR = HERE.parents[1]                     # .../md-to-excel-plain
SCRIPTS = SKILL_DIR / "scripts"
sys.path.insert(0, str(SCRIPTS))
import md_to_plain_excel as M                   # noqa: E402

# 项目根（.agents/skills/<skill>/tests/ → 上溯 4 级）；结构不同就退到技能目录
try:
    ROOT = HERE.parents[4]
    if not (ROOT / "output").exists():
        ROOT = SKILL_DIR
except IndexError:
    ROOT = SKILL_DIR
TEMP = ROOT / "temp"
TEMP.mkdir(parents=True, exist_ok=True)

REAL_DIR = Path(r"C:\Goonet\Code\goo-net\docs\batch_pgsq_to_mysql")

SRC = "\n".join([
    "# 見出し",
    "",
    "段落文字，含 **粗体** 與 `行內代碼`。",
    "",
    "| 欄位 | 說明 |",
    "|:---|:---:|",
    "| a | 1 |",
    "| b | 2 |",
    "",
    "## 表格緊接標題（前面無空行）",
    "| NO | A | B | C | D | 說明 |",
    "|---:|---:|---|---|---|---|",
    "| 1 | x | `p\\|q` | | `COUNT(CASE...)` | 含 \\| 轉義 |",
    "| 2 | y | z | | | |",
    "",
    "- 清單項目 1",
    "- 清單項目 2",
    "",
    "```sql",
    "SELECT a|b FROM t;  /* 管道符不應被當成表格 */",
    "| 這行也不能當表格 |",
    "```",
    "",
    "```mermaid",
    "graph LR",
    "  A --> B",
    "```",
    "",
    "| 只有表頭 | 沒有資料 |",
    "|---|---|",
    "",
    "| 單欄 |",
    "|:---:|",
    "| 單欄值1 |",
    "",
    "---",
    "",
    "> 引用文字 `code`",
    "",
    "最後一行",
])
SRC_PATH = TEMP / "_plain_test.md"
SRC_PATH.write_text(SRC, encoding="utf-8")
OUT_PATH = TEMP / "_plain_test.xlsx"

fails = []


def check(name, cond, extra=""):
    print(("PASS  " if cond else "FAIL  ") + name + ("" if cond else f"   <- {extra}"))
    if not cond:
        fails.append(name)


def table_syntax_line_ids(src_lines):
    """独立（另写一遍）判定源文件里哪些行属于表格语法，只用于比对"""
    ids, i, n, fence = set(), 0, len(src_lines), None
    while i < n:
        ln = src_lines[i]
        st = ln.strip()
        if fence is None and (st.startswith("```") or st.startswith("~~~")):
            fence = st[0]
            i += 1
            continue
        if fence is not None:
            if st.startswith(fence * 3):
                fence = None
            i += 1
            continue
        if "|" in ln and i + 1 < n and "-" in src_lines[i + 1] and "|" in src_lines[i + 1] \
                and all(c in " -:|" for c in src_lines[i + 1]) and src_lines[i + 1].strip():
            ids.add(i)          # 表头行
            ids.add(i + 1)      # 分隔行
            j = i + 2
            while j < n and src_lines[j].strip() and "|" in src_lines[j]:
                ids.add(j)
                j += 1
            i = j
            continue
        i += 1
    return ids


def fidelity(src_lines, ws):
    """A 列（非表格行）逐行回读 == 源文件非表格行（顺序、内容完全一致）

    用「A 列有没有细边框」独立判定该行是不是表格行——不依赖脚本自己的解析结果。
    """
    plain_rows = [r for r in range(1, ws.max_row + 1)
                  if ws.cell(row=r, column=1).border.left.style is None]
    got = [M.rich_text_to_str(ws.cell(row=r, column=1).value) for r in plain_rows]
    skip = table_syntax_line_ids(src_lines)
    exp = [ln if ln != "" else None for k, ln in enumerate(src_lines) if k not in skip]
    if got == exp:
        return True, f"{len(got)} 行一致"
    diff = next(((k, a, b) for k, (a, b) in enumerate(zip(got, exp)) if a != b), None)
    return False, f"len {len(got)} vs {len(exp)}; first diff={diff}; got尾3={got[-3:]}"


# ---------------------------------------------------------------- 1) 解析层
lines = SRC.split("\n")
if lines and lines[-1] == "":
    lines.pop()
blocks = M.parse(lines)
tables = [b for b in blocks if b[0] == "table"]
texts = [b[1] for b in blocks if b[0] == "text"]

check("表格张数 = 4", len(tables) == 4, f"got {len(tables)}")
check("无表格被误吞（分隔线 --- 不算表格）", "---" in texts)
check("代码围栏内的 | 行仍是纯文本", "| 這行也不能當表格 |" in texts)

t1 = tables[0]
check("表1 表头", t1[1] == ["欄位", "說明"], t1[1])
check("表1 对齐 left/center", t1[3] == ["left", "center"], t1[3])
check("表1 数据", t1[2] == [["a", "1"], ["b", "2"]], t1[2])

t2 = tables[1]
check("表2 列数 = 6", len(t2[1]) == 6, len(t2[1]))
check("表2 末列无对齐标记 → None", t2[3][5] is None and t2[3][0] == "right", t2[3])
check("表2 转义 \\| 还原为 |", t2[2][0][2] == "`p|q`", t2[2][0][2])
check("表2 补齐空单元格", t2[2][1] == ["2", "y", "z", "", "", ""], t2[2][1])

t3 = tables[2]
check("只有表头的表：0 数据行", t3[2] == [] and t3[1] == ["只有表頭", "沒有資料"], t3)

t4 = tables[3]
check("单列表格", t4[1] == ["單欄"] and t4[2] == [["單欄值1"]] and t4[3] == ["center"], t4)

# ---------------------------------------------------------------- 2) 渲染层
ws = M.build(blocks, "宋体", 11, "keep", False, 5, 60, "test")
ws.parent.save(OUT_PATH)

colA = [M.rich_text_to_str(ws.cell(row=r, column=1).value) for r in range(1, ws.max_row + 1)]

check("表格行数 = 表数 + 数据行数",
      sum(len(b[2]) + 1 for b in tables) ==
      len([r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).border.left.style == "thin"]))
ok, detail = fidelity(lines, ws)
check("A 列（非表格行）逐行回读 == 源文件非表格行（顺序、内容完全一致）", ok, detail)

check("标题行保留 # 原样", "# 見出し" in colA)
check("加粗标记未被剥掉", any(isinstance(v, str) and "**粗体**" in v for v in colA))
check("行内反引号未被剥掉", any(isinstance(v, str) and "`行內代碼`" in v for v in colA))
check("列表标记未被剥掉", "- 清單項目 1" in colA)
check("引用标记未被剥掉", any(isinstance(v, str) and v.startswith("> ") for v in colA))
check("代码围栏行本身原样保留", colA.count("```sql") == 1 and colA.count("```") == 2)
check("主题分隔线 --- 原样保留", "---" in colA)

# 新增：标题行加粗、**段红色
heading_row = next(r for r in range(1, ws.max_row + 1)
                   if isinstance(ws.cell(row=r, column=1).value, str)
                   and ws.cell(row=r, column=1).value.startswith("# "))
check("标题行(#)整行加粗", ws.cell(row=heading_row, column=1).font.bold is True)
bold_line_row = next(r for r in range(1, ws.max_row + 1)
                     if M.rich_text_to_str(ws.cell(row=r, column=1).value) == "段落文字，含 **粗体** 與 `行內代碼`。")
bold_val = ws.cell(row=bold_line_row, column=1).value
check("含**的行用CellRichText渲染", hasattr(bold_val, "__iter__") and not isinstance(bold_val, str))


t1_row = next(r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "欄位")
t2_row = next(r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "NO")
check("表1 表头加粗", ws.cell(row=t1_row, column=1).font.bold is True)
check("表1 第1列表头跟随 :--- → 左对齐", ws.cell(row=t1_row, column=1).alignment.horizontal == "left")
check("表2 无对齐标记的列表头默认居中", ws.cell(row=t2_row, column=3).alignment.horizontal == "center")
check("表1 数据列按 :--- 左对齐", ws.cell(row=t1_row + 1, column=1).alignment.horizontal == "left")
check("表1 2 列有细边框",
      all(ws.cell(row=t1_row + 1, column=c).border.left.style == "thin" for c in (1, 2)))
check("非表格行不换行（可溢出显示完整）", ws.cell(row=1, column=1).alignment.wrap_text is False)
check("非表格行无边框", ws.cell(row=1, column=1).border.left.style is None)
check("非表格行右侧全空（Excel 才会溢出显示完整文本）",
      all(ws.cell(row=r, column=c).value is None
          for r in range(1, ws.max_row + 1)
          if ws.cell(row=r, column=1).border.left.style is None
          for c in range(2, ws.max_column + 1)))

w1 = ws.column_dimensions["A"].width
check("列宽已写入且在 [5,60] 内", w1 is not None and 5 <= w1 <= 60, w1)

# ---------------------------------------------------------------- 3) --excel-table
ws2 = M.build(blocks, "宋体", 11, "keep", True, 5, 60, "test2")
names = list(ws2.tables)
check("--excel-table：3 张有数据的表都注册了 ListObject（只有表头的跳过）", len(names) == 3, names)
check("ListObject 名称连续且唯一", names == ["Tbl1", "Tbl2", "Tbl3"], names)

# ---------------------------------------------------------------- 4) 真实文件冒烟
real = None
if len(sys.argv) > 1:
    real = Path(sys.argv[1])
elif REAL_DIR.is_dir():
    cands = sorted(REAL_DIR.glob("*.merge*.md")) or sorted(REAL_DIR.glob("*.md"))
    real = cands[0] if cands else None

if real and real.is_file():
    rl = real.read_text(encoding="utf-8").split("\n")
    if rl and rl[-1] == "":
        rl.pop()
    rb = M.parse(rl)
    rt = [b for b in rb if b[0] == "table"]
    check(f"真实文件（{real.name}）：检出表格 > 0", len(rt) > 0, f"tables={len(rt)}")
    rws = M.build(rb, "宋体", 11, "keep", False, 5, 60, "real")
    ok, detail = fidelity(rl, rws)
    check("真实文件：A 列纯文本逐行 == 源文件非表格行（原文照搬）", ok, detail)
    # 表格内容与源文件表格语法行对照（分隔行不属于内容行）
    # 用脚本自己的 split_row（正确处理 \| 转义），并把每行补齐到 xlsx 全局列数
    import re as _re
    src_cells_raw = []
    for k, ln in enumerate(rl):
        if k not in table_syntax_line_ids(rl) or not ln.strip():
            continue
        cells = M.split_row(ln)
        if all(_re.fullmatch(r":?-+:?", c or "") for c in cells):
            continue                     # `|:---|---:|` 分隔行
        src_cells_raw.append(cells)
    xls_cells = [[rws.cell(row=r, column=c).value for c in range(1, rws.max_column + 1)]
                 for r in range(1, rws.max_row + 1)
                 if rws.cell(row=r, column=1).border.left.style == "thin"]
    ncol_xls = len(xls_cells[0]) if xls_cells else 0
    src_cells = [(r + [""] * ncol_xls)[:ncol_xls] for r in src_cells_raw]
    n_diff = sum(1 for a, b in zip(src_cells, xls_cells)
                 if [x if x is not None else "" for x in b] != a)
    check("真实文件：表格行数 = 源表格内容行数（表头+数据）", len(src_cells) == len(xls_cells),
          f"{len(src_cells)} vs {len(xls_cells)}")
    check("真实文件：每一表格行逐格与源文件一致（不丢列、不串格）", n_diff == 0, f"{n_diff} 行不一致")
else:
    print(f"SKIP  真实文件冒烟（未找到可用的 md：{real}）")

print()
print("FAILED:", fails if fails else "none")
sys.exit(1 if fails else 0)
