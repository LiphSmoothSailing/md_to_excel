---
name: md-to-excel-plain
description: >-
  普通的 Markdown 转 Excel（通用 / 轻量版）：只把 Markdown 里的管道表格转成 Excel 表格，
  其余内容（标题、段落、列表、引用、代码块、分隔线等）一律以 Markdown 源码纯文本原样写进单元格，
  不做标题样式、不做配色、不做模板版式。适用于"markdown 转 excel"、"md 转表格"、"只把表格转成
  excel"、"其余保持 markdown 源码"、"普通/通用 markdown to excel"等任务。
  带固定设计书版式（前 6 行头部、列段合并、配色体系）的转换请改用 doubao-design-doc-excel，二者互不影响。
---

# 普通 Markdown → Excel

一句话需求：**表格变成 Excel 表格，其余全部按 Markdown 源码原文显示。**

## 与 doubao-design-doc-excel 的分工（重要）

| | 本技能 `md-to-excel-plain` | `doubao-design-doc-excel` |
| --- | --- | --- |
| 模板 | **不用任何模板**，输出干净的新工作簿 | 强制使用 `assets/design_template.xlsx` 版式 |
| 表格 | 管道表格 → Excel 表格 | 管道表格 → 模板列段合并的表格 |
| 非表格内容 | **Markdown 源码原文照搬**，无样式 | 渲染成章节 / 子节 / 正文 / 代码样式 |
| mermaid | 不做处理，保留源码文本 | 渲染为图片嵌入 |

**本技能绝不改动 `doubao-design-doc-excel`（不动它的 SKILL.md / 脚本 / 模板）。**

## 输出规则（用户强约束，不可违反）

1. **只有表格被"转换"**：识别 GFM 管道表格（表头行 + `|---|:--:|---:|` 分隔行 + 数据行），
   写成真正的 Excel 单元格表格（表头 + 数据 + 细边框）。
   - 分隔行里的 `:---` / `:---:` / `---:` 对齐标记会作为该列的对齐方式生效，分隔行本身不写进单元格。
   - 单元格文字**原样保留**（不剥掉 `` ` ``、`**`、`\|` 之外的转义也不动）；`\|` 还原成 `|`。
   - 表格列数不一致时按最宽的行补齐空单元格，不丢内容。
2. **其余内容一律原样显示**：标题行保留 `#`、列表保留 `-` / `1.`、加粗保留 `**`、行内代码保留反引号、
   代码块保留 ``` 围栏、分隔线保留 `---`、空行保留空行。
   **样式例外（仅作用于 A 列非表格行）**：
   - 标题行（`#` / `##` / `###`…）整行**加粗**，`#` 字符保留；
   - 行内 `**...**` 段渲染为**红色**，`**` 字符保留；标题行里的 `**` 段红且粗；
   - 表格单元格内的 `#`、`**` 一律不做任何处理。
   **不加粗（除标题行外）、不配色（除 ** 段红色外）、不合并单元格、不改写任何字符。**
3. **不加换行、让文字自然溢出**：非表格行的文字只写到 A 列且 `wrap_text=False`，
   Excel 会自动溢出到右侧空白单元格，因此即使 A 列很窄也能完整看到（不会被截断）。
4. **表格单元格自动换行**：列宽按内容计算（下限 5、上限 60 字符），超出部分折行，行高按内容撑开。
   同一列被多张表共用时取所有表的最大需求。
5. **一张 sheet 承载全文**，按源码顺序自上而下排列（表格就地展开成多列）。

## 用法

```powershell
python .agents\skills\md-to-excel-plain\scripts\md_to_plain_excel.py --input <源.md> [--output <目标.xlsx>]
```

脚本路径若在别的项目，用该技能目录下的 `scripts/md_to_plain_excel.py`。

| 参数 | 说明 |
| --- | --- |
| `--input` | 输入 Markdown 路径（UTF-8，必需） |
| `--output` | 输出 xlsx（缺省 `./output/<源文件名>.xlsx`，目录自动创建） |
| `--sheet-name` | sheet 名（缺省取源文件名，超 31 字符截断） |
| `--font` / `--font-size` | 字体名 / 字号（缺省 `宋体` / `11`） |
| `--blank-lines` | `keep`（缺省，源里几个空行就留几个）/ `collapse`（连续空行并成 1 行）/ `drop` |
| `--excel-table` | 额外把每张表注册为真正的 Excel 表格对象（ListObject，带筛选）；表头为空或重名时会被规范化成 `Column2` / `名称_2` 并同步回写单元格 |
| `--min-col-width` / `--max-col-width` | 表格列宽上下限（缺省 5 / 60） |
| `--verify` | 生成后自动回读 xlsx 做保真校验：A 列纯文本逐行一致 + 表格行数一致；失败时退出码 1 |
| `--force` | 输出文件已存在时允许覆盖（**缺省拒绝**，见下） |

> **注意默认输出与同名冲突**：本技能默认输出 `./output/<源文件名>.xlsx`，与
> `doubao-design-doc-excel` 的默认输出**同名**。为免误盖别处已交付的产物，脚本在目标文件
> 已存在时会直接报错退出、不覆盖。确认要覆盖加 `--force`；想两个产物并存就显式指定
> `--output ...<源文件名>.plain.xlsx`。**转换前先确认目标文件是不是自己该覆盖的那个。**

## 工作流

1. **先看源文件**，确认是"普通转换"需求（没有指定设计书模板/版式）。
   若用户要求设计书版式，改用 `doubao-design-doc-excel`，不要用本技能硬套。
2. 运行脚本，**建议加 `--verify`**（默认输出到 `./output/<源文件名>.xlsx`）：
   ```powershell
   python .agents\skills\md-to-excel-plain\scripts\md_to_plain_excel.py \
     --input <源.md> --verify [--force]
   ```
3. **验证**（`--verify` 已自动完成，退出码 0 = 通过）：
   - A 列纯文本逐行 == 源文件非表格行（顺序、内容完全一致）
   - 表格行数（表头+数据）与 parse 结果一致
   - 如需更严格的逐格对照（含 `\|` 转义、列数补齐），跑自测脚本见下。
4. **交付**：`present_files` 交付 xlsx 本地路径，并说明表格数、源码行数、sheet 名。

## 自测（改过脚本后必跑）

```powershell
python .agents\skills\md-to-excel-plain\tests\test_md_to_plain_excel.py [真实md路径]
```

36 项断言，退出码 0 = 全通过。覆盖：解析边界（代码围栏里的 `|`、`---` 分隔线、只有表头的表、
单栏表、`\|` 转义还原、列数不一致补齐）、表头/数据对齐与边框、**非表格行右侧必须全空**、
以及最关键的「A 列纯文本逐行 == 源文件非表格行」。给了真实 md（或从
`C:\Goonet\Code\goo-net\docs\batch_pgsq_to_mysql` 自动挑一个）时，还会做全部表格**逐格对照**。

## 目录约定

| 用途 | 位置 | 覆盖参数 |
| --- | --- | --- |
| 脚本 | 本技能 `scripts/md_to_plain_excel.py` | — |
| 自测脚本 | 本技能 `tests/test_md_to_plain_excel.py` | 可选参数：真实 md 路径 |
| 转换产物 | `./output/<源文件名>.xlsx`（自动创建） | `--output` |
| 临时文件 / 验证脚本 | `./temp/` | — |

## 依赖

`openpyxl>=3.1`（生成 `.xlsx`）。不依赖模板、不依赖 Node.js / mermaid。
