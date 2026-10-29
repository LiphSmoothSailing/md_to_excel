---
name: excel-to-md-plain
description: >-
  普通的 Excel 转 Markdown（与 md-to-excel-plain 互逆）：把 Excel 里有细边框的单元格区域
  还原成 GFM 管道表格，A 列无边框的文本行原样写回 Markdown。适用于"excel 转 md"、
  "xlsx 转 markdown"、"把 md-to-excel-plain 产出的 xlsx 还原成 md"等场景。
---

# 普通 Excel → Markdown

与 `md-to-excel-plain` 互逆：

| md-to-excel-plain | excel-to-md-plain（本技能） |
| --- | --- |
| GFM 管道表格 → 带边框 Excel 单元格 | 带边框 Excel 区域 → GFM 管道表格 |
| 其余 Markdown 源码 → A 列纯文本 | A 列纯文本 → Markdown 文本行 |
| 标题行 `#` 原样（不做标题样式） | 标题行 `#` 原样保留 |
| `**bold**` 真粗体（--strip-bold 模式） | 粗体 run → 包回 `**bold**`；表头/标题行的结构性加粗不包 |
| `--merge-layout` 合并单元格 | 合并区还原：锚点列为准、非锚点格跳过，多出的空列不进表格 |

## 输出规则

1. **表格识别**：A 列有细左边框（`border.left.style == "thin"`）的连续行视为一张表。
   - 第一行 = 表头，后续行 = 数据行
   - 列数以表头行的逻辑列数为准，数据行多出的单元格不进表格（md-to-excel-plain 产出的 xlsx 不会出现这种情况）
   - 分隔行对齐标记（`:---` / `:---:` / `---:`）根据表头和数据单元格的水平对齐反推
   - 单元格里的字面 `|` 转义为 `\|`；单元格内换行还原为 `<br>`
2. **非表格行**：A 列值原样写为一行 Markdown；粗体 run 包成 `**...**`。
   - None 值 → 空行
   - 标题行（`#` 开头）的整行加粗是结构性的，不包 `**`
   - 表格表头行的加粗也是结构性的，不包 `**`
3. **合并格还原**：遇到 `--merge-layout` 产出的合并单元格，以表头行的合并锚点列为准，跳过非锚点格，多出的空列不进表格（竖向合并不处理）。
4. **不做的事**：不讀字體顏色/填充/列寬/行高；不處理豎向合併；多 sheet 只讀第一個 active sheet。

## 用法

```powershell
python .agents\skills\excel-to-md-plain\scripts\excel_to_md_plain.py --input <源.xlsx> [--output <目标.md>] [--force]
```

| 参数 | 说明 |
| --- | --- |
| `--input` | 输入 xlsx（必需） |
| `--output` | 输出 md（缺省 `./output/<源文件名>.md`） |
| `--force` | 输出已存在时允许覆盖（缺省拒绝） |

## 工作流

1. 运行脚本
2. 人工对照：非表格文本行应与源 Markdown 一致，表格数一致（脚本不做自动校验）
3. `present_files` 交付 md 文件

## 目录

| 用途 | 位置 |
| --- | --- |
| 脚本 | `scripts/excel_to_md_plain.py` |
| 自测 | `tests/test_excel_to_md_plain.py` |

## 依赖

`openpyxl>=3.1`。
