---
name: doubao-design-doc-excel-to-md
description: >-
  把 doubao-design-doc-excel 生成的设计书 Excel（design_template 版式）还原为 Markdown，与 doubao-design-doc-excel 互逆。
  适用于"设计书 excel 转 md"、"把 design_template 版式的 xlsx 还原成 markdown"、"設計書 Excel を Markdown に戻す"、
  "excel 设计书转 markdown"等任务。只处理按该版式生成的文件；普通 Excel 请用 excel-to-md-plain。
---

# 設計書 Excel → Markdown（design_template 版式）

与 `doubao-design-doc-excel` 互逆：按版式样式识别章节、子节、表格、代码、正文与图片，还原为 Markdown。

## 版式识别规则（与生成端 `md_to_design_excel.py` 对应）

内容区为 B..BC 列。块类型由 B 列单元格的样式决定：

| Excel 样式 | 还原为 Markdown |
| --- | --- |
| A1 标题 | `# 标题` |
| B:BC 整段合并 + 深灰底 `FF666666` + 白字粗体、无边框 | `## 章节` |
| B:BC 整段合并 + 浅灰底 `FFD9D9D9` | `### 子节` |
| B 列有细左边框的连续行（首行为表头） | GFM 管道表格 |
| B:BC 整段合并 + 浅灰底 `FFF2F2F2` + Consolas | 代码围栏（JSON 自动标 `json`） |
| B:BC 整段合并、无底色、无边框 | 正文段落（多行按行还原） |
| 工作表内嵌图片 | 导出 PNG，并以 `![流程图 N](<目标名>_images/mermaid_N.png)` 引用 |

表格单元格：Excel 中的换行还原为 `<br>`，`|` 转义为 `\|`。
第 1–6 行的模板头部（サービス名、作成者等）不进入正文，作为 HTML 注释 `<!-- 版式头部（模板信息） ... -->` 保留。

## 无法还原的信息（转换时已丢失，Excel 中不存在）

- **mermaid 源码**：只能导出为图片，无法还原成 mermaid 代码块。
- **行内反引号与粗体**：生成端会去掉 `` `code` `` 的反引号、不保留 `**粗体**`，还原后为纯文本。
- **表格对齐**：统一输出 `---`。
- **超过 3 列的表格**：生成端已把第 4 列起并入第 3 列（` / ` 连接），还原后不会拆回。
- **第 4 级及以下标题**：生成端已降为子节，还原为 `###`。

## 用法

```powershell
python .agents\skills\doubao-design-doc-excel-to-md\scripts\design_excel_to_md.py --input <源.xlsx> [--output <目标.md>] [--force]
```

| 参数 | 说明 |
| --- | --- |
| `--input` | 输入 xlsx（design_template 版式，必需） |
| `--output` | 输出 md（缺省 `./output/<源文件名>.md`，目录自动创建） |
| `--force` | 输出已存在时允许覆盖（缺省拒绝） |

图片导出到 `<输出文件名>_images/`（与 .md 同级）；没有图片时不建目录。

## 工作流

1. 确认输入是 `doubao-design-doc-excel` 生成的文件（A1 为标题，内容从第 7 行起，B:BC 合并区）。
2. 运行脚本转换。
3. 验证：
   - 章节数（`## `）、子节数（`### `）与原设计书一致。
   - 表格行数与列数抽查；单元格中的 `<br>`、`\|` 正确。
   - 有图片时确认 `*_images/` 下的 PNG 数量与 md 中的引用一致。
   - 按上面的"无法还原"清单，告知用户哪些信息未能恢复。
4. 交付：报告输出的 md 路径（及图片目录）。

## 目录

| 用途 | 位置 |
| --- | --- |
| 脚本 | `scripts/design_excel_to_md.py` |
| 自测（往返校验） | `tests/test_design_excel_to_md.py`，可选参数：真实 md 路径（用于章节数冒烟） |

自测依赖同级的 `doubao-design-doc-excel` 技能（复用它的 `md_to_design_excel.py` 生成测试 xlsx），并需要 `openpyxl`、`Pillow`。

## 依赖

`openpyxl>=3.1`、`Pillow`（导出图片时 openpyxl 读取嵌入图片需要）。
