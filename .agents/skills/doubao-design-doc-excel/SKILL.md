---
name: doubao-design-doc-excel
description: 把 Markdown 设计书（API 設計書 / 要件定義書 / 詳細設計書 / 详细设计书）转换为遵循 design_template.xlsx 版式的 Excel 文件。适用于用户要求"MD转Excel"、"设计书转Excel"、"markdown转表格"、"按模板生成设计书Excel"、"設計書をExcelに変換"、"要件定義書をエクセル化"等任务；也适用于任何要求保留模板前6行头部、模板列宽、以合并单元格承载长内容、不改列宽的表格转换。转换前必须先读本 SKILL.md。
---

# MD → 設計書 Excel（design_template 版式）

将 Markdown 设计书渲染为 Excel，版式严格以模板基准 `design_template.xlsx` 为准。该模板由用户微调确认，**任何后续设计书转换都使用它，不得改用其他模板或新建版式**。

## 目录约定（项目内，脚本内置默认）

| 用途 | 位置 | 覆盖参数 |
| ---- | ---- | ---- |
| 模板文件（唯一权威，只留这一份） | Skill 内 `assets/design_template.xlsx`（`--template` 可显式覆盖） | `--template` |
| 转换过程临时文件 / 验证脚本 | `./temp/` | — |
| 转换产物 | `./output/<源文件名>.xlsx`（自动创建目录） | `--output` |

## 模板版式铁律（用户强约束，不可违反）

1. **保留模板前 6 行头部**：A1 标题（Meiryo 14 粗体）、サービス名/案件名/作成日/作成者/更新日/更新者信息区、プロジェクト種別/ID/予定工数/リリース予定日及复选框区。值、样式、合并、行高原样保留；模板自带的 `=#REF!` 断链公式清空值但保留样式；模板首 sheet 原为 hidden，必须改为 visible。
2. **列宽与模板完全一致，绝不改列宽**：默认列宽 0.8cm（≈4.3 字符）；显式列宽仅 `A=3.25`、`AF=4.5`、`AG=3.25`。转换后不得写入任何新列宽。
3. **长内容一律用合并单元格承载**：内容从第 7 行起，按模板列段合并（见下），禁止拉宽列或拆列。
4. **配色/字体沿用模板体系**：章节标题 = 深灰 `FF666666` 底 + 白字 Meiryo 11 粗体；子节 = `FFD9D9D9` 底 + 黑字 Meiryo 10 粗体；表头 = 深灰底白字 Meiryo 10 粗体 + thin 黑边框；数据 = 白底 Meiryo 9 + thin 黑边框；代码 = `FFF2F2F2` 底 + Consolas 10；正文 = Meiryo 10 无填充。垂直居中，短列（≤10 显示宽）居中、长列左对齐。
5. **输出单 sheet**：删除模板其余 sheet，文件只保留一个设计书 sheet，且可见。
6. **mermaid 代码块渲染为图片嵌入**：MD 中的 ```` ```mermaid ```` 块用 mermaid-cli（`mmdc`）渲染为 PNG 后嵌入 Excel 对应行（行高撑开、按行高上限自动缩放、不溢出）。mmdc 缺失或渲染失败时降级为该代码块文本（浅灰底 Consolas）。依赖：Node.js + `@mermaid-js/mermaid-cli`（提供 `mmdc`）+ Pillow。

## 内容区列段（1-based，B..BC）

- 2/3 列表格：**动态分配**——脚本统计各列「表头+数据」最大显示宽度，按占比把 B..BC 切成 N 段（每段至少 1 列、边界取整列、不改任何列宽），内容多的列宽、内容少的列窄
- 1 列表格 / 章节 / 子节 / 正文 / 代码：`B:BC`（约 682）
- 表格 >3 列：脚本自动把第 4 列起并入第 3 列（表头/内容用 ` / ` 连接），不丢内容

## 工作流

1. **读全源 MD**（UTF-8），确认文档结构：`#` 标题 → 章节/文档标题（按首个标题级别动态推导 `##`=章节、`###`=子节）、`|` 表格、``` 代码块（含 ```mermaid 流程图）、段落。
2. **确定语言**：源 MD 已是目标语言 → 用 `--input` 模式直接转换；用户要求日文/中文而 MD 是另一语言 → 先翻译（专业术语、PHP 代码、URL、文件路径、页面文案原样保留），把翻译结果整理为 blocks JSON 后用 `--blocks` 模式。blocks 结构见脚本 docstring。
3. **运行脚本**：

```bash
python <skill>/scripts/md_to_design_excel.py --input <源.md> [--output <目标.xlsx>] [--title <A1标题>]
python <skill>/scripts/md_to_design_excel.py --blocks <blocks.json> --output <目标.xlsx> [--title ...] [--sheet-name ...]
```

   - `--output` 缺省时输出到 `./output/<源文件名>.xlsx`（CWD 下自动创建）；`--title` 缺省时 MD 模式取首个 `#` 标题、blocks 模式保留模板 A1 标题；`--sheet-name` 缺省时取 A1 标题截断 31 字符。转换产生的临时文件（mermaid 图片等）放 `./temp/` 下并在保存后清理。
4. **验证**（交付前必做，openpyxl 回读）：
   - 单 sheet 且 `sheet_state == 'visible'`；sheet 名正确。
   - 列宽：`A=3.25 / AF=4.5 / AG=3.25`、`defaultColWidth≈0.8cm（4.3 字符）`，无其他显式列宽。
   - 前 6 行合并区与模板逐一相等（17 个）；值仅允许 A1 标题与 `#REF!` 清理两处差异。
   - 内容从第 7 行起完整：章节标题计数、关键表格数据抽查；合并区无重叠。
   - 行高：按内容估算，长代码块（如 JSON 示例）行高 = 行数×13+8，正文 = 行数×14+8，不截断。
   - mermaid 图已嵌入：`ws._images` 非空、锚点行高撑开（≤ 约 520px 高，不溢出覆盖其他内容）；若期望有图而 `_images` 为空，检查 `mmdc` 是否可用。
   - 用 artifact-preview 渲染文本核对全部章节与关键术语。
5. **交付**：`present_files` 交付 xlsx 本地路径。

## 资源

- `assets/design_template.xlsx` — 版式基准模板（用户微调确认版），**模板唯一权威**，脚本默认从本 Skill assets 加载；勿修改、勿另存副本。
- `scripts/md_to_design_excel.py` — 转换脚本（解析 MD / 渲染 blocks JSON），已用真实设计书 MD 与日文 blocks 双路径测试通过。
