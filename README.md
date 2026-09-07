# Markdown → 設計書 Excel 转换工具

> 版权所有 © 李鹏虎 (Li Penghu)　保留所有权利
> 仅供个人学习与非商业用途；商业使用须经作者书面同意。
> 作者邮箱：gs_lipenghu@qq.com

将 Markdown 设计书（API 設計書 / 要件定義書 / 詳細設計書 / 详细设计书）转换为
遵循 `design_template.xlsx` 版式的 Excel 文件。版式基准由用户微调确认，
**所有设计书转换一律使用本模板，不得改用其他模板或另起版式**。

---

## 目录结构

```
md_to_excel/
├── temp/                             # 转换过程的临时文件 / 验证脚本（可随时清理）
├── output/                           # 转换产物默认输出目录
├── .agents/skills/
│   └── doubao-design-doc-excel/      # 转换 Skill（SKILL.md + 转换脚本 + 模板）
│       ├── SKILL.md                  # 版式铁律与工作流说明
│       ├── scripts/md_to_design_excel.py   # 转换脚本（MD 解析 / blocks 渲染）
│       └── assets/design_template.xlsx     # 设计书 Excel 模板基准（唯一权威，勿改）
├── requirements.txt                  # Python 依赖清单
├── LICENSE                           # 许可协议
└── README.md                         # 本文件
```

目录约定（脚本内置默认，可用参数覆盖）：

| 用途          | 位置                                    | 覆盖参数         |
| ----------- | ------------------------------------- | ------------ |
| 模板文件（唯一权威）  | Skill 内 `assets/design_template.xlsx` | `--template` |
| 临时文件 / 验证脚本 | `./temp/`                             | —            |
| 转换产物        | `./output/<源文件名>.xlsx`                | `--output`   |

---

## 环境要求

| 组件      | 版本要求           | 用途            | 是否必需           |
| ------- | -------------- | ------------- | -------------- |
| Python  | 3.8 及以上        | 运行主程序         | 必需             |
| Node.js | 18 及以上（建议 LTS） | 渲染 Mermaid 图表 | 仅用 Mermaid 时需要 |

检查是否已安装：

```powershell
python --version
node --version
npm --version
```

---

## 安装 Python 依赖（pip）

在 `md_to_excel` 文件夹下执行：

```powershell
pip install -r requirements.txt
```

或手动逐个安装：

```powershell
pip install openpyxl      # 生成 Excel（必需）
pip install Pillow        # 嵌入图片 / Mermaid 图（嵌图必需）
```

| 包名         | 作用               | 不装的后果                  |
| ---------- | ---------------- | ---------------------- |
| `openpyxl` | 生成 `.xlsx`       | 程序无法运行                 |
| `Pillow`   | 让 openpyxl 能嵌入图片 | 图片与 Mermaid 图会退化为文字占位符 |

---

## 安装 Mermaid 渲染工具（npm）

Mermaid 是文本描述语言，Excel 无法直接显示，需先用官方命令行工具
`mermaid-cli` 把图渲染成 PNG，再嵌入 Excel。它依赖 Node.js，请按下面两步顺序安装。

### 第 1 步：安装 Node.js（npm 随 Node.js 一起装）

1. 打开官网 https://nodejs.org
2. 下载 **LTS（长期支持）** 版本并运行安装程序
3. 安装过程中保持勾选「Add to PATH / 添加到环境变量」（默认已勾选）
4. 安装完成后 **重启终端（PowerShell / 命令提示符）**

验证 Node.js 与 npm 都可用：

```powershell
node --version
npm --version
```

> npm 是 Node.js 自带的包管理器，装好 Node.js 后即可使用，无需单独安装。

### 第 2 步：用 npm 安装 mermaid-cli

```powershell
npm install -g @mermaid-js/mermaid-cli
```

- 需要安装的 npm 包名：**`@mermaid-js/mermaid-cli`**
- 安装后提供命令：**`mmdc`**
- `-g` 表示全局安装；首次安装会自动下载一个 Chromium 内核（约几百 MB），请耐心等待并保持联网。

验证安装成功（**装完请新开一个终端**再执行）：

```powershell
mmdc --version
```

能打印版本号即表示成功。

> **说明**：如果不需要 Mermaid 功能，可跳过本节。届时遇到 ` ```mermaid ` 代码块时，
> 程序会自动**回退**为普通代码块显示原始 mermaid 文本，内容不会丢失，只是不渲染成图。

---

## 模板版式铁律（转换时不可违反）

1. **保留模板前 6 行头部**：标题、サービス名/案件名/作成日/作成者/更新日/更新者信息区、
   プロジェクト種別/ID/予定工数/リリース予定日及复选框区——值、样式、合并、行高原样保留；
   模板自带 `=#REF!` 断链公式清空值但保留样式；模板首 sheet 原为 hidden，必须改为 visible。
2. **列宽与模板完全一致，绝不改列宽**：默认列宽 0.8cm（≈4.3 字符）；显式列宽仅
   `A=3.25`、`AF=4.5`、`AG=3.25`。转换后不得写入任何新列宽。
3. **长内容一律用合并单元格承载**：内容从第 7 行起，按模板列段合并，禁止拉宽列或拆列。
4. **配色/字体沿用模板体系**：章节标题深灰 `FF666666` 底白字 Meiryo 11 粗体；
   子节 `FFD9D9D9` 底黑字 Meiryo 10 粗体；表头深灰底白字 Meiryo 10 粗体 + thin 黑边框；
   数据白底 Meiryo 9 + thin 黑边框；代码 `FFF2F2F2` 底 Consolas 10；正文 Meiryo 10 无填充。
5. **输出单 sheet**：只保留一个设计书 sheet，且可见。

内容区列段（B..BC，1-based）：

| 表格列数               | 列段合并                                       |
| ------------------ | ------------------------------------------ |
| 3 列                | 动态分配（按各列内容显示宽度占比）                        |
| 2 列                | 动态分配（同上）                                  |
| 1 列 / 章节 / 正文 / 代码 | `B:BC`（约 682）                              |

2/3 列表格按内容**动态分配列段**：脚本统计各列「表头+数据」的最大显示宽度，
按占比把 B..BC 切成 N 段（每段至少 1 列、边界取整列、**不改任何列宽**），
内容多的列更宽、内容少的列更窄。表格超过 3 列时，脚本自动把第 4 列起并入
第 3 列（表头/内容用 ` / ` 连接），不丢内容。

---

## 使用方法

### 推荐：大模型 Agent 直接调用 Skill

在豆包等大模型 Agent 中，把任务一句话说给 Agent 即可——Agent 会自动读取
`.agents/skills/doubao-design-doc-excel/` 技能，并按其中的版式铁律执行转换：

> 请按照 .agents/skills/doubao-design-doc-excel/ 技能将 XXXXX.md 转换成 excel

把 `XXXXX.md` 换成实际文件路径（可带完整路径或相对路径）。Agent 按 Skill 自动完成：

- 加载模板（Skill 内 `assets/design_template.xlsx`，唯一权威）
- 保留模板前 6 行头部、列宽与模板完全一致、长内容用合并单元格承载
- 渲染表格 / 代码块 / 段落，```mermaid 流程图渲染为图片嵌入
- 产物默认输出到 `./output/XXXXX.xlsx`（不手动指定时）

该方式适用于任何"MD 转设计书 Excel"任务（API 設計書 / 要件定義書 / 詳細設計書 等），
无需记住脚本参数。

### 方式一：直接转换 Markdown（保留原文语言）

```powershell
python .agents\skills\doubao-design-doc-excel\scripts\md_to_design_excel.py --input <源.md>
```

- 缺省输出到 `./output/<源文件名>.xlsx`（目录自动创建）
- 文档标题（A1）取 MD 首个 `#` 标题

指定输出路径与标题：

```powershell
python .agents\skills\doubao-design-doc-excel\scripts\md_to_design_excel.py `
  --input <源.md> --output <目标.xlsx> --title "<A1标题>"
```

### 方式二：渲染 blocks JSON（适合先翻译/整理内容）

当需要把 MD 翻译成日文后再转换（专业术语、PHP 代码、URL、文件路径、页面文案原样保留），
先把内容整理为 blocks JSON，再渲染：

```powershell
python .agents\skills\doubao-design-doc-excel\scripts\md_to_design_excel.py `
  --blocks <blocks.json> --output <目标.xlsx> --sheet-name "<sheet名>"
```

blocks JSON 结构（列表，元素按序渲染）：

```json
[
  ["chapter", "1. ドキュメント概要"],
  ["table", ["項目", "内容"], [["機能", "キャンセル"], ["対象ページ", "cancel.php"]]],
  ["note", "正文段落，可含 \\n 多行"],
  ["code", "代码 / JSON 示例"]
]
```

- `chapter`：章节标题（深灰底白字）
- `section`：子节标题（浅灰底）
- `table`：表格（1~3 列；`["列1","列2",...], [[值...], ...]`）
- `note`：正文段落
- `code`：代码 / JSON 示例（浅灰底 Consolas）

### 参数一览

| 参数             | 说明                                             |
| -------------- | ---------------------------------------------- |
| `--input`      | 输入 Markdown 路径（与 `--blocks` 二选一）               |
| `--blocks`     | blocks JSON 路径（与 `--input` 二选一）                |
| `--output`     | 输出 xlsx 路径（缺省：`./output/<源文件名>.xlsx`）          |
| `--template`   | 模板路径（缺省：Skill 内 `assets/design_template.xlsx`） |
| `--title`      | A1 文档标题（缺省：MD 模式取首个 `#` 标题，blocks 模式保留模板标题）    |
| `--sheet-name` | sheet 名（缺省：取 A1 标题截断 31 字符）                    |

---

## 环境要求

| 组件       | 版本   | 用途         | 必需  |
| -------- | ---- | ---------- | --- |
| Python   | 3.8+ | 运行转换脚本     | 必需  |
| openpyxl | 3.x  | 读写 `.xlsx` | 必需  |

安装依赖：

```powershell
pip install -r requirements.txt
```

`requirements.txt` 内容：

```
openpyxl>=3.1
```

---

## 验证与交付

转换完成后建议按以下要点核验（可用 openpyxl 回读）：

- 单 sheet 且可见；sheet 名正确
- 列宽：`A=3.25 / AF=4.5 / AG=3.25`、默认 0.8cm（≈4.3 字符），无其他显式列宽
- 前 6 行合并区与模板逐一相等；值仅允许 A1 标题与 `#REF!` 清理两处差异
- 内容从第 7 行起完整；合并区无重叠
- 行高按内容估算，长代码块（JSON 示例）不截断

## 版权声明

本软件仅供个人学习与非商业用途使用。未经作者书面同意，不得用于任何商业目的
（包括但不限于销售、集成到商业产品、提供商业服务等）。如需商业授权，请先联系作者征得同意。

作者：李鹏虎　邮箱：gs_lipenghu@qq.com
