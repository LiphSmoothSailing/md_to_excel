# Markdown → Excel 转换工具 · 安装说明

> 版权所有 © 李鹏虎 (Li Penghu)　保留所有权利
> 仅供个人学习与非商业用途；商业使用须经作者书面同意。
> 作者邮箱：gs_lipenghu@qq.com

本工具把结构化 Markdown 转换为格式化的 Excel（`.xlsx`），支持标题层级缩进、
表格、列表、任务清单、代码块、图片，以及 **Mermaid 流程图渲染成图片嵌入**。

---

## 目录结构

```
md_to_excel/
├── md_to_excel.py     # 主程序
├── sample.png         # 内置示例引用的配图
├── requirements.txt   # Python 依赖清单
└── 安装说明.md         # 本文件
```

---

## 一、环境要求

| 组件 | 版本要求 | 用途 | 是否必需 |
| --- | --- | --- | --- |
| Python | 3.8 及以上 | 运行主程序 | 必需 |
| Node.js | 18 及以上（建议 LTS） | 渲染 Mermaid 图表 | 仅用 Mermaid 时需要 |

检查是否已安装：

```powershell
python --version
node --version
npm --version
```

---

## 二、安装 Python 依赖（pip）

在 `md_to_excel` 文件夹下执行：

```powershell
pip install -r requirements.txt
```

或手动逐个安装：

```powershell
pip install openpyxl      # 生成 Excel（必需）
pip install Pillow        # 嵌入图片 / Mermaid 图（嵌图必需）
```

| 包名 | 作用 | 不装的后果 |
| --- | --- | --- |
| `openpyxl` | 生成 `.xlsx` | 程序无法运行 |
| `Pillow` | 让 openpyxl 能嵌入图片 | 图片与 Mermaid 图会退化为文字占位符 |

---

## 三、安装 Mermaid 渲染工具（npm）

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

## 四、使用方法

```powershell
# 1) 转换指定文件（输出同名 .xlsx）
python md_to_excel.py 输入.md

# 2) 转换并指定输出文件名
python md_to_excel.py 输入.md 输出.xlsx

# 3) 不带参数：生成内置示例 demo.md -> demo.xlsx（含表格、任务清单、代码块、Mermaid 流程图）
python md_to_excel.py
```

运行时会先显示版权信息，再输出转换结果。

---

## 五、Mermaid 使用示例

在 Markdown 中写：

````markdown
```mermaid
flowchart TD
    A[开始] --> B{是否已登录?}
    B -->|是| C[进入主页]
    B -->|否| D[跳转登录页]
    C --> F[结束]
```
````

转换后，该流程图会以 **图片** 形式嵌入 Excel 对应位置。

---

## 六、常见问题

**Q1：提示「未找到 mmdc」或 Mermaid 变成了代码块文字？**
说明 `mermaid-cli` 未装好或不在 PATH。请重新执行第三节的 `npm install -g @mermaid-js/mermaid-cli`，
然后**新开一个终端**再运行（安装后需重启终端才能识别 `mmdc` 命令）。

**Q2：图片 / Mermaid 图变成灰色文字占位符？**
多半是没装 `Pillow`。执行 `pip install Pillow` 后重试。

**Q3：`npm install` 卡在下载 Chromium？**
这是 mermaid-cli 首次安装的正常现象（需下载浏览器内核）。请保持网络畅通、耐心等待；
如网络受限，可为 npm 配置代理或使用国内镜像源后重试。

**Q3.1：安装报 `npm error code 13` / `EPERM: operation not permitted`？**
这是权限不足（全局目录在受保护路径，如 `C:\Program Files\...`）。二选一解决：

- 方式 A：右键 PowerShell →「以管理员身份运行」，再执行
  `npm install -g @mermaid-js/mermaid-cli`
- 方式 B（免管理员）：把 npm 全局目录改到用户可写位置后重装
  ```powershell
  npm config set prefix "$env:APPDATA\npm"
  npm install -g @mermaid-js/mermaid-cli
  ```

装完都请 **新开终端** 再执行 `mmdc --version` 验证。若之前留下残缺目录，
可先删除 `<npm 全局目录>\node_modules\@mermaid-js` 再重装。

**Q4：Node.js 没装怎么办？**
到 https://nodejs.org 下载并安装 LTS 版本，安装时勾选「加入 PATH」，装完重启终端。

---

## 版权声明

本软件仅供个人学习与非商业用途使用。未经作者书面同意，不得用于任何商业目的
（包括但不限于销售、集成到商业产品、提供商业服务等）。如需商业授权，请先联系作者征得同意。

作者：李鹏虎　邮箱：gs_lipenghu@qq.com
