# 计算机实践类实验报告 Skill（Computer Lab Report Writer）

面向计算机实践类课程（当前先支持数据库实验）的实验报告写作与 Word 排版 Skill。目标是解决 AI 生成实验报告时常见的格式问题：

- 表格、标题出现蓝色
- 全文使用“微软雅黑”或依赖 Word 主题字体
- 正文没有首行缩进
- 代码、表格、图片的样式互相继承，导致排版混乱

## 当前项目结构

```
experimental-skill/
├── README.md                     # 本文件
├── Skill.md                      # Skill 入口说明（待改写为 SKILL.md，见“待办”）
├── references/
│   ├── word-format-spec.md       # Word 排版规范（强制）
│   └── font/                     # 本地字体文件，仅供参考，不随 Skill 分发
│       ├── Kaiti.ttc
│       ├── STHeiti Medium.ttc
│       └── Times New Roman.ttf
├── scripts/
│   ├── docx_styles.py            # 样式与字体定义（规范常量的唯一来源）
│   ├── table_builder.py          # 表格一（非数据类）/ 表格二（数据类，三线表）
│   ├── code_block.py             # 代码块：Pygments 高亮 + 行号两列表格
│   └── validate_docx.py          # 生成后的字体与格式扫描
└── templates/
    ├── excel_template.docx       # 表格样式参考（表格一、表格二）
    ├── code_template.docx        # 代码展示样式参考
    ├── E-R_template.md           # Mermaid 黑白主题参考（实际内容为 UML 类图）
    └── ER-crowsfoot_template.md  # E-R 图（鸡爪画法）模板，样式参考 UML 模板
```

## 排版规范摘要

完整数值见 `references/word-format-spec.md`，以下为要点。

| 对象 | 规范 |
|---|---|
| 页面 | A4，上下边距 2.54 cm，左右边距 3.17 cm |
| 正文 | 中文宋体、西文 Times New Roman，小四（12 pt），首行缩进 2 字符，1.25 倍行距，段前段后各 0.5 行（按 12 pt 折算为 6 pt） |
| 标题 | 黑体、黑色、无缩进，不使用主题色 |
| 表题 | 位于表格上方，宋体五号，居中，格式“表1 …” |
| 图题 | 位于图片下方，宋体五号，居中，格式“图1 …” |
| 表格、图片 | 与前后正文各留 0.5 行（6 pt）；题注与对象之间约 3 pt |
| 表格内文字 | 不继承正文首行缩进，使用独立段落样式 |
| 代码 | Courier New 9 pt，行号列与代码列分开，使用独立样式 |

### 两种表格

- **表格一（非数据类，如环境配置）**：全框线 0.5 pt，表头灰底 D7D7D7，黑体 12 pt 加粗居中。
- **表格二（数据类）**：三线表，表头上线 3 pt、表头下线 1 pt，表头黑体 12 pt 不加粗，跨页重复表头。

### 代码块

- 两列表格：左列行号（Word 自动编号 `01.`，底色 E7E5DC），右列代码。
- 保留模板中的绿色列间竖线（65E266）。
- 偶数行底色 F8F8F8，语法高亮沿用模板的 GitHub 风格配色。
- 每个代码块的行号从 01 重新开始。
- 代码行不使用首行缩进。

### 字体规则

- 允许：宋体、黑体、Times New Roman、Courier New。
- 禁止：楷体（暂不使用，待另行说明用途）、微软雅黑。
- 不依赖主题字体或默认字体，每处都显式写出 `ascii`、`hAnsi`、`eastAsia`、`cs` 四个属性。
- 字体文件不放进 Skill，只做检测。

## 图表约定

- **E-R 图**：使用鸡爪（Crow's Foot）表示法，黑白配色。
- **UML 图、流程图**：使用 Mermaid，沿用 `templates/E-R_template.md` 中的黑白主题（白底、黑线、黑字）。
- 图片统一居中，图题在下方。

## 生成后检查

`scripts/validate_docx.py` 对生成的 DOCX 做全文扫描，范围包括：

- 正文段落、各级标题
- 表格所有单元格
- 图题和表题
- 代码块
- 页眉和页脚
- 样式继承、主题字体以及 `fontTable.xml`

发现楷体、微软雅黑、主题字体或未明确指定的字体时，输出错误报告。

## 依赖

- Python 3
- `python-docx`、`lxml`
- `Pygments`（代码高亮）
- 计划中：`mermaid-cli`（Mermaid 渲染：UML 类图、流程图，以及鸡爪画法 E-R 图，使用 `erDiagram`）、`pytest`

## 开发状态

| 模块 | 状态 |
|---|---|
| `word-format-spec.md` | 已编写，题注格式为“表1 / 图1”（全文连续编号），E-R 图采用鸡爪画法 |
| `docx_styles.py` | 已编写，尚未运行测试 |
| `table_builder.py` | 已编写，题注格式待同步 |
| `code_block.py` | 已编写，尚未运行测试 |
| `validate_docx.py` | 已编写，尚未运行测试 |
| 单元测试、集成测试 | 未开始 |
| 图表渲染、截图管理 | 未开始 |
| `build_docx.py`（总入口） | 未开始 |
| `SKILL.md` 与数据库实验专用说明 | 未开始 |

> 开发环境里 Python 与依赖的检查命令曾被拦截，脚本目前只做了静态编写，尚未在本机实际运行。

## 待办

1. 将 `Skill.md` 改写为 `SKILL.md`，内容改为计算机实践类实验报告的触发条件和流程。
2. 统一题注格式为“表1 / 图1”，并同步规范文档与脚本。
3. 在本机安装依赖并运行脚本，补充单元测试，尤其是字体检测的反例。
4. 新增 `build_docx.py`、图表渲染和截图管理脚本。
5. 新增 `references/lab-types/database.md`，写明数据库实验的章节、E-R 图和 SQL 代码展示要求。
6. 用 Word 实际打开生成的文件，核对字体、缩进和间距。
