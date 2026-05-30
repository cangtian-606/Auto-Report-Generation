# Auto-Report-Generation

Word 文档自动化生成系统。将 Word 模板（.docx）结合 Excel / YAML 数据，通过 Jinja2 模板引擎渲染生成填充好的 Word 报告。

> 本项目代码由 **DeepSeek V4 Pro** 辅助编写。

---

## 这是什么？

准备好一个 Word 模板（含 Jinja2 占位符）和一份数据文件（Excel 或 YAML），即可生成排版完成的 Word 报告，消除手工复制粘贴，避免多表数据遗漏。

```
┌──────────────┐       ┌──────────────┐
│ 模板 .docx     │       │ 数据 .xlsx    │
│ 含 {{ 变量 }}  │       │ / .yaml      │
│ 和 {% for %}  │       │              │
└──────┬───────┘       └──────┬───────┘
       │                      │
       └──────────┬───────────┘
                  ▼
        Auto-Report-Generation
                  │
                  ▼
         ┌──────────────┐
         │  报告 .docx   │
         └──────────────┘
```

### 适用场景

| 场景 | 说明 |
|------|------|
| 尽调报告 | 企业财务、法律、业务数据自动填入报告模板 |
| 审计报告 | 被审计单位多张 Excel 表数据汇总生成 Word 报告 |
| 合同批量 | 同一合同模板 + 不同客户数据 → 批量生成 |
| 定期报表 | 每月/每季度固定格式 Word 报表自动更新 |

---

## 快速开始

### 安装

```bash
pip install -r requirements.txt
```

### CLI 用法

```bash
python -m src --data <数据文件> --template <模板文件> --output <输出文件>
```

| 参数 | 简写 | 必填 | 说明 |
|------|:--:|:--:|------|
| `--data` | `-d` | ✅ | 数据文件路径，支持 `.xlsx` 或 `.yaml` |
| `--template` | `-t` | ✅ | Word 模板文件路径（`.docx`） |
| `--output` | `-o` | ✅ | 输出 Word 文件路径（`.docx`） |
| `--strict` | — | — | 严格模式：Schema 校验不通过时报错退出 |
| `--schema` | — | — | 指定自定义 Schema 校验文件路径 |
| `--log-level` | — | — | 日志级别，可选 `DEBUG` / `INFO` / `WARNING` / `ERROR`（默认 `INFO`） |
| `--log-file` | — | — | 日志输出到指定文件（默认自动生成到 `logs/` 目录） |

**示例 — Excel 数据源：**

```bash
python -m src -d data/数据.xlsx -t templates/报告模板.docx -o output/报告.docx
```

**示例 — YAML 数据源：**

```bash
python -m src -d data/数据.yaml -t templates/报告模板.docx -o output/报告.docx
```

**示例 — 严格模式 + 指定日志文件：**

```bash
python -m src -d data/数据.xlsx -t templates/报告模板.docx -o output/报告.docx --strict --log-file logs/run.log
```

### 批量模式

```bash
# 对 data/按编号拆分/ 目录下所有 xlsx 文件，使用同一模板批量生成
python -m src -d data/按编号拆分/ -t templates/报告模板.docx -o output/按编号拆分输出/
```

---

## 核心能力

### 变量替换

模板占位符自动替换为数据值：

| 模板写法 | Excel 值 | 渲染结果 |
|------|------|------|
| `{{ 全局.公司名称 }}` | `东海精密制造有限公司` | 东海精密制造有限公司 |
| `{{ 全局.注册资本 \| num }}` | `50000000` | 50,000,000.00 |
| `{{ 全局.成立日期 \| date }}` | `2020-03-15` | 2020年03月15日 |
| `{{ 全局.成立日期 \| date("chinese") }}` | `2020-03-15` | 二〇二〇年三月十五日 |

### 表格循环

一条模板行自动展开为 N 条数据行：

```
{%tr for i in 历史沿革.股权结构 %}
{{ i.股东名称 }}   {{ i.认缴比例 | percent }}
{%tr endfor %}
```

### 条件显隐

根据数据自动显示或隐藏段落：

```
{%p if 税率政策.有高新优惠 %}
本公司适用高新技术企业优惠税率。
{%p endif %}
```

### 标签层级体系

| 层级 | 前缀 | 能控制什么 | 示例 |
|------|:--:|------|------|
| **块级** | 无 | 多段落 + 表格 | `{% for %}` `{% if %}` |
| **段落区** | `p` | 多个纯文字段落（不含表格） | `{%p if %}` `{%p else %}` |
| **行内** | `r` | 段落中几个字的显隐 | `{%r if %}` `{%r endif %}` |

### 过滤器

| 过滤器 | 效果 |
|--------|------|
| `\| num` | 千分位金额，`1234567.89` → `1,234,567.89` |
| `\| num(0)` | 取整，`1234567.89` → `1,234,568` |
| `\| num(2, "万")` | 除以万单位，`12345678` → `1,234.57` |
| `\| percent` | 小数转百分比，`0.25` → `25.00%` |
| `\| date` | 日期格式化，`2025-01-15` → `2025年01月15日` |
| `\| date("short")` | 短格式，→ `2025/01/15` |
| `\| date("month")` | 月份格式，→ `2025年01月` |
| `\| date("chinese")` | 中文大写，→ `二〇二五年一月十五日` |
| `\| int` | 截断取整，`123.45` → `123` |
| `\| default("-")` | 空值替换为 `-` |
| `\| default_dash` | 空值显示为空字符串 |
| `\| paragraphs` | `\n` 转为 Word 真实段落换行 |

---

## Demo 参考

项目提供了一份完整的 Demo 文件，可作为模板和数据格式的参考：

| 文件 | 说明 |
|------|------|
| `data/demo_data.xlsx` | 23 个 Sheet 的示例数据（键值对 + 表格 + 嵌套） |
| `data/demo_data.yaml` | 等价 YAML 格式的示例数据 |
| `templates/demo_template.docx` | 16 个章节的专业排版模板 |
| `output/demo_output.docx` | 预生成的参考输出 |

Demo 覆盖全部 **30 项语法**：块级/段落区/行内三级标签、行循环/列循环/嵌套循环、布尔/否定/等于/数值比较四种条件、11 种过滤器。

> 公司和项目均为**假名指代**（东海精密、南山控股、A区B市），不含真实商业信息。

---

## 技术栈

- **Python 3.13**
- **docxtpl** — Word 模板渲染引擎
- **Jinja2** — 模板表达式求值（SandboxedEnvironment 防 SSTI）
- **pandas** + **openpyxl** — Excel 数据读取
- **lxml** — Word XML 预处理
- **PyYAML** — YAML 数据源

---

## 文档

| 文档 | 适合谁 | 内容 |
|------|------|------|
| [01-快速入门](docs/01-快速入门.md) | 所有人 | 5 分钟从零到生成 |
| [02-数据格式规范](docs/02-数据格式规范.md) | 业务 / 开发 | Excel/YAML 格式、嵌套、过滤器唯一事实源 |
| [03-用户操作指南](docs/03-用户操作指南.md) | 业务 | 填数据 → 跑命令 → 排查 |
| [04-模板制作规范](docs/04-模板制作规范.md) | 模板制作者 | 报告样本 → 模板 → 数据字典三步法 |
| [05-开发者手册](docs/05-开发者手册.md) | 开发者 | 架构、API、扩展、安全 |
| [06-附录-速查表](docs/06-附录-速查表.md) | 所有人 | 标签层级、过滤器、CLI 参数一览 |

---

## 架构

```
数据文件             模板文件
(.xlsx / .yaml)       (.docx)
     │                    │
     ▼                    ▼
  Reader              Preprocessor
  (自动识别类型)         (Tr→Tc 串联)
     │                    │
     ▼                    ▼
  RenderContext          Pipeline
  (构建+扁平化+校验)       (临时文件管理)
           │                    │
           └──────────┬─────────┘
                      ▼
              DocumentGenerator
              (Sandboxed Jinja2)
                      │
                      ▼
                  output.docx
```

### 设计要点

- **深模块架构** — RenderContext、TemplateAnalyzer、RenderPipeline 聚合为高内聚模块
- **SandboxedEnvironment** — Jinja2 沙箱防 SSTI 注入
- **路径安全** — `validate_path()` 限制文件访问到 `data/`、`templates/`、`output/` 目录
- **双数据源** — Excel 和 YAML 自动识别，统一 Reader 接口
- **结构化日志** — run_id 追踪 + 阶段耗时，支持批量模式故障定位

---

## 测试

```bash
pytest tests/ -q
```

> 当前 121 项测试全部通过。

---

## 许可

MIT
