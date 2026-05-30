<p align="center">
  <img src="https://img.shields.io/badge/Python-3.13-3776AB?style=flat-square&logo=python" alt="Python 3.13" />
  <img src="https://img.shields.io/badge/test-121%20passed-34D058?style=flat-square&logo=pytest" alt="121 passed" />
  <img src="https://img.shields.io/badge/license-MIT-blue?style=flat-square" alt="MIT" />
</p>

<h1 align="center">Auto-Report-Generation</h1>
<p align="center"><strong>Word 文档自动化引擎</strong> — 一份模板 + 一张 Excel，一键生成专业报告</p>

---

## 这是什么？

**Auto-Report-Generation** 是一个轻量的 Word 报告自动化系统。你只需要准备：

- 🧾 一个 **Word 模板**（含 Jinja2 占位符）
- 📊 一份 **Excel / YAML 数据**

就能在 **1 秒内** 生成一份排版精良的 Word 报告，消除手工复制粘贴，避免分散表数据遗漏。

```
┌──────────────┐       ┌──────────────┐
│ 模板 .docx     │       │ 数据 .xlsx    │
│ 含 {{ 变量 }}  │       │ 含 23 Sheet  │
│ 和 {% for %}  │       │ 和 .yaml     │
└──────┬───────┘       └──────┬───────┘
       │                      │
       └──────────┬───────────┘
                  ▼
        Auto-Report-Generation
                  │
                  ▼
         ┌──────────────┐
         │  报告 .docx   │
         │  排版完毕 ✅   │
         └──────────────┘
```

### 🎯 适用场景

| 场景 | 示例 |
|------|------|
| 尽调报告 | 企业财务、法律、业务数据自动填入报告模板 |
| 审计报告 | 被审计单位多张 Excel 表数据汇总生成报告 |
| 合同批量 | 同一合同模板 + 不同客户数据 → 批量生成 |
| 定期报表 | 每月/每季度固定格式报表自动更新数 |

> 内置 **demo 项目**，开箱即跑，无需准备任何文件。

---

## 🚀 30 秒上手

```bash
# 1. 安装
pip install -r requirements.txt

# 2. 使用内置 demo 数据运行
python -m src -d data/demo_data.xlsx -t templates/demo_template.docx -o output/demo_output.docx

# 3. 打开 output/demo_output.docx，查看生成结果
```

也支持 YAML 数据源：

```bash
python -m src -d data/demo_data.yaml -t templates/demo_template.docx -o output/demo_output.docx
```

---

## ✨ 核心能力

### 数据 → 变量替换

Excel 中的字段自动填入 Word 模板对应位置：

| Excel 单元格 | 模板写法 | 渲染结果 |
|:---:|------|------|
| `东海精密制造有限公司` | `{{ 全局.公司名称 }}` | 东海精密制造有限公司 |
| `50000000` | `{{ 全局.注册资本 \| num }}` | 50,000,000.00 |
| `2020-03-15` | `{{ 全局.成立日期 \| date }}` | 2020年03月15日 |
| `2020-03-15` | `{{ 全局.成立日期 \| date("chinese") }}` | 二〇二〇年三月十五日 |
| `0.25` | `{{ 税率 \| percent }}` | 25.00% |

### 表格循环

一条模板行 → 自动展开为 N 条数据行：

```
{%tr for i in 历史沿革.股权结构 %}
{{ i.股东名称 }}   {{ i.认缴比例 | percent }}
{%tr endfor %}
```

### 条件显隐

根据数据有无，自动显示或隐藏段落：

```
{%p if 税率政策.有高新优惠 %}
本公司适用高新技术企业优惠税率。
{%p endif %}
```

### 层级标签体系

| 层级 | 前缀 | 能控制什么 | 示例 |
|------|:--:|------|------|
| **块级** | 无 | 多段落 + 表格 | `{% for %}` `{% if %}` |
| **段落区** | `p` | 多个纯文字段落 | `{%p if %}` `{%p else %}` |
| **行内** | `r` | 段落中几个字 | `{%r if %}` `{%r endif %}` |

### 内置过滤器

| 过滤器 | 效果 |
|--------|------|
| `\| num` | 千分位金额，`1234567.89` → `1,234,567.89` |
| `\| num(2,"万")` | 除以万单位，`12345678` → `1,234.57` |
| `\| percent` | 小数转百分比，`0.25` → `25.00%` |
| `\| date` | 日期格式化，`2025-01-15` → `2025年01月15日` |
| `\| date("chinese")` | 中文大写，→ `二〇二五年一月十五日` |
| `\| int` | 截断取整，`123.45` → `123` |
| `\| default("-")` | 空值替换为 `-` |
| `\| paragraphs` | `\n` 转为 Word 真实段落 |

---

## 📦 内置 Demo

项目自带一个 **完整的制造业尽调报告 demo**，无需准备任何文件即可验证系统所有功能：

| 文件 | 说明 |
|------|------|
| `data/demo_data.xlsx` | 23 个 Sheet，模拟完整的尽调数据 |
| `data/demo_data.yaml` | 等价 YAML 格式数据源 |
| `templates/demo_template.docx` | 16 个章节的专业排版 Jinja2 模板 |
| `output/demo_output.docx` | 预生成的参考输出 |

Demo 覆盖全部 **30 项语法**：块级/段落区/行内三级标签、行循环/列循环/嵌套循环、布尔/否定/等于/数值比较四种条件、11 种过滤器。

> 公司和项目均使用**假名指代**（东海精密、南山控股、A区B市），不含真实商业信息。

---

## 📖 文档

| 文档 | 适合谁 | 内容 |
|------|------|------|
| [01-快速入门](docs/01-快速入门.md) | 所有人 | 5 分钟从零到生成 |
| [02-数据格式规范](docs/02-数据格式规范.md) | 业务 / 开发 | Excel/YAML 格式、嵌套、过滤器唯一事实源 |
| [03-用户操作指南](docs/03-用户操作指南.md) | 业务 | 填数据 → 跑命令 → 排查 |
| [04-模板制作规范](docs/04-模板制作规范.md) | 模板制作者 | 报告样本 → 模板 → 数据字典三步法 |
| [05-开发者手册](docs/05-开发者手册.md) | 开发者 | 架构、API、扩展、安全 |
| [06-附录-速查表](docs/06-附录-速查表.md) | 所有人 | 标签层级、过滤器、CLI 参数一览 |

---

## 🏗️ 架构

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

### 设计要

- **深模块架构** — RenderContext、TemplateAnalyzer、RenderPipeline 聚合为高内聚模块
- **SandboxedEnvironment** — Jinja2 沙箱防 SSTI 注入
- **路径安全** — 文件访问限制在 `data/`、`templates/`、`output/` 目录
- **双数据源** — Excel 和 YAML 各取所需，自动识别
- **结构日志** — run_id 追踪 + 阶段耗时，支持批量模式故障定位

---

## 🧪 测试

```bash
pytest tests/ -q          # 121 tests
```

---

## 📄 许可

MIT
