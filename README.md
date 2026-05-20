# Word模板自动化生成系统

## 文档导航

| 文档 | 读者 | 说明 |
|------|------|------|
| [用户指南](docs/用户指南.md) | 业务人员 | Excel数据填写、数据格式、报告生成 |
| [开发手册](docs/开发手册.md) | 开发人员/模板维护者 | 模板制作、系统架构、API、扩展开发、替换规范 |

---

## 系统概述

本系统是一套**通用的Word文档自动化生成解决方案**，通过将Word模板与Excel数据结合，实现标准化文档的批量自动生成。

### 核心能力

| 能力 | 说明 |
|------|------|
| 变量替换 | 模板中的 `{{变量}}` 自动替换为Excel数据 |
| 条件显示 | 根据数据条件显示/隐藏整段或整行 |
| 循环生成 | 根据数据列表动态生成表格行 |
| 批量处理 | 一键生成多份文档 |
| 格式保留 | 完全保留Word模板的排版样式 |

### 典型应用

| 场景 | 说明 |
|------|------|
| 合同/协议 | 批量生成个性化合同、协议 |
| 证书/证明 | 批量制作证书、证明文件 |
| 报告生成 | 定期报告模板填充 |
| 数据汇总 | Excel数据转Word文档 |
| 函件处理 | 批量生成通知函、确认函 |

---

## 工作流程

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  Word模板   │     │  Excel数据   │     │   Python    │
│  (含变量)   │  +  │  (键值对)    │  +  │  渲染脚本    │
└─────────────┘     └─────────────┘     └─────────────┘
                                                │
                                                ▼
                                        ┌─────────────┐
                                        │  填充好的   │
                                        │  Word文档   │
                                        └─────────────┘
```

---

## 项目结构

```
Auto-Report-Generation/
├── src/                         # 源码
│   ├── __init__.py
│   ├── renderer.py              # 模板渲染引擎
│   └── converter.py             # 模板转换工具
├── tests/                       # 测试
│   ├── __init__.py
│   └── test_report.py
├── docs/                        # 文档
│   ├── 用户指南.md              # 业务人员操作指南
│   └── 开发手册.md              # 模板制作+开发+规范
├── templates/                   # Word模板
├── data/                        # Excel数据文件
├── config/                      # 配置文件
│   └── example_config.json
├── output/                      # 输出目录
├── requirements.txt
├── .gitignore
└── README.md
```

---

## 快速入门

### 环境准备

创建虚拟环境并安装依赖：

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

依赖项：

```bash
pip install docxtpl pandas openpyxl python-docx
```

验证安装：

```bash
python -c "from docxtpl import DocxTemplate; print('OK')"
```

---

## 快速示例

### 步骤1：创建Word模板

新建Word文档，输入以下内容：

```
合同编号：{{ date.合同编号 }}
签订日期：{{ date.签订日期 }}

甲方：{{ date.甲方 }}
乙方：{{ date.乙方 }}

{%p if date.有附件 %}
附件：{{ date.附件列表 }}
{%p endif %}

{%tr for item in form.商品列表 %}
{{ item.名称 }}  {{ item.数量 }}件  单价¥{{ item.单价 }}  合计¥{{ item.金额 }}
{%tr endfor %}

金额合计：¥{{ form.合计金额 }}

甲方签字：___________    乙方签字：___________
```

保存为 `templates/template.docx`。

### 步骤2：创建Excel数据

新建Excel文件，Sheet名称 `date.全局`，包含：

| 字段编码 | 值 |
|----------|-----|
| date.合同编号 | HT-2025-001 |
| date.签订日期 | 2025年1月15日 |
| date.甲方 | 北京科技有限公司 |
| date.乙方 | 上海贸易有限公司 |
| date.有附件 | TRUE |
| date.附件列表 | 附件1：产品规格书 |

第二个Sheet命名为 `form.商品列表`（用于循环数据），包含：

| 名称 | 数量 | 单价 | 金额 |
|------|----------|-------|--------|
| 产品A | 100 | 50.00 | 5000.00 |
| 产品B | 50 | 80.00 | 4000.00 |
| 产品C | 200 | 30.00 | 6000.00 |

保存为 `data/data.xlsx`。

### 步骤3：生成文档

```bash
python -m src.renderer --data data/data.xlsx templates/template.docx output/output.docx
```

### 步骤4：查看结果

打开 `output/output.docx`，内容应为：

```
合同编号：HT-2025-001
签订日期：2025年1月15日

甲方：北京科技有限公司
乙方：上海贸易有限公司

附件：附件1：产品规格书

产品A  100件  单价¥50.00  合计¥5000.00
产品B  50件  单价¥80.00  合计¥4000.00
产品C  200件  单价¥30.00  合计¥6000.00

金额合计：¥15000.00

甲方签字：___________    乙方签字：___________
```

---

## 批量生成

将多个Excel数据文件放入同一目录：

```
data/batch/
├── 合同1.xlsx
├── 合同2.xlsx
└── 合同3.xlsx
```

执行批量生成：

```bash
python -m src.renderer --batch data/batch/ templates/template.docx output/
```

生成结果保存在 `output/` 目录。

---

## 常见问题

| 问题 | 解决方法 |
|------|---------|
| 变量显示为空 | 检查Excel字段编码是否匹配 |
| 金额格式错误 | 使用 `{{ value \| money }}` 过滤器 |
| 条件不生效 | 布尔值用 `TRUE`/`FALSE`（大写）|
| 表格行没循环 | 检查Sheet命名是否为 `form.名称` |

---

## 版本信息

- 版本：1.0.0
- Python：3.8+
- 许可证：内部使用
