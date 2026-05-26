# Auto-Report-Generation

Word 文档自动化生成系统。将 Word 模板（.docx）结合 Excel/YAML 数据，通过 Jinja2 模板引擎渲染生成填充好的 Word 文档。

## Tech stack

- **Python 3.13**
- **docxtpl** — Word 模板渲染引擎，基于 python-docx + Jinja2
- **jinja2** — 模板表达式求值（SandboxedEnvironment 防 SSTI）
- **pandas** + **openpyxl** — Excel 数据读取
- **lxml** — Word XML 预处理（表格标签注入）
- **pyyaml** — YAML 数据源支持

## Architecture at a glance

```
数据文件             模板文件
(.xlsx / .yaml)      (.docx)
     │                   │
     ▼                   ▼
  Reader            Preprocessor
  (xlsx.py/)        (table_preprocessors.py)
  (yaml.py)         Tr → Tc 串联执行
     │                   │
     ▼                   │
  DataMapper              │
  (mapper.py)             │
  ┌─ 键值对 → dict        │
  ├─ 表格   → list[dict]  │
  └─ 嵌套   → 子表挂载    │
     │                   │
     ▼                   │
  _flatten_nested_lists() │
  └─ 子表扁平化为顶级路径   │
     │                   │
     ▼                   ▼
        DocumentGenerator
        (generator.py + filters.py)
        SandboxedEnvironment
              │
              ▼
         output.docx
```

### Render context 数据结构

渲染上下文是一个嵌套 dict：

```
context["全局"]                              = {"公司名称": "...", "报告文号": "..."}
context["释义"]                               = [{"全称": "...", "简称": "..."}, ...]
context["历史沿革"]                            = {"注册资本": 1000, "股权结构": [...], ...}
context["项目基本情况"]                         = [{"项目名称": "...", "备案文件": "..."}, ...]
context["EPC合同"]                            = [{"承包单位": "...", "epc金额": [...]}, ...]
context["EPC合同.epc金额"]     ← 扁平化后      = [...所有子项合并的顶级列表...]
context["税率政策"]                            = {"有三免三减半": True, "企业所得税税率": 0.25, ...}
```

## Domain glossary

| 术语 | 含义 |
|------|------|
| **模板** | 含 Jinja2 变量占位符的 .docx 文件 |
| **数据文件** | 提供渲染数据的 .xlsx 或 .yaml 文件 |
| **context** | Mapper 产出的嵌套 dict，传给 Jinja2 模板渲染 |
| **Sheet** | Excel 中的一个工作表，对应 context 中的一个顶级键 |
| **键值对 Sheet** | 首列是 `字段编码` 的 Sheet，每个字段映射为一个值 |
| **表格 Sheet** | 首行是列标题的 Sheet，每行是一条数据记录 |
| **嵌套 Sheet** | Sheet 名含 `.` 表示父子关系（如 `历史沿革.股权结构`） |
| **_parent_ 列** | 子表中以 `_parent_` 开头的列，值匹配父表同名列以建立关联 |
| **扁平化** | `_flatten_nested_lists()` 将列表内的子表合并到顶级 context |
| **{%tr for %}** | 表格行循环标签 |
| **{%tc for %}** | 表格列循环标签 |
| **{%p if %}** | 段落级条件标签 |
| **{% for %}** | 段落级循环标签 |

## Document index

| 文档 | 读者 | 说明 |
|------|------|------|
| [01-快速入门](docs/01-快速入门.md) | 所有人 | 5 分钟从零到生成第一份文档 |
| [02-数据格式规范](docs/02-数据格式规范.md) | 所有人 | Excel/YAML 格式、变量、过滤器、嵌套 — **唯一事实源** |
| [03-用户操作指南](docs/03-用户操作指南.md) | 业务人员 | 填写数据、执行命令、排查问题 |
| [04-模板制作规范](docs/04-模板制作规范.md) | AI / 模板制作者 | 报告样本 → 模板 + 数据字典 |
| [05-开发者手册](docs/05-开发者手册.md) | 开发者 | 架构、API、扩展、部署 |
| [06-附录-速查表](docs/06-附录-速查表.md) | 所有人 | 语法、过滤器、CLI 参数一览 |

## Project structure

```
templates/                   ← 模板文件 (.docx)
data/                        ← 数据文件 (.xlsx / .yaml)
output/                      ← 生成输出
docs/                        ← 项目文档 (01 ~ 06)
src/
├── __main__.py              # python -m src 入口
├── orchestrator.py          # 编排：_build_context → _flatten → _preprocess → _render
├── cli.py                   # CLI 参数解析 + 批量
├── path_guard.py            # 路径安全校验
├── logging_config.py        # 日志配置
├── exceptions.py            # 自定义异常体系
├── reader/
│   ├── xlsx.py              # Excel 读取（自动识别键值对/表格/嵌套）
│   └── yaml.py              # YAML 读取
├── processing/
│   ├── mapper.py            # DataMapper：原始数据 → context dict
│   ├── schema.py            # SchemaValidator：context 完整性校验
│   └── table_preprocessors.py  # TcInheritance + TrInheritance 预处理器
└── render/
    ├── generator.py         # DocumentGenerator：docxtpl 封装 + 变量检查
    └── filters.py           # 自定义过滤器：money / percent / num / date / int / default
tests/                       # 107 个测试（pytest）
```

## Key design decisions

1. **数据格式文档是唯一事实源** — 过滤器定义、变量语法等所有规范以 `02-数据格式规范.md` 为准，其他文档通过引用避免重复。
2. **中文 Sheet 名和字段名** — 全部使用见名知义的中文命名，降低业务人员使用门槛。
3. **自动识别 Sheet 类型** — Reader 根据首行首列是否匹配 `KEY_NAMES`（字段编码/key/名称等）自动判断键值对还是表格。
4. **Tr → Tc 串联预处理** — 渲染前通过 XML 操作自动注入循环标签，模板制作者只需首行声明。
5. **SandboxedEnvironment** — Jinja2 沙箱防止 SSTI 攻击。
6. **路径安全** — `validate_path()` 限制文件访问到 `data/`、`templates/`、`output/` 及项目根目录。

## Agent skills

### Issue tracker

Issues 以本地 Markdown 文件存放在 `.scratch/<feature-slug>/` 目录下。参见 `docs/agents/issue-tracker.md`。

### Triage labels

使用标准五标签词汇（needs-triage / needs-info / ready-for-agent / ready-for-human / wontfix），与默认值一致无映射。参见 `docs/agents/triage-labels.md`。

### Domain docs

Single-context 布局：`CONTEXT.md` 在仓库根目录，`docs/adr/` 存放架构决策记录。参见 `docs/agents/domain.md`。
