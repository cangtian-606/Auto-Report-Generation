# 重构与功能扩充实施计划

> **Goal:** 拆分单文件模块、添加类型提示、迁移 pytest、新增日期过滤器、Schema 验证器、模板语法检查
> **Architecture:** 将 renderer.py 垂直拆分为 reader/mapper/generator 三模块，generator 引入 Schema 验证和语法检查
> **Tech Stack:** Python 3, pytest, docxtpl, pandas

---

## 一、架构变更

### 新目录结构

```
src/
├── __init__.py                    # 公开 API + 向后兼容别名
├── reader.py                      # ExcelDataReader（从 renderer.py 提取）
├── mapper.py                      # DataMapper + _convert_value（从 renderer.py 提取）
├── generator.py                   # DocumentGenerator + 过滤器 + CLI（重命名自 TemplateRenderer）
├── schema.py                      # DataSchema + SchemaValidator（新增）
└── exceptions.py                  # 自定义异常类（新增）
```

### 命名变更

| 旧名 | 新名 |
|------|------|
| `renderer.py` | 删除（拆分） |
| `TemplateRenderer` | `DocumentGenerator` |
| `render_single()` | `generate()` |
| `python -m src.renderer` | `python -m src.generator` |
| `{%p if %}` | 不变 |

### 向后兼容

`src/__init__.py` 导出别名，确保外部导入不受影响：

```python
TemplateRenderer = DocumentGenerator
render_single = DocumentGenerator.generate
```

---

## 二、决策汇总

| 问题 | 决策 |
|------|------|
| Q1 日期过滤器 opt-in | 显式 `\| date` 才格式化 |
| Q2 日期输入格式 | 仅字符串，YYYY-MM-DD / YYYY/MM/DD |
| Q3 Type Hints 严格程度 | 宽松模式，public 方法强制，private 可选 |
| Q4 Schema 格式 | 字典/数据类描述（代码）+ JSON 文件（配置） |
| Q5 pytest 覆盖范围 | 先集成测试，后续按需补充单元测试 |
| Q6 语法检查报错方式 | logger 输出摘要 + 方法返回布尔 + strict 抛异常 |
| Q7 未使用字段报告 | 通过参数控制 |
| Q8 CLI 参数 | `--validate` + `--strict` 独立 |
| Q9 Schema 来源 | 独立 JSON 文件，`--schema schema.json` 指定 |
| Q10 目录结构 | 按职责垂直拆分 |
| Q11 改名 | `generator` 系列 |
| Q12 提交节奏 | 3 个提交（重构 1 + 功能 2） |

---

## 三、提交计划

### 提交1：重构（无功能变化）

```
refactor: 拆分 renderer.py 为 reader/mapper/generator 模块
```

包含：
- 新建 `src/reader.py`、`src/mapper.py`、`src/generator.py`、`src/exceptions.py`
- 删除 `src/renderer.py`
- `src/__init__.py` 导出公开 API + 别名
- 添加 public API Type Hints（宽松模式）
- CLI 入口从 `src.renderer` 改为 `src.generator`
- 原有测试通过

### 提交2：新功能（过滤器 + 语法检查）

```
feat: 新增 | date 过滤器、模板语法检查、--unused 参数
```

包含：
- `_filter_date()` 在 `generator.py` 中
- `DocumentGenerator.check_syntax()` 方法
- `--unused` CLI 参数（报告冗余数据字段）
- 日期过滤器集成测试
- 语法检查集成测试

### 提交3：Schema 验证器

```
feat: 新增 DataSchema 验证器 + --schema/--validate CLI 参数
```

包含：
- `src/schema.py`（`DataSchema` 数据类 + `SchemaValidator`）
- `SchemaValidator.validate()` 方法
- `--schema` 参数（JSON Schema 文件）
- `--validate` 参数（默认中断）
- `--strict-validate` 参数（抛异常）
- Schema 验证集成测试

---

## 四、实施步骤

### Task 1: 重构 — 拆分模块

**文件:**
- 创建: `src/reader.py`
- 创建: `src/mapper.py`
- 创建: `src/generator.py`
- 创建: `src/exceptions.py`
- 创建: `src/__init__.py`
- 删除: `src/renderer.py`
- 修改: `docs/开发手册.md`（模块名变更）
- 测试: `tests/test_report.py` 仍通过

---

### Task 2: 类型提示

**文件:**
- 修改: `src/reader.py`（public 方法加 Type Hints）
- 修改: `src/mapper.py`（public 方法加 Type Hints）
- 修改: `src/generator.py`（public 方法加 Type Hints）
- 修改: `src/schema.py`（新建时即加）

---

### Task 3: pytest 迁移

**文件:**
- 创建: `tests/conftest.py`（pytest fixtures）
- 重写: `tests/test_report.py`（用 pytest 函数式写法）

---

### Task 4: 日期过滤器

**文件:**
- 修改: `src/generator.py`（新增 `_filter_date`）
- 修改: `tests/test_report.py`（日期格式测试）
- 修改: `docs/开发手册.md`（过滤器列表）

---

### Task 5: 模板语法检查

**文件:**
- 修改: `src/generator.py`（新增 `check_syntax()` 方法）
- 修改: `src/generator.py`（新增 `--unused` CLI 参数）

---

### Task 6: Schema 验证器

**文件:**
- 创建: `src/schema.py`
- 修改: `src/generator.py`（`validate` 参数 + 调用验证器）
- 修改: `src/__init__.py`（导出 `DataSchema`/`SchemaValidator`）
- 创建: `examples/schema.json`（示例 Schema 文件）
- 测试: Schema 验证集成测试

---

## 五、自测清单

每个提交后运行：

```bash
.venv/Scripts/python.exe -m pytest tests/test_report.py -v
```

最终全部通过后：

```bash
git log --oneline -5
# 确认提交结构正确
```
