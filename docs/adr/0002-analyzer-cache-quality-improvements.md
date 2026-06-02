# ADR-0002: 分析器、缓存与 DocxEditor 质量改进

## 日期

2026-06-01

## 状态

已提出（等待用户确认）

## 背景

2026-05-31 的代码审查（参见对话上下文）发现以下问题，按修复优先级排序：

### P0 真实 Bug

- **`analyzer.get_unused` 误报**：`--report-unused` 把模板里 `{% for item in 列表 %}{{ item.x }}` 实际使用的字段全部报为未使用。原因是函数把列表下标 `[0]` 拼进路径（如 `产销情况.月度产销[0].月份`），而 Jinja2 模板中"裸"路径是 `产销情况.月度产销` 或 `item.月份`，不包含 `[0]`。实测对 demo 数据误报 **315 个**未使用字段。

### P1 缺失的测试与文档

- **`analyzer.get_unused` 缺乏测试覆盖**——这正是它一直没被发现的原因（`tests/test_generator.py:41` 的 `test_find_unused_data` 没有覆盖 `{% for %}` 场景）。
- **测试计数文档过时**：`AGENTS.md` 与 `docs/adr/0001-deep-module-extraction.md` 都写"107 个测试"，实际跑出来是 **121 个**。

### P2 可维护性改进

- **模板缓存无法感知文件变更**：`DocumentGenerator._template_cache` 用路径字符串作 key，模板文件被外部修改后仍返回旧对象。一次性 CLI 调用无影响，但若作为库被 watch/serve 场景调用会出错。
- **`DocxEditor.__exit__` 状态不一致**：`__exit__` 后 `_tmp_dir` 未置 `None`，`save_to` 的 assert 文案声称"with 块内或退出后"都能调用，但 `os.walk` 在已删除目录上会静默拿不到任何文件，错误位置不在 assert。

### P3 范围外（暂不处理，未来 ADR 评估）

- `analyzer.get_undeclared` 吞掉异常，调用方难以区分"无未定义"与"检查失败"
- `YamlDataReader` 延后导入 yaml 的写法与项目其他模块不一致
- `generator._extract_error_info` 解析 Jinja2 错误消息过于脆弱
- `mapper.py` 排序对同 `.` 数 Sheet 隐式依赖输入顺序

## 决策

### D1：修复 `get_unused` 误报（红绿重构）

**核心思路**：list 内部字段是否被使用需要理解 Jinja2 的循环变量作用域，纯正则无法做到。改为**不下钻 list 内部**：list 整体被模板引用就视为"已用"，不再逐字段报告。

**行为变更**：
- 修复前：list 内每个字段都生成形如 `A.B[0].field` 的路径，命中率极低 → 全部报为未用
- 修复后：list 整体 `A.B` 是否在模板中引用，是→视整个 list 为已用（不再展开其内字段），否→把 `A.B` 报为未用

**取舍**：放弃"识别 list 内具体哪个字段未被使用"的精度（这个能力本就不准确），换取"误报清零"。

### D2：为 `get_unused` 增加测试

在 `tests/test_generator.py` 补两个测试用例：
- 含 `{% for %}` 块引用的列表字段不应被误报为未使用
- 真正未被引用的 list 整体应被报为未用

并保留现有 `test_find_unused_data`（它覆盖 dict 路径场景，必须继续通过）。

### D3：模板缓存引入 mtime 感知

`_template_cache` 的 key 从纯路径改为 `(path, mtime_ns)` 元组。文件 mtime 变化时自动重建。CLI 单次调用场景缓存只增一个条目，开销可忽略。

```python
_template_cache: Dict[Tuple[str, int], "DocxTemplate"] = {}

def _load_template(self) -> "DocxTemplate":
    mtime = os.stat(self.template_path).st_mtime_ns
    key = (self.template_path, mtime)
    if key not in self._template_cache:
        self._template_cache[key] = DocxTemplate(self.template_path)
    return self._template_cache[key]
```

### D4：`DocxEditor.__exit__` 状态置空

```python
def __exit__(self, exc_type, exc_val, exc_tb) -> None:
    if self._tmp_dir:
        shutil.rmtree(self._tmp_dir, ignore_errors=True)
        self._tmp_dir = None  # 新增
    return None
```

`save_to` 的 assert 文案简化为"只能在 with 块内调用"，删除"退出后"语义。

### D5：同步测试计数文档

更新 `AGENTS.md` 与 `docs/adr/0001-deep-module-extraction.md` 中的"107 个测试"为 "121 个"（D2 完成后再 +2 测试，最终为 123 个）。

## 实施策略

按用户工作规则采用 **TDD 与小步提交**：

| 提交 | 内容 | 顺序 |
|------|------|------|
| 1 | D2 测试先行（红）→ D1 修复（绿）→ 重构 | 一个原子提交 |
| 2 | D3 缓存 mtime 改造 | 含新测试 |
| 3 | D4 DocxEditor 状态清理 | 含新测试 |
| 4 | D5 文档同步 | 纯文本 |

每次提交后运行 `pytest`，确保 121 + 新增测试全绿。

**回退方案**：若 D1 的"不下钻 list"被业务方反对（确实有人想看 list 内具体未用字段），降级方案为"只报告 list 顶层路径"，D2 测试相应调整。

## 后果

### 正面

- `--report-unused` 输出可被业务人员信任，去除 315 条误报噪音
- 模板缓存对开发态热更新友好（修模板后无需重启进程）
- `DocxEditor` 状态机清晰，with 块外调用会立即抛明确异常
- 测试覆盖补齐，防止 `get_unused` 类问题回归
- 测试计数文档与现实一致

### 负面

- 失去"识别 list 内具体未用字段"的能力（D1 之前也不准确，D1 之后是"承认不能"）
- 模板缓存多一次 `os.stat` 调用（< 1ms，可忽略）
- `__exit__` 后不能再调用 `save_to`（行为变更，需在 [docs/05-开发者手册.md](file:///d:/ProgramWorkingSpace/Auto-Report-Generation/docs/05-开发者手册.md) 中说明）

### 风险

- mtime 感知缓存：NTFS 上 mtime 精度通常为 100ns，不影响；Linux ext4 为 ns 级，无问题
- 业务方对"不再下钻 list"的接受度：若反馈负面，按回退方案调整

## 不采纳的方案

- **重写 `get_unused` 为基于 docxtpl 的 AST 解析**：docxtpl 不暴露完整 Jinja2 AST，工作量与风险不成正比
- **直接删除 `--report-unused` 功能**：解决了误报但失去价值
- **D6-D9 一起打包修复**：单次 ADR 范围过大不利于回退，按用户"小步重构"规则拆分

## 关联文件

- 实施代码：`src/render/analyzer.py`、`src/render/generator.py`、`src/processing/docx_editor.py`
- 测试文件：`tests/test_generator.py`
- 文档：`AGENTS.md`、`docs/adr/0001-deep-module-extraction.md`
