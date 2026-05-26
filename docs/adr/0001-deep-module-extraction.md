# ADR-0001: 深度模块提炼

## 日期

2026-05-27

## 状态

已采纳

## 背景

v1.0 版代码库存在以下架构摩擦：

1. **orchestrator.py** 承担过多职责（215 行），包含上下文构建、扁平化、校验、模板统计、临时文件管理、汇总输出等分散逻辑
2. **table_preprocessors.py** 中 Tc/Tr 预处理器的 `process()` 各自管理 zip 解压/打包，`_repack()` 存在跨类调用（Tr 调用 `TcInheritancePreprocessor._repack`）
3. **generator.py** 中 `render()`、`get_undeclared_variables()`、`check_syntax()` 各自重复创建 `SandboxedEnvironment` 并注册过滤器
4. **logging_config.py** 使用模块级全局变量 `_run_id` / `_stage_times`，通过 loose functions 访问

## 决策

提炼四个深度模块：

### 1. DocxEditor（`src/processing/docx_editor.py`）

zip 解包/打包封装为上下文管理器，预处理器专注于 XML 变换。

```python
with DocxEditor("template.docx") as editor:
    root = editor.root
    # XML 变换...
    editor.save_to("output.docx", modified=True)
```

### 2. RenderPipeline（`src/render/pipeline.py`）

阶段序列执行器，统一管理临时文件生命周期。

```python
pipeline = RenderPipeline()
pipeline.add_stage("读数据", read_fn)
pipeline.add_stage("构建", build_fn)
state = pipeline.run()
pipeline.cleanup_temps()
```

### 3. RenderContext（`src/render/context.py`）

DataMapper + 扁平化 + Schema 校验聚合。

```python
rctx = RenderContext(raw_data)
context = rctx.build()
errors = rctx.validate(schema_path)
```

### 4. TemplateAnalyzer（`src/render/analyzer.py`）

模板自省聚合：stats、变量检查、未使用数据检测。SandboxedEnvironment 只创建一次。

```python
analyzer = TemplateAnalyzer("template.docx")
tables, tags = analyzer.get_stats()
undeclared = analyzer.get_undeclared(context)
```

## 后果

- **正面**: orchestrator 从 215 行减至 ~180 行，各模块职责清晰；消除跨类耦合和重复代码
- **负面**: 新增 4 个模块文件（12 个源文件总计）；P2 候选（统一 DataMapper 挂载策略、LogTracer）被评估为风险大于收益，暂不实施
- **风险**: 无，107 个测试全部通过，冒烟测试渲染成功

## 未采纳的替代方案

- **LogTracer**: 需修改全局日志系统状态管理，涉及多模块耦合，收益不抵风险
- **统一 DataMapper 挂载**: 挂载逻辑涉及 6 条分支路径，深层重构风险高，需专门设立测试保障后再进行
