# CONTEXT.md — 项目领域模型

## 核心概念

| 术语 | 定义 |
|------|------|
| 模板 | 含 Jinja2 变量占位符的 .docx 文件 |
| 数据文件 | 提供渲染数据的 .xlsx 或 .yaml 文件 |
| Sheet | Excel 中的一个工作表，对应渲染上下文的一个顶级键 |
| context | 传给 Jinja2 模板渲染的嵌套 dict |
| 键值对 Sheet | 首列 `字段编码` 的 Sheet，每个字段映射为一个值 |
| 表格 Sheet | 首行是列标题的 Sheet，每行是一条数据记录 |
| DocxEditor | zip 解/打包上下文管理器，预处理器通过它操作 docx XML |
| RenderContext | 将 DataMapper + 扁平化 + Schema 校验聚合为深度模块 |
| TemplateAnalyzer | 模板自省：表格数/标签数统计、变量检查、未使用数据检测 |
| RenderPipeline | 阶段序列执行器，统一管理临时文件生命周期 |

## 领域规则

- 所有 Sheet 名、字段名使用中文，见名知义
- 嵌套 Sheet 用 `.` 分隔父表和子表（如 `历史沿革.股权结构`）
- 布尔值用 TRUE/FALSE 表示，不区分大小写
- 过滤器 `\| money` / `\| percent` 对 0 返回空字符串
- 预处理器使用 DocxEditor 管理 docx I/O，不直接操作 zip/临时文件
- 运行自动生成日志到 `logs/`，每次执行有独立 run_id

## 已记录决策

参见 [docs/adr/](docs/adr/) 目录。
