# Auto-Report-Generation

Word 文档自动化生成系统 —— 将 Word 模板 + Excel/YAML 数据 → 填充好的 Word 文档。

## 文档导航

| 文档 | 读者 | 说明 |
|------|------|------|
| [01-快速入门](docs/01-快速入门.md) | 所有人 | 5 分钟从零到生成第一份文档 |
| [02-数据格式规范](docs/02-数据格式规范.md) | 所有人 | Excel/YAML 格式、变量、过滤器、嵌套 |
| [03-用户操作指南](docs/03-用户操作指南.md) | 业务人员 | 填写数据、执行命令、排查问题 |
| [04-模板制作规范](docs/04-模板制作规范.md) | AI / 模板制作者 | 报告样本→模板+数据字典 |
| [05-开发者手册](docs/05-开发者手册.md) | 开发者 | 架构、API、扩展、部署 |
| [06-附录-速查表](docs/06-附录-速查表.md) | 所有人 | 语法规格一览 |

## 快速开始

```bash
pip install -r requirements.txt
python -m src -d data.xlsx -t template.docx -o output.docx
```

## 项目结构

```
src/
├── orchestrator.py     # 编排：读取→校验→渲染
├── cli.py              # CLI 入口
├── path_guard.py       # 路径安全校验
├── logging_config.py   # 日志配置
├── exceptions.py       # 自定义异常
├── reader/             # 数据读取（Excel / YAML）
├── processing/         # 数据映射 + 校验
└── render/             # 模板渲染 + 过滤器
```

## 测试

```bash
pytest tests/ -q
```
