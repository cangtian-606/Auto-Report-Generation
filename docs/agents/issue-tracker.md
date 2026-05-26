# Issue Tracker: Local Markdown

Issues 和 PRD 以 Markdown 文件存放在 `.scratch/` 目录下。

## 约定

- 每个功能一个目录：`.scratch/<feature-slug>/`
- PRD 文件：`.scratch/<feature-slug>/PRD.md`
- 实施 issue：`.scratch/<feature-slug>/issues/<NN>-<slug>.md`，从 `01` 开始编号
- 分诊状态记录在 issue 文件顶部的 `Status:` 行中（标签名称参见 `triage-labels.md`）
- 评论和对话历史追加到文件底部 `## Comments` 标题下

## 当技能说"发布到 issue tracker"

在 `.scratch/<feature-slug>/` 下创建新文件（如有必要先创建目录）。

## 当技能说"获取相关 issue"

读取指定路径的文件。用户通常会直接传递路径或 issue 编号。
