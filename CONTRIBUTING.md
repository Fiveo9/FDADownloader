# Contributing

感谢你愿意改进这个项目。

这个仓库的目标很务实：稳定地下载 FDA guidance metadata、保存原始文件，并按可配置规则整理成本地资料库。欢迎围绕这个目标提交改进。

## 提交前先确认

- 先搜索现有 issue，避免重复反馈。
- 如果是 bug，请尽量附上运行环境、复现步骤、日志片段和预期行为。
- 如果是功能建议，请说明使用场景，以及为什么它适合放进这个仓库的默认工作流。

## 本地开发

1. 安装 Python 3.8+。
2. 安装项目依赖：

```bash
pip install -r requirements.txt
```

3. 运行测试：

```bash
python -m unittest discover -s tests -q
```

## Pull Request 指南

- 保持改动聚焦，不把无关重构混进同一个 PR。
- 修改行为时，请同步更新 README 或相关文档。
- 修改分类逻辑时，请同时更新 `classification_rules.csv`、默认行为说明和必要测试。
- 不要提交本地运行产物，例如：
  - `FDA_Downloads/`
  - `FDA_Guidance_Library/`
  - `FDA_Guidance_Data_*.xlsx`
- 如果改动影响下载或整理流程，请说明你是如何验证的。

## 代码风格

- 优先保持脚本可读、直接、易维护。
- 除非确有必要，否则不要引入重量级新依赖。
- 优先复用现有命名和工作流，避免把仓库变成一个框架。

## 沟通方式

欢迎提出小修小补，也欢迎提出结构性改进建议。只要问题定义清楚，我们就能比较高效地把它落下来。
