# 文档索引

本索引用于梳理仓库内的文档与说明文件，并标注哪些内容属于上游镜像，避免误编辑。

## 项目主文档
- `README.md`：项目入口与快速开始
- `PLAN.md`：架构设计、模块说明与实施状态
- `.env.example`：环境变量示例
- `examples/demo.py`：可运行的端到端示例
- `tests/`：单元测试与多模态测试用例

## Benchmark 与数据集（上游镜像）
- `SpreadsheetBench-NoDocker/README.md`：SpreadsheetBench 官方说明
- `SpreadsheetBench-NoDocker/README_LOCAL.md`：No-Docker 版本使用指南
- `SpreadsheetBench-NoDocker/CHANGELOG_CN.md`：变更记录（中文）
- `SpreadsheetBench-NoDocker/CLAUDE.md`：上游说明文件

## Skills（上游镜像 + 内置副本）
- `skills/xlsx/SKILL.md`：ExcelAgent 运行时使用的内置 xlsx skill
- `skills/xlsx/LICENSE.txt`：内置 skill 许可证
- `skill/skills/README.md`：Anthropic skills 仓库说明（上游镜像）
- `skill/skills/spec/agent-skills-spec.md`：Agent Skills 规范（上游镜像）
- `skill/skills/skills/xlsx/SKILL.md`：上游 xlsx skill 参考实现
- `skill/skills/skills/docx|pdf|pptx/`：上游文档类 skills 参考实现

## 维护约定
- `SpreadsheetBench-NoDocker/` 与 `skill/` 为上游镜像目录，默认不在此项目内改动。
- 需要定制或裁剪时，优先在项目主文档中补充说明，或将修改记录在本仓库的变更日志中。
