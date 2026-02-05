# ExcelAgent - 设计与现状

## 项目目标

构建一个基于 Agent Skills 规范的智能体项目，支持通过工具与 Skill 对 Excel 文件进行读取、编辑、分析等操作，同时支持通过 OpenRouter（OpenAI 兼容格式）调用多模态大模型，并可在 SpreadsheetBench 上进行评测。

---

## 架构设计

```
┌─────────────────────────────────────────────┐
│                  ExcelAgent                 │
├─────────────┬───────────────┬───────────────┤
│  Skill 层   │   LLM 层      │   工具层      │
│             │               │               │
│ xlsx skill  │ OpenRouter    │ openpyxl      │
│ (内置副本)  │ (OpenAI格式)   │ pandas        │
│             │ 多模态支持     │ matplotlib    │
└─────────────┴───────────────┴───────────────┘
```

---

## 核心模块

| 模块 | 说明 |
|------|------|
| `agent/` | 智能体核心逻辑（主循环、工具调度） |
| `llm/` | OpenRouter/OpenAI 兼容客户端与多模态封装 |
| `tools/` | Excel 读写、格式化、分析、Python 执行工具 |
| `skills/` | 运行时使用的内置 skill（当前为 `xlsx`） |
| `skill/` | 上游 Anthropic skills 仓库镜像（参考资料） |
| `benchmark/` | SpreadsheetBench 运行与评测脚本 |
| `SpreadsheetBench-NoDocker/` | 数据集与评测依赖（上游仓库镜像） |
| `examples/` | 使用示例 |
| `tests/` | 单元测试 |

---

## Skill 与工具关系

- 运行时能力主要通过 `tools/` 提供（读写/格式化/分析/图表/代码执行）。
- `skills/xlsx/` 作为内置 Skill，用于公式重算等与 Office 生态配套的脚本。
- `skill/` 是完整的上游 skills 仓库镜像，仅作参考，不参与 ExcelAgent 运行。

---

## 项目结构

```
ExcelAgent/
├── README.md
├── PLAN.md
├── docs/
│   └── README.md
├── .env.example
├── pyproject.toml
├── agent/
│   ├── core.py
│   └── dispatcher.py
├── llm/
│   ├── client.py
│   └── multimodal.py
├── tools/
│   ├── reader.py
│   ├── writer.py
│   ├── formatter.py
│   ├── analyzer.py
│   └── code_executor.py
├── skills/
│   └── xlsx/
├── skill/
│   └── skills/...
├── benchmark/
│   ├── run_benchmark.py
│   └── evaluate.py
├── SpreadsheetBench-NoDocker/
├── examples/
│   └── demo.py
└── tests/
    ├── test_reader.py
    ├── test_writer.py
    └── test_llm.py
```

---

## 依赖

主要依赖（见 `pyproject.toml`）：
- openai (OpenRouter 兼容客户端)
- openpyxl (Excel 读写)
- pandas (数据分析)
- Pillow (图片处理)
- matplotlib (图表生成)
- python-dotenv (环境变量管理)
- tqdm (Benchmark 进度)

可选依赖：
- dev: pytest, pytest-asyncio
- office: defusedxml, lxml

---

## 实现状态

| 能力 | 状态 | 说明 |
|------|------|------|
| Excel 读取/写入 | 已完成 | `tools/reader.py`, `tools/writer.py` |
| 格式化 | 已完成 | `tools/formatter.py` |
| 分析/图表 | 已完成 | `tools/analyzer.py` |
| 代码执行 | 已完成 | `tools/code_executor.py` |
| LLM 调用 | 已完成 | `llm/client.py` |
| 多模态 | 已完成 | `llm/multimodal.py` |
| 智能体主循环 | 已完成 | `agent/core.py` |
| Benchmark 运行 | 已完成 | `benchmark/run_benchmark.py` |
| Benchmark 评测 | 已完成 | `benchmark/evaluate.py` |
| 示例/测试 | 已完成 | `examples/`, `tests/` |

---

## 测试方案

1. 单元测试：Excel 读写、格式化、公式计算
2. 集成测试：Skill 调度 + LLM 调用端到端流程
3. 多模态测试：生成图表 -> 截图 -> OpenRouter 多模态模型解读
4. SpreadsheetBench：`benchmark/` + `SpreadsheetBench-NoDocker/`

---

## 关键设计原则

- **公式优先**：始终使用 Excel 公式，不在 Python 中硬编码计算结果
- **Skill 规范兼容**：遵循 Anthropic Agent Skills 标准格式
- **LLM 可替换**：通过 OpenRouter 统一接口，可随时切换模型
- **渐进式复杂度**：从简单读写开始，逐步加入分析和多模态能力
