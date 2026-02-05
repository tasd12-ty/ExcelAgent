# ExcelAgent

ExcelAgent 是一个面向 Excel 文件读写、分析与图表生成的智能体项目，采用“Agent + 工具 + LLM”的分层设计，支持 OpenRouter（OpenAI 兼容格式）调用多模态模型，并内置 SpreadsheetBench 评测流程。

## 主要能力
- 读取/写入 Excel（含公式、格式、样式）
- 数据分析与图表生成
- 多模态图表解读
- 基于工具调用的智能体交互
- SpreadsheetBench 基准测试与评测

## 快速开始
1. 安装依赖：`pip install -e .`
2. 配置环境变量：至少设置 `OPENROUTER_API_KEY`，其它可选项见 `.env.example`
3. 运行示例：`python examples/demo.py`
4. 运行测试：`pytest`（需要 `dev` 依赖）

## 项目结构
- `agent/`：智能体主循环与工具调度
- `llm/`：OpenRouter/OpenAI 兼容客户端与多模态封装
- `tools/`：Excel 读写、格式化、分析与 Python 执行工具
- `skills/xlsx/`：内置 xlsx skill（可用于公式重算等）
- `skill/`：上游 Anthropic skills 仓库镜像（参考资料，不参与运行）
- `benchmark/`：SpreadsheetBench 运行与评测脚本
- `SpreadsheetBench-NoDocker/`：数据集与评测依赖（上游仓库镜像）
- `examples/`：功能演示脚本
- `tests/`：单元测试
- `PLAN.md`：设计与里程碑
- `docs/README.md`：文档索引

## 架构设计
```
┌─────────────────────────────────────────────┐
│                  ExcelAgent                 │
├─────────────┬───────────────┬───────────────┤
│   Skill层   │    LLM层      │    工具层     │
│             │               │               │
│ xlsx skill  │ OpenRouter    │ openpyxl      │
│ (内置)      │ (OpenAI格式)   │ pandas        │
│             │ 多模态支持     │ matplotlib    │
└─────────────┴───────────────┴───────────────┘
```

- Agent 核心：`agent/core.py` 负责对话、解析 `tool_call`、执行工具并回传结果。
- 工具调度：`agent/dispatcher.py` 统一注册工具与参数描述，便于扩展。
- 工具实现：`tools/` 负责 Excel 读写、格式化、分析与图表生成。

## 官方 xlsx skill 迁移说明
官方 xlsx skill 原始版本位于 `skill/skills/skills/xlsx/`，已迁移到项目可直接使用的 `skills/xlsx/`：
- `skills/xlsx/SKILL.md`
- `skills/xlsx/LICENSE.txt`
- `skills/xlsx/scripts/`（包含 `recalc.py` 与 office 相关脚本）

该副本作为 ExcelAgent 的内置 skill 目录，便于在项目内直接引用与扩展。

## 常见工作流（公式优先）
1. 读取/分析：优先使用 `tools/reader.py`（pandas 读取与摘要）
2. 修改/写入：使用 `tools/writer.py` 写入数据与公式
3. 格式化：使用 `tools/formatter.py`
4. 公式重算：使用 `skills/xlsx/scripts/recalc.py`
   ```bash
   python skills/xlsx/scripts/recalc.py <excel_file> [timeout_seconds]
   ```

## 环境变量
- `OPENROUTER_API_KEY`：OpenRouter API Key（必填）
- `OPENROUTER_BASE_URL`：默认 `https://openrouter.ai/api/v1`
- `OPENROUTER_MODEL`：默认 `google/gemini-3-flash-preview`
- `LLM_MAX_TOKENS`：默认 `4096`

## Benchmark（SpreadsheetBench）
数据目录：`SpreadsheetBench-NoDocker/data/<dataset>/dataset.json`

运行基准测试：
```bash
python benchmark/run_benchmark.py --dataset sample_data_200 --model <model-id> --max_steps 15
```

评测结果：
```bash
python benchmark/evaluate.py --dataset sample_data_200 --model <model-id> --setting agent
```

Benchmark 逻辑概述：
- 对每个任务运行 3 个 test case（输入复制到输出后运行 agent）
- 记录对话与工具调用到 `benchmark/logs/`
- 输出文件保存到 `SpreadsheetBench-NoDocker/data/<dataset>/outputs/`

## 文档索引
完整文档索引与上游文档说明见 `docs/README.md`。
