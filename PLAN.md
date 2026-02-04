# ExcelAgent - 智能Excel操作代理方案

## 项目目标

构建一个基于 Agent Skills 规范的智能体项目，支持通过 Skill 对 Excel 文件进行读取、编辑、分析等操作，同时支持通过 OpenRouter（OpenAI 兼容格式）调用多模态大模型。

---

## 架构设计

```
┌─────────────────────────────────────────────┐
│                  ExcelAgent                  │
├─────────────┬───────────────┬───────────────┤
│  Skill 层   │   LLM 层      │   工具层       │
│             │               │               │
│ xlsx skill  │ OpenRouter    │ openpyxl      │
│ (Anthropic) │ (OpenAI格式)   │ pandas        │
│             │               │ Pillow        │
│ 自定义skill  │ 多模态支持     │ matplotlib    │
└─────────────┴───────────────┴───────────────┘
```

### 核心模块

| 模块 | 说明 |
|------|------|
| `skills/` | Skill 定义（基于 Anthropic skills 规范） |
| `agent/` | 智能体核心逻辑，调度 Skill 和 LLM |
| `llm/` | LLM 客户端，封装 OpenRouter 调用 |
| `tools/` | Excel 操作工具集（读写、格式化、公式） |
| `tests/` | 测试用例 |

---

## Skill 设计

### 1. excel-reader（Excel 读取）

基于 Anthropic xlsx skill 规范，支持：
- 读取 .xlsx / .csv / .tsv 文件
- 提取表结构、数据摘要
- 识别公式和格式信息

### 2. excel-writer（Excel 写入/编辑）

- 创建新 Excel 文件
- 修改现有文件（单元格、公式、样式）
- 遵循金融模型色彩规范（蓝色输入、黑色公式等）
- 使用 Excel 公式而非硬编码计算值

### 3. excel-analyzer（Excel 分析）

- 数据统计分析
- 图表生成
- 调用多模态 LLM 对图表截图进行解读

---

## LLM 集成方案

### OpenRouter 调用（OpenAI 兼容格式）

```python
# 配置示例
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
OPENROUTER_API_KEY = "sk-or-..."

# 支持的模型（通过 OpenRouter）
# - google/gemini-3-flash-preview
```

使用 `openai` Python SDK，仅修改 `base_url` 和 `api_key` 即可接入 OpenRouter：

```python
from openai import OpenAI

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=OPENROUTER_API_KEY,
)
```

多模态调用：将 Excel 图表截图以 base64 传入 `image_url`，让模型解读。

---

## 项目结构

```
ExcelAgent/
├── PLAN.md                # 本文档
├── CLAUDE.md              # Claude Code 项目配置
├── pyproject.toml         # 项目依赖
├── .env.example           # 环境变量模板
├── skills/
│   └── xlsx/
│       ├── SKILL.md       # Skill 定义（基于 Anthropic 规范）
│       └── scripts/
│           └── recalc.py  # 公式重算脚本
├── agent/
│   ├── __init__.py
│   ├── core.py            # 智能体主循环
│   └── dispatcher.py      # Skill 调度器
├── llm/
│   ├── __init__.py
│   ├── client.py          # OpenRouter 客户端
│   └── multimodal.py      # 多模态（图片）处理
├── tools/
│   ├── __init__.py
│   ├── reader.py          # Excel 读取
│   ├── writer.py          # Excel 写入
│   ├── analyzer.py        # 数据分析
│   └── formatter.py       # 格式化工具
├── tests/
│   ├── test_reader.py
│   ├── test_writer.py
│   ├── test_llm.py
│   └── fixtures/
│       └── sample.xlsx    # 测试用 Excel 文件
└── examples/
    └── demo.py            # 使用示例
```

---

## 依赖

```toml
[project]
dependencies = [
    "openai>=1.0",        # OpenRouter 兼容客户端
    "openpyxl>=3.1",      # Excel 读写（格式、公式）
    "pandas>=2.0",        # 数据分析
    "Pillow>=10.0",       # 图片处理（截图）
    "matplotlib>=3.7",    # 图表生成
    "python-dotenv>=1.0", # 环境变量管理
]
```

---

## 实施步骤

| 步骤 | 内容 | 产出 |
|------|------|------|
| 1 | 初始化项目结构、依赖 | pyproject.toml, .env |
| 2 | 实现 Excel 读写工具 (tools/) | reader.py, writer.py |
| 3 | 定义 xlsx Skill (skills/) | SKILL.md |
| 4 | 封装 OpenRouter LLM 客户端 | client.py, multimodal.py |
| 5 | 实现智能体核心逻辑 | core.py, dispatcher.py |
| 6 | 数据分析+图表+多模态解读 | analyzer.py |
| 7 | 编写测试 | tests/ |
| 8 | 编写使用示例 | demo.py |

---

## 测试方案

1. **单元测试**：Excel 读写、格式化、公式计算
2. **集成测试**：Skill 调度 + LLM 调用端到端流程
3. **多模态测试**：生成图表 -> 截图 -> OpenRouter 多模态模型解读
4. **使用 Anthropic xlsx Skill**：验证与官方 Skill 规范的兼容性

---

## 关键设计原则

- **公式优先**：始终使用 Excel 公式，不在 Python 中硬编码计算结果
- **Skill 规范兼容**：遵循 Anthropic Agent Skills 标准格式
- **LLM 可替换**：通过 OpenRouter 统一接口，可随时切换模型
- **渐进式复杂度**：从简单读写开始，逐步加入分析和多模态能力
