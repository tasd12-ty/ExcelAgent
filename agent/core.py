"""ExcelAgent 核心 — 智能体主循环"""

import json
import re

from llm.client import chat, DEFAULT_MODEL
from agent.dispatcher import dispatch, get_tools_description

SYSTEM_PROMPT = """你是 ExcelAgent，一个专业的 Excel 操作智能体。

你可以通过调用工具来读取、创建、编辑和分析 Excel 文件。

{tools}

## 调用格式

当你需要使用工具时，请输出如下 JSON 块:
```tool_call
{{"tool": "工具名", "args": {{"参数名": "参数值"}}}}
```

## 规则
1. 公式优先：始终使用 Excel 公式，不硬编码计算结果
2. 先读取再修改：修改文件前先读取了解结构
3. 一步一步来：每次只调用一个工具，等待结果后再决定下一步
4. 完成后说明：操作完成后告知用户结果
"""

TOOL_CALL_PATTERN = re.compile(r"```tool_call\s*\n(.+?)\n```", re.DOTALL)


def parse_tool_call(text: str) -> dict | None:
    """从 LLM 回复中解析工具调用。"""
    match = TOOL_CALL_PATTERN.search(text)
    if not match:
        return None
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError:
        return None


def run(user_input: str, file_path: str | None = None, max_steps: int = 10, model: str | None = None) -> str:
    """运行智能体处理用户请求。

    Args:
        user_input: 用户的自然语言请求
        file_path: 关联的 Excel 文件路径（可选）
        max_steps: 最大工具调用轮次
        model: LLM 模型ID

    Returns:
        最终回复文本
    """
    system = SYSTEM_PROMPT.format(tools=get_tools_description())

    messages = [{"role": "system", "content": system}]

    # 构建初始用户消息
    content = user_input
    if file_path:
        content += f"\n\n目标文件: {file_path}"
    messages.append({"role": "user", "content": content})

    for step in range(max_steps):
        response = chat(messages, model=model)
        messages.append({"role": "assistant", "content": response})

        # 尝试解析工具调用
        tool_call = parse_tool_call(response)
        if not tool_call:
            # 没有工具调用，智能体已完成
            return response

        # 执行工具
        tool_name = tool_call["tool"]
        args = tool_call.get("args", {})

        print(f"[Step {step + 1}] 调用工具: {tool_name}({args})")

        try:
            result = dispatch(tool_name, **args)
            # 将结果转为字符串
            if hasattr(result, "to_string"):
                result_str = result.to_string()
            elif isinstance(result, dict):
                result_str = json.dumps(result, ensure_ascii=False, default=str, indent=2)
            else:
                result_str = str(result)
        except Exception as e:
            result_str = f"错误: {e}"

        messages.append({"role": "user", "content": f"工具执行结果:\n{result_str}"})

    return messages[-1]["content"] if messages else "达到最大步数限制"
