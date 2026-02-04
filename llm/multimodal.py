"""多模态 LLM 调用 — 图片理解"""

import base64
import os

from llm.client import chat, DEFAULT_MODEL


def encode_image(image_path: str) -> str:
    """将图片文件编码为 base64。"""
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def analyze_image(
    image_path: str,
    prompt: str = "请分析这张图表，描述主要趋势和关键数据点。",
    model: str | None = None,
) -> str:
    """发送图片给多模态模型进行分析。

    Args:
        image_path: 图片文件路径
        prompt: 分析提示词
        model: 模型ID
    """
    ext = os.path.splitext(image_path)[1].lower()
    mime_map = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}
    mime_type = mime_map.get(ext, "image/png")

    b64 = encode_image(image_path)

    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:{mime_type};base64,{b64}"},
                },
            ],
        }
    ]
    return chat(messages, model=model)


def analyze_excel_chart(
    image_path: str,
    context: str = "",
    model: str | None = None,
) -> str:
    """专门分析 Excel 图表截图。

    Args:
        image_path: 图表截图路径
        context: 额外上下文（如数据来源说明）
    """
    prompt = "你是一位数据分析专家。请分析这张 Excel 图表截图:\n"
    if context:
        prompt += f"\n背景信息: {context}\n"
    prompt += "\n请提供:\n1. 图表类型\n2. 主要趋势\n3. 关键数据点\n4. 分析建议"

    return analyze_image(image_path, prompt=prompt, model=model)
