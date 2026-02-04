"""OpenRouter LLM 客户端（OpenAI 兼容格式）"""

import os

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

DEFAULT_MODEL = os.getenv("OPENROUTER_MODEL", "google/gemini-3-flash-preview")
BASE_URL = os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")
API_KEY = os.getenv("OPENROUTER_API_KEY", "")


def get_client() -> OpenAI:
    """获取 OpenRouter 客户端实例。"""
    if not API_KEY:
        raise ValueError("请设置 OPENROUTER_API_KEY 环境变量")
    return OpenAI(base_url=BASE_URL, api_key=API_KEY)


def chat(
    messages: list[dict],
    model: str | None = None,
    temperature: float = 0.7,
    max_tokens: int = 4096,
) -> str:
    """发送聊天请求，返回回复文本。

    Args:
        messages: [{"role": "user", "content": "..."}]
        model: 模型ID，默认使用环境变量配置
    """
    client = get_client()
    response = client.chat.completions.create(
        model=model or DEFAULT_MODEL,
        messages=messages,
        temperature=temperature,
        max_tokens=max_tokens,
    )
    return response.choices[0].message.content


def chat_stream(
    messages: list[dict],
    model: str | None = None,
    temperature: float = 0.7,
    max_tokens: int = 4096,
):
    """流式聊天，yield 每个文本片段。"""
    client = get_client()
    stream = client.chat.completions.create(
        model=model or DEFAULT_MODEL,
        messages=messages,
        temperature=temperature,
        max_tokens=max_tokens,
        stream=True,
    )
    for chunk in stream:
        if chunk.choices and chunk.choices[0].delta.content:
            yield chunk.choices[0].delta.content
