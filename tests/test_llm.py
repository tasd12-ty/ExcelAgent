"""测试 LLM 客户端 — 需要有效的 OPENROUTER_API_KEY"""

import os
import pytest
from dotenv import load_dotenv

load_dotenv()

# 如果没有 API key 则跳过
pytestmark = pytest.mark.skipif(
    not os.getenv("OPENROUTER_API_KEY"),
    reason="需要设置 OPENROUTER_API_KEY"
)


def test_chat():
    from llm.client import chat
    response = chat([{"role": "user", "content": "回复 'OK'，不要说其他内容。"}])
    assert response is not None
    assert len(response) > 0


def test_chat_stream():
    from llm.client import chat_stream
    chunks = list(chat_stream([{"role": "user", "content": "回复 'OK'，不要说其他内容。"}]))
    assert len(chunks) > 0
    full_text = "".join(chunks)
    assert len(full_text) > 0


def test_multimodal():
    """测试多模态调用（需要先生成图表）"""
    from tools.writer import create_workbook
    from tools.analyzer import create_chart
    from llm.multimodal import analyze_image

    fixtures = os.path.join(os.path.dirname(__file__), "fixtures")
    os.makedirs(fixtures, exist_ok=True)
    xlsx_path = os.path.join(fixtures, "mm_test.xlsx")
    chart_path = os.path.join(fixtures, "mm_chart.png")

    try:
        create_workbook(xlsx_path, {
            "Data": [
                ["Month", "Sales"],
                ["Jan", 100],
                ["Feb", 150],
                ["Mar", 200],
                ["Apr", 180],
            ]
        })
        create_chart(xlsx_path, "Data", chart_type="bar", x_col="Month", y_col="Sales", output_path=chart_path)
        assert os.path.exists(chart_path)

        result = analyze_image(chart_path, "用一句话描述这张图表。")
        assert result is not None
        assert len(result) > 0
    finally:
        for p in [xlsx_path, chart_path]:
            if os.path.exists(p):
                os.remove(p)
