"""ExcelAgent 使用示例

运行前请确保:
1. pip install -e .
2. 复制 .env.example 为 .env 并填入 OPENROUTER_API_KEY

用法:
    python examples/demo.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from dotenv import load_dotenv
load_dotenv()

from tools.writer import create_workbook, write_cells
from tools.reader import get_summary
from tools.analyzer import analyze_data, create_chart
from tools.formatter import apply_header_style, auto_fit_columns, add_table_borders


def demo_create_and_read():
    """演示: 创建 Excel 并读取摘要"""
    print("=" * 50)
    print("Demo 1: 创建 Excel 文件")
    print("=" * 50)

    output = "examples/demo_output.xlsx"

    create_workbook(output, {
        "销售数据": [
            ["产品", "Q1", "Q2", "Q3", "Q4"],
            ["笔记本电脑", 150, 200, 180, 250],
            ["手机", 300, 350, 280, 400],
            ["平板", 80, 100, 90, 120],
        ]
    })

    # 添加公式行
    write_cells(output, "销售数据", [
        {"cell": "A5", "value": "合计", "style": "formula"},
        {"cell": "B5", "value": "=SUM(B2:B4)", "style": "formula"},
        {"cell": "C5", "value": "=SUM(C2:C4)", "style": "formula"},
        {"cell": "D5", "value": "=SUM(D2:D4)", "style": "formula"},
        {"cell": "E5", "value": "=SUM(E2:E4)", "style": "formula"},
    ])

    # 格式化
    apply_header_style(output, "销售数据")
    auto_fit_columns(output, "销售数据")
    add_table_borders(output, "销售数据", "A1:E5")

    print(f"文件已创建: {output}\n")

    # 读取摘要
    summary = get_summary(output)
    print("文件摘要:")
    print(summary)
    return output


def demo_analyze_and_chart(file_path: str):
    """演示: 数据分析与图表生成"""
    print("\n" + "=" * 50)
    print("Demo 2: 数据分析与图表")
    print("=" * 50)

    result = analyze_data(file_path, "销售数据")
    print(f"数据维度: {result['shape']}")
    print(f"列: {result['columns']}")
    print(f"\n统计摘要:\n{result['describe']}")

    chart_path = create_chart(
        file_path, "销售数据",
        chart_type="bar",
        x_col="产品",
        y_col="Q1",
        title="Q1 产品销售对比",
    )
    print(f"\n图表已保存: {chart_path}")
    return chart_path


def demo_multimodal(chart_path: str):
    """演示: 多模态 LLM 图表解读"""
    print("\n" + "=" * 50)
    print("Demo 3: 多模态 LLM 图表解读")
    print("=" * 50)

    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        print("跳过: 未设置 OPENROUTER_API_KEY")
        return

    from llm.multimodal import analyze_excel_chart
    result = analyze_excel_chart(chart_path, context="这是一份产品Q1销售数据的柱状图")
    print(f"\nLLM 分析结果:\n{result}")


def demo_agent():
    """演示: 使用智能体自然语言交互"""
    print("\n" + "=" * 50)
    print("Demo 4: 智能体自然语言交互")
    print("=" * 50)

    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        print("跳过: 未设置 OPENROUTER_API_KEY")
        return

    from agent.core import run
    result = run(
        "请读取这个 Excel 文件，告诉我哪个产品在Q1表现最好。",
        file_path="examples/demo_output.xlsx",
    )
    print(f"\n智能体回复:\n{result}")


if __name__ == "__main__":
    file_path = demo_create_and_read()
    chart_path = demo_analyze_and_chart(file_path)
    demo_multimodal(chart_path)
    demo_agent()
