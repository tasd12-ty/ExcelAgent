"""Skill 调度器 — 根据用户意图分发到对应工具函数"""

from tools import reader, writer, formatter, analyzer


# 注册可用的工具函数
TOOL_REGISTRY = {
    "read_excel": {
        "fn": reader.read_excel,
        "description": "读取Excel文件内容",
        "params": ["file_path", "sheet_name"],
    },
    "read_formulas": {
        "fn": reader.read_excel_formulas,
        "description": "读取Excel中的公式",
        "params": ["file_path", "sheet_name"],
    },
    "get_summary": {
        "fn": reader.get_summary,
        "description": "生成Excel文件摘要",
        "params": ["file_path"],
    },
    "create_workbook": {
        "fn": writer.create_workbook,
        "description": "创建新Excel文件",
        "params": ["file_path", "sheets"],
    },
    "write_cells": {
        "fn": writer.write_cells,
        "description": "写入单元格数据",
        "params": ["file_path", "sheet_name", "cells"],
    },
    "apply_number_format": {
        "fn": writer.apply_number_format,
        "description": "应用数字格式",
        "params": ["file_path", "sheet_name", "formats"],
    },
    "auto_fit": {
        "fn": formatter.auto_fit_columns,
        "description": "自动调整列宽",
        "params": ["file_path", "sheet_name"],
    },
    "apply_header_style": {
        "fn": formatter.apply_header_style,
        "description": "应用表头样式",
        "params": ["file_path", "sheet_name"],
    },
    "add_borders": {
        "fn": formatter.add_table_borders,
        "description": "添加表格边框",
        "params": ["file_path", "sheet_name", "cell_range"],
    },
    "analyze_data": {
        "fn": analyzer.analyze_data,
        "description": "统计分析Excel数据",
        "params": ["file_path", "sheet_name"],
    },
    "create_chart": {
        "fn": analyzer.create_chart,
        "description": "生成图表",
        "params": ["file_path", "sheet_name", "chart_type", "x_col", "y_col", "output_path"],
    },
}


def get_tools_description() -> str:
    """生成工具列表描述，用于 LLM system prompt。"""
    lines = ["可用工具:"]
    for name, info in TOOL_REGISTRY.items():
        params = ", ".join(info["params"])
        lines.append(f"- {name}({params}): {info['description']}")
    return "\n".join(lines)


def dispatch(tool_name: str, **kwargs):
    """执行指定工具。"""
    if tool_name not in TOOL_REGISTRY:
        raise ValueError(f"未知工具: {tool_name}，可用: {list(TOOL_REGISTRY.keys())}")

    fn = TOOL_REGISTRY[tool_name]["fn"]
    return fn(**kwargs)
