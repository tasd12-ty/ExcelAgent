"""Excel 数据分析与图表工具"""

import os

import matplotlib
matplotlib.use("Agg")  # 非交互式后端
import matplotlib.pyplot as plt
import pandas as pd

from tools.reader import read_excel


def analyze_data(file_path: str, sheet_name: str | None = None) -> dict:
    """对 Excel 数据进行统计分析。

    Returns:
        {
            "shape": (rows, cols),
            "columns": [...],
            "dtypes": {...},
            "describe": 统计摘要文本,
            "missing": {col: count},
        }
    """
    info = read_excel(file_path, sheet_name)
    target = sheet_name or info["active_sheet"]
    df = info["data"][target]

    missing = {col: int(df[col].isna().sum()) for col in df.columns if df[col].isna().any()}

    return {
        "shape": df.shape,
        "columns": list(df.columns),
        "dtypes": {str(k): str(v) for k, v in df.dtypes.items()},
        "describe": df.describe(include="all").to_string(),
        "missing": missing,
    }


def create_chart(
    file_path: str,
    sheet_name: str | None = None,
    chart_type: str = "bar",
    x_col: str | None = None,
    y_col: str | None = None,
    output_path: str | None = None,
    title: str = "",
) -> str:
    """从 Excel 数据生成图表并保存为图片。

    Args:
        chart_type: bar, line, pie, scatter, hist
        x_col: X轴列名
        y_col: Y轴列名
        output_path: 图片保存路径，默认同目录下 chart.png

    Returns:
        生成的图片路径
    """
    info = read_excel(file_path, sheet_name)
    target = sheet_name or info["active_sheet"]
    df = info["data"][target]

    if not output_path:
        base = os.path.splitext(file_path)[0]
        output_path = f"{base}_chart.png"

    plt.rcParams["font.sans-serif"] = ["Arial Unicode MS", "SimHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(10, 6))

    if chart_type == "bar":
        if x_col and y_col:
            df.plot.bar(x=x_col, y=y_col, ax=ax)
        else:
            df.select_dtypes(include="number").plot.bar(ax=ax)
    elif chart_type == "line":
        if x_col and y_col:
            df.plot.line(x=x_col, y=y_col, ax=ax)
        else:
            df.select_dtypes(include="number").plot.line(ax=ax)
    elif chart_type == "pie":
        col = y_col or df.select_dtypes(include="number").columns[0]
        labels = df[x_col] if x_col else df.index
        ax.pie(df[col], labels=labels, autopct="%1.1f%%")
    elif chart_type == "scatter":
        if x_col and y_col:
            df.plot.scatter(x=x_col, y=y_col, ax=ax)
        else:
            cols = df.select_dtypes(include="number").columns
            df.plot.scatter(x=cols[0], y=cols[1], ax=ax)
    elif chart_type == "hist":
        col = y_col or df.select_dtypes(include="number").columns[0]
        df[col].plot.hist(ax=ax, bins=20)

    ax.set_title(title or f"{target} - {chart_type}")
    plt.tight_layout()
    plt.savefig(output_path, dpi=150)
    plt.close(fig)

    return output_path
