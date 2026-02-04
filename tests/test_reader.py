"""测试 Excel 读取工具"""

import os
import pytest

from tools.reader import read_excel, read_excel_formulas, get_summary
from tools.writer import create_workbook, write_cells

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
SAMPLE = os.path.join(FIXTURES, "sample.xlsx")


@pytest.fixture(autouse=True)
def setup_sample():
    """创建测试用 Excel 文件"""
    os.makedirs(FIXTURES, exist_ok=True)
    create_workbook(SAMPLE, {
        "销售数据": [
            ["产品", "Q1", "Q2", "Q3", "Q4"],
            ["产品A", 100, 200, 150, 300],
            ["产品B", 80, 120, 90, 180],
            ["产品C", 200, 250, 300, 400],
        ],
        "汇总": [
            ["指标", "值"],
            ["总计", "=SUM(销售数据!B2:E4)"],
        ],
    })
    yield
    if os.path.exists(SAMPLE):
        os.remove(SAMPLE)


def test_read_excel():
    info = read_excel(SAMPLE)
    assert "销售数据" in info["sheets"]
    assert "汇总" in info["sheets"]
    assert info["shape"]["销售数据"] == (3, 5)


def test_read_specific_sheet():
    info = read_excel(SAMPLE, sheet_name="销售数据")
    assert len(info["data"]) == 1
    assert "销售数据" in info["data"]


def test_read_formulas():
    formulas = read_excel_formulas(SAMPLE, sheet_name="汇总")
    assert len(formulas["汇总"]) == 1
    assert formulas["汇总"][0]["formula"].startswith("=SUM")


def test_get_summary():
    summary = get_summary(SAMPLE)
    assert "销售数据" in summary
    assert "产品" in summary


def test_file_not_found():
    with pytest.raises(FileNotFoundError):
        read_excel("/nonexistent/file.xlsx")
