"""Python 代码执行工具 — 用于处理预定义工具无法覆盖的复杂操作"""

import os
import subprocess
import sys
import tempfile


def run_python(code: str, timeout: int = 60) -> str:
    """执行 Python 代码，返回 stdout + stderr。

    代码可使用 openpyxl、pandas、numpy 等已安装的库。
    适用于预定义工具无法处理的复杂电子表格操作。

    Args:
        code: 要执行的 Python 代码
        timeout: 最大执行时间（秒）

    Returns:
        执行输出或错误信息
    """
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".py", delete=False, encoding="utf-8"
    ) as f:
        f.write(code)
        temp_path = f.name

    try:
        result = subprocess.run(
            [sys.executable, temp_path],
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=os.getcwd(),
        )
        output = ""
        if result.stdout:
            output += result.stdout
        if result.stderr:
            if output:
                output += "\n"
            output += result.stderr
        if result.returncode != 0 and not result.stderr:
            output += f"\n[exit code: {result.returncode}]"
        return output.strip() or "[Code executed successfully with no output]"
    except subprocess.TimeoutExpired:
        return f"[Execution timed out after {timeout} seconds]"
    except Exception as e:
        return f"[Execution error: {e}]"
    finally:
        try:
            os.unlink(temp_path)
        except OSError:
            pass
