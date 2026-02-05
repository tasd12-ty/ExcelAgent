"""
在 SpreadsheetBench 上运行 ExcelAgent。

Usage:
    python benchmark/run_benchmark.py \
        --dataset sample_data_200 \
        --model <model-name> \
        --max_steps 15 \
        [--resume] \
        [--task_ids id1,id2]
"""

import os
import sys
import json
import shutil
import argparse
import time
from datetime import datetime
from tqdm import tqdm

# 将项目根目录加入 sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv

load_dotenv()

from agent.core import run_benchmark, parse_tool_call

# ---------------------------------------------------------------------------
# 路径工具
# ---------------------------------------------------------------------------

BENCH_DATA_ROOT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "SpreadsheetBench-NoDocker",
    "data",
)


def find_dataset_path(dataset_name: str) -> str:
    """定位 dataset.json 所在目录。"""
    candidates = [
        os.path.join(BENCH_DATA_ROOT, dataset_name),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", dataset_name),
    ]
    for p in candidates:
        if os.path.exists(os.path.join(p, "dataset.json")):
            return p
    raise FileNotFoundError(
        f"Dataset '{dataset_name}' not found. Tried: {candidates}\n"
        "Have you extracted the data archive? "
        "e.g.: cd SpreadsheetBench-NoDocker/data && tar -xzf sample_data_200.tar.gz"
    )


def get_file_paths(task_id: str, dataset_path: str, test_case_idx: int, output_dir: str):
    """返回 (input_path, output_path)。"""
    base = f"{test_case_idx}_{task_id}"
    input_path = os.path.join(dataset_path, "spreadsheet", task_id, f"{base}_input.xlsx")
    output_path = os.path.join(output_dir, f"{base}_output.xlsx")
    return input_path, output_path


# ---------------------------------------------------------------------------
# 断点续跑
# ---------------------------------------------------------------------------


def load_completed_ids(log_path: str) -> set:
    """从 JSONL 日志中读取已完成的任务 ID。"""
    completed = set()
    if os.path.exists(log_path):
        with open(log_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    record = json.loads(line)
                    if record.get("status") == "success":
                        completed.add(str(record["id"]))
                except json.JSONDecodeError:
                    continue
    return completed


# ---------------------------------------------------------------------------
# 单任务执行
# ---------------------------------------------------------------------------


def run_single_task(
    data: dict,
    dataset_path: str,
    output_dir: str,
    model: str | None,
    max_steps: int,
) -> dict:
    """对一个 benchmark 任务运行 ExcelAgent（3 个 test case）。"""
    task_id = str(data["id"])
    instruction = data["instruction"]
    instruction_type = data["instruction_type"]
    answer_position = data["answer_position"]

    result = {
        "id": task_id,
        "instruction_type": instruction_type,
        "status": "pending",
        "test_cases": {},
        "timestamp": datetime.now().isoformat(),
    }

    # ---- Test case 1: 完整 agent 运行 ----
    input_path, output_path = get_file_paths(task_id, dataset_path, 1, output_dir)
    if not os.path.exists(input_path):
        result["status"] = "error"
        result["error"] = f"Input not found: {input_path}"
        return result

    shutil.copy2(input_path, output_path)

    t0 = time.time()
    try:
        final_response, messages = run_benchmark(
            instruction=instruction,
            file_path=output_path,
            instruction_type=instruction_type,
            answer_position=answer_position,
            max_steps=max_steps,
            model=model,
        )
        elapsed = round(time.time() - t0, 2)

        # 提取工具调用链
        tool_calls = []
        for msg in messages:
            if msg["role"] == "assistant":
                tc = parse_tool_call(msg["content"])
                if tc:
                    tool_calls.append(tc)

        result["test_cases"]["1"] = {
            "output_exists": os.path.exists(output_path),
            "elapsed_seconds": elapsed,
        }
        result["conversation"] = messages
        result["tool_calls"] = tool_calls
        result["num_steps"] = len(tool_calls)

    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)
        return result

    # ---- Test cases 2 & 3: 重新运行 agent ----
    for tc_idx in [2, 3]:
        inp, out = get_file_paths(task_id, dataset_path, tc_idx, output_dir)
        if not os.path.exists(inp):
            result["test_cases"][str(tc_idx)] = {"error": "input not found"}
            continue

        shutil.copy2(inp, out)
        try:
            run_benchmark(
                instruction=instruction,
                file_path=out,
                instruction_type=instruction_type,
                answer_position=answer_position,
                max_steps=max_steps,
                model=model,
            )
            result["test_cases"][str(tc_idx)] = {
                "output_exists": os.path.exists(out),
            }
        except Exception as e:
            result["test_cases"][str(tc_idx)] = {"error": str(e)}

    result["status"] = "success"
    return result


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description="Run ExcelAgent on SpreadsheetBench")
    parser.add_argument("--dataset", type=str, default="sample_data_200")
    parser.add_argument("--model", type=str, default=None, help="Model ID (uses env default if not set)")
    parser.add_argument("--max_steps", type=int, default=15)
    parser.add_argument("--resume", action="store_true", help="Skip already completed tasks")
    parser.add_argument("--task_ids", type=str, default=None, help="Comma-separated task IDs to run")
    parser.add_argument("--setting", type=str, default="agent", help="Setting name for output dir")
    args = parser.parse_args()

    dataset_path = find_dataset_path(args.dataset)
    with open(os.path.join(dataset_path, "dataset.json"), "r", encoding="utf-8") as f:
        dataset = json.load(f)

    safe_model = (args.model or os.getenv("OPENROUTER_MODEL", "unknown")).replace("/", "_")
    setting_name = f"{args.setting}_{safe_model}"

    output_dir = os.path.join(dataset_path, "outputs", setting_name)
    os.makedirs(output_dir, exist_ok=True)

    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)
    conv_log_path = os.path.join(log_dir, f"conv_{setting_name}.jsonl")

    # 断点续跑
    completed_ids: set = set()
    if args.resume:
        completed_ids = load_completed_ids(conv_log_path)
        print(f"Resume: {len(completed_ids)} tasks already completed, skipping.")

    # 过滤任务
    tasks = dataset
    if args.task_ids:
        target_ids = {s.strip() for s in args.task_ids.split(",")}
        tasks = [d for d in dataset if str(d["id"]) in target_ids]
        print(f"Running {len(tasks)} selected tasks: {args.task_ids}")

    print(f"Dataset: {args.dataset} ({len(tasks)} tasks)")
    print(f"Model: {args.model or os.getenv('OPENROUTER_MODEL', 'default')}")
    print(f"Output dir: {output_dir}")
    print(f"Log file: {conv_log_path}")
    print()

    success_count = 0
    error_count = 0

    for data in tqdm(tasks, desc="ExcelAgent Benchmark"):
        if str(data["id"]) in completed_ids:
            continue

        try:
            result = run_single_task(data, dataset_path, output_dir, args.model, args.max_steps)
        except Exception as e:
            result = {
                "id": data["id"],
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat(),
            }

        if result["status"] == "success":
            success_count += 1
        else:
            error_count += 1

        with open(conv_log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(result, ensure_ascii=False) + "\n")

    print(f"\nDone. success={success_count}, error={error_count}")
    print(f"Results: {conv_log_path}")
    print(f"Outputs: {output_dir}")
    print(f"\nNext step: python benchmark/evaluate.py --dataset {args.dataset} --model {safe_model} --setting {args.setting}")


if __name__ == "__main__":
    main()
