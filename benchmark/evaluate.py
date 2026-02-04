"""
评估 ExcelAgent 在 SpreadsheetBench 上的输出。

Usage:
    python benchmark/evaluate.py \
        --dataset sample_data_200 \
        --model <model-name> \
        --setting agent
"""

import os
import sys
import json
import argparse
from tqdm import tqdm
from collections import defaultdict

# 将 SpreadsheetBench 的 evaluation 目录加入 path
EVAL_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "SpreadsheetBench-NoDocker",
    "evaluation",
)
sys.path.insert(0, EVAL_DIR)

from evaluation import compare_workbooks  # noqa: E402

BENCH_DATA_ROOT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "SpreadsheetBench-NoDocker",
    "data",
)


def evaluate(dataset_name: str, setting: str, model: str):
    """比较输出文件与标准答案，计算 soft/hard accuracy。"""
    dataset_path = os.path.join(BENCH_DATA_ROOT, dataset_name)
    if not os.path.exists(dataset_path):
        print(f"Dataset path not found: {dataset_path}")
        return

    with open(os.path.join(dataset_path, "dataset.json"), "r", encoding="utf-8") as f:
        dataset = json.load(f)

    safe_model = model.replace("/", "_")
    output_dir = os.path.join(dataset_path, "outputs", f"{setting}_{safe_model}")

    if not os.path.exists(output_dir):
        print(f"Output directory not found: {output_dir}")
        print("Have you run benchmark/run_benchmark.py first?")
        return

    eval_results = []
    stats = defaultdict(lambda: {"total": 0, "soft_sum": 0.0, "hard_sum": 0})

    for data in tqdm(dataset, desc="Evaluating"):
        task_id = str(data["id"])
        instruction_type = data["instruction_type"]
        answer_position = data["answer_position"]

        test_case_results = []
        for tc_idx in range(1, 4):
            gt_path = os.path.join(
                dataset_path, "spreadsheet", task_id, f"{tc_idx}_{task_id}_answer.xlsx"
            )
            proc_path = os.path.join(output_dir, f"{tc_idx}_{task_id}_output.xlsx")

            try:
                result, _ = compare_workbooks(
                    gt_path, proc_path, instruction_type, answer_position
                )
            except Exception:
                result = False

            test_case_results.append(int(result))

        soft = test_case_results.count(1) / len(test_case_results) if test_case_results else 0
        hard = 0 if 0 in test_case_results else 1

        eval_results.append(
            {
                "id": task_id,
                "instruction_type": instruction_type,
                "test_case_results": test_case_results,
                "soft_restriction": soft,
                "hard_restriction": hard,
            }
        )

        stats[instruction_type]["total"] += 1
        stats[instruction_type]["soft_sum"] += soft
        stats[instruction_type]["hard_sum"] += hard
        stats["overall"]["total"] += 1
        stats["overall"]["soft_sum"] += soft
        stats["overall"]["hard_sum"] += hard

    # 保存详细结果
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)
    result_path = os.path.join(log_dir, f"eval_{setting}_{safe_model}.json")
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump(eval_results, f, indent=2, ensure_ascii=False)

    # 打印汇总
    print()
    print("=" * 60)
    print(f"  Evaluation: {setting}_{safe_model}")
    print(f"  Dataset: {dataset_name}")
    print("=" * 60)
    for category in sorted(stats.keys()):
        s = stats[category]
        n = s["total"]
        soft_avg = s["soft_sum"] / n if n > 0 else 0
        hard_avg = s["hard_sum"] / n if n > 0 else 0
        print(f"  {category:30s}  n={n:4d}  soft={soft_avg:.4f}  hard={hard_avg:.4f}")
    print("=" * 60)
    print(f"  Detailed results: {result_path}")
    print()


def main():
    parser = argparse.ArgumentParser(description="Evaluate ExcelAgent on SpreadsheetBench")
    parser.add_argument("--dataset", type=str, default="sample_data_200")
    parser.add_argument("--model", type=str, required=True, help="Model name")
    parser.add_argument("--setting", type=str, default="agent")
    args = parser.parse_args()

    evaluate(args.dataset, args.setting, args.model)


if __name__ == "__main__":
    main()
