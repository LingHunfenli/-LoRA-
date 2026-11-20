# -*- coding: utf-8 -*-
"""生成完整测试集数据报告（所有 test_clean.jsonl 条目）。

输出：
    results/test_dataset_report.md   ← 完整报告（题目 + 参考答案）
    results/test_dataset_stats.json  ← 统计摘要
"""
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config


def main():
    p = config.TEST_CLEAN
    if not p.exists():
        print(f"[ERROR] 测试集不存在: {p}")
        sys.exit(1)

    items = []
    with open(p, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                d = json.loads(line)
                items.append({
                    "instruction": d.get("instruction", ""),
                    "output": d.get("output", ""),
                    "input": d.get("input", ""),
                })
            except json.JSONDecodeError:
                continue

    # ---------- 统计 ----------
    q_lens = [len(it["instruction"]) for it in items]
    a_lens = [len(it["output"]) for it in items]
    stats = {
        "total_count": len(items),
        "question_len": {
            "min": min(q_lens), "max": max(q_lens),
            "avg": round(sum(q_lens) / len(q_lens), 1),
        },
        "answer_len": {
            "min": min(a_lens), "max": max(a_lens),
            "avg": round(sum(a_lens) / len(a_lens), 1),
        },
    }

    # 答案长度分桶
    buckets = {"0-50字": 0, "51-100字": 0, "101-200字": 0, "201-500字": 0, "500+字": 0}
    for l in a_lens:
        if l <= 50: buckets["0-50字"] += 1
        elif l <= 100: buckets["51-100字"] += 1
        elif l <= 200: buckets["101-200字"] += 1
        elif l <= 500: buckets["201-500字"] += 1
        else: buckets["500+字"] += 1
    stats["answer_len_distribution"] = buckets

    stats_path = config.RESULTS_DIR / "test_dataset_stats.json"
    with open(stats_path, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)
    print(f"[OK] 统计摘要 → {stats_path}")

    # ---------- Markdown 报告 ----------
    md_lines = []
    md_lines.append("# 测试集完整数据报告\n")
    md_lines.append(f"**数据文件**: `test_clean.jsonl`  ")
    md_lines.append(f"**总条数**: {len(items)}  ")
    md_lines.append(f"**生成时间**: {Path(__file__).name}  ")
    md_lines.append("")

    # 统计表
    md_lines.append("## 一、统计摘要\n")
    md_lines.append("| 指标 | 最小值 | 最大值 | 平均值 |")
    md_lines.append("|---|---|---|---|")
    md_lines.append(f"| 问题长度（字符） | {stats['question_len']['min']} | {stats['question_len']['max']} | {stats['question_len']['avg']} |")
    md_lines.append(f"| 答案长度（字符） | {stats['answer_len']['min']} | {stats['answer_len']['max']} | {stats['answer_len']['avg']} |")
    md_lines.append("")

    md_lines.append("### 答案长度分布\n")
    md_lines.append("| 区间 | 条数 | 占比 |")
    md_lines.append("|---|---|---|")
    for k, v in buckets.items():
        pct = f"{v / len(items) * 100:.1f}%"
        md_lines.append(f"| {k} | {v} | {pct} |")
    md_lines.append("")

    # 完整数据
    md_lines.append("## 二、完整测试数据\n")
    md_lines.append("> 以下为 test_clean.jsonl 全部 {} 条数据，每条包含问题和参考答案。\n".format(len(items)))

    for i, it in enumerate(items, 1):
        md_lines.append(f"### {i}. {it['instruction']}\n")
        if it["input"]:
            md_lines.append(f"**输入补充**: {it['input']}\n")
        md_lines.append(f"**参考答案**:\n\n> {it['output']}\n")

    report_path = config.RESULTS_DIR / "test_dataset_report.md"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    print(f"[OK] 完整报告 → {report_path}")
    print(f"     大小: {report_path.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    main()
