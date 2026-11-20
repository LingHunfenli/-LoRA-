# -*- coding: utf-8 -*-
"""阶段2：中医问答数据清洗、字段规范化、切分与 CPU 训练子集采样。

处理规则：
1. 剔除标题式噪声（output 为《……》且长度 < 30 字）；
2. 按 instruction 去重（保留首条）；跨集去重（test 中的问题不出现在 train/val）；
3. 长度过滤：instruction >= 5 字、output 10~400 字；字段类型校验；
4. 字段规范化：统一 {instruction, input, output} 三字段，input 改写为图片版标准系统提示语；
5. 清洗后训练集 95:5 划出验证集（seed=42）；
6. 采样 CPU 训练子集（默认 1200 条）。

用法：
    python data_clean.py [--subset-size 1200] [--val-ratio 0.05]
"""
import argparse
import json
import random
import statistics

import config

TITLE_MAX_LEN = 30          # 《……》且短于该长度判为标题噪声
INSTRUCTION_MIN_LEN = 5
OUTPUT_MIN_LEN = 10
OUTPUT_MAX_LEN = 400


def load_jsonl(path):
    data = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                data.append(json.loads(line))
    return data


def save_jsonl(data, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for d in data:
            f.write(json.dumps(d, ensure_ascii=False) + "\n")


def clean_split(raw, name, stats):
    """清洗单个数据分片，返回规范化后的记录列表。"""
    result = []
    seen = set()
    for d in raw:
        # 字段校验
        if not all(isinstance(d.get(k), str) for k in ("instruction", "input", "output")):
            stats[name]["bad_field"] += 1
            continue
        instruction = d["instruction"].strip()
        output = d["output"].strip()

        # 标题式噪声
        if (output.startswith("《") and output.endswith("》")
                and len(output) < TITLE_MAX_LEN):
            stats[name]["title"] += 1
            continue
        # 长度过滤
        if len(instruction) < INSTRUCTION_MIN_LEN:
            stats[name]["short_instruction"] += 1
            continue
        if not (OUTPUT_MIN_LEN <= len(output) <= OUTPUT_MAX_LEN):
            stats[name]["bad_output_len"] += 1
            continue
        # 去重
        if instruction in seen:
            stats[name]["dup"] += 1
            continue
        seen.add(instruction)

        # 字段规范化（严格按任务图片格式）
        result.append({
            "instruction": instruction,
            "input": config.SYSTEM_PROMPT,
            "output": output,
        })
    return result


def length_summary(data, key):
    lens = sorted(len(d[key]) for d in data)
    if not lens:
        return {"count": 0}
    n = len(lens)
    return {
        "count": n,
        "mean": round(statistics.mean(lens), 1),
        "p50": lens[n // 2],
        "p90": lens[int(n * 0.9)],
        "max": lens[-1],
    }


def main():
    parser = argparse.ArgumentParser(description="中医问答数据清洗与切分")
    parser.add_argument("--subset-size", type=int, default=1200, help="CPU 训练子集条数")
    parser.add_argument("--val-ratio", type=float, default=0.05, help="验证集占清洗后训练集比例")
    args = parser.parse_args()

    rng = random.Random(config.RANDOM_SEED)
    config.PROCESSED_DATA_DIR.mkdir(parents=True, exist_ok=True)

    report_lines = []

    def log(msg=""):
        print(msg)
        report_lines.append(msg)

    log("=" * 60)
    log("中医问答数据清洗报告")
    log("=" * 60)

    raw_train = load_jsonl(config.TRAIN_RAW)
    raw_test = load_jsonl(config.TEST_RAW)
    log(f"原始数据：train {len(raw_train)} 条 / test {len(raw_test)} 条")

    stats = {
        "train": {"title": 0, "dup": 0, "bad_field": 0, "short_instruction": 0, "bad_output_len": 0},
        "test":  {"title": 0, "dup": 0, "bad_field": 0, "short_instruction": 0, "bad_output_len": 0},
    }
    train = clean_split(raw_train, "train", stats)
    test = clean_split(raw_test, "test", stats)

    for name in ("train", "test"):
        s = stats[name]
        log(f"\n[{name}] 清洗剔除明细：标题噪声 {s['title']}、重复 {s['dup']}、"
            f"字段异常 {s['bad_field']}、问题过短 {s['short_instruction']}、回答长度越界 {s['bad_output_len']}")
    log(f"清洗后（跨集去重前）：train {len(train)} 条 / test {len(test)} 条")

    # 跨集去重：test 中出现过的问题，从 train 中剔除，保证评测纯净
    test_questions = {d["instruction"] for d in test}
    before = len(train)
    train = [d for d in train if d["instruction"] not in test_questions]
    log(f"跨集去重：从 train 剔除 {before - len(train)} 条与 test 重复的问题")

    # 训练集 95:5 切出验证集
    rng.shuffle(train)
    val_size = max(1, int(len(train) * args.val_ratio))
    val = train[:val_size]
    train_keep = train[val_size:]
    log(f"\n切分结果：train {len(train_keep)} / val {len(val)} / test {len(test)}")

    # CPU 训练子集采样
    subset_size = min(args.subset_size, len(train_keep))
    subset = rng.sample(train_keep, subset_size)
    rng.shuffle(subset)
    log(f"CPU 训练子集：{len(subset)} 条 -> {config.TRAIN_SUBSET.name}")

    # 落盘
    save_jsonl(train_keep, config.TRAIN_CLEAN)
    save_jsonl(val, config.VAL_CLEAN)
    save_jsonl(test, config.TEST_CLEAN)
    save_jsonl(subset, config.TRAIN_SUBSET)

    log("\n长度分布（字符数）：")
    for nm, ds in [("train_clean", train_keep), ("val_clean", val), ("test_clean", test)]:
        log(f"  [{nm}] instruction {length_summary(ds, 'instruction')}")
        log(f"  [{nm}] output      {length_summary(ds, 'output')}")

    # 格式断言：三字段齐全且 input 全部等于标准系统提示语
    for ds in (train_keep, val, test, subset):
        assert all(set(d.keys()) == {"instruction", "input", "output"} for d in ds)
        assert all(d["input"] == config.SYSTEM_PROMPT for d in ds)
    log("\n格式断言通过：全部记录均为 {instruction, input, output} 三字段，且 input == 标准系统提示语。")

    with open(config.DATA_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    print(f"\n报告已保存：{config.DATA_REPORT}")


if __name__ == "__main__":
    main()
