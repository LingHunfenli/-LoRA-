# -*- coding: utf-8 -*-
"""阶段3（加分项）：问题侧保守规则增强——只改写 instruction，output 原样沿用，绝不编造医学答案。

增强手段：
1. 敬语前缀包裹：请问 / 麻烦您 / 想请教您 / 您好；
2. 请求式后缀：可以详细说说吗 / 期待您的解答 / 谢谢；
3. 同义开头模板替换：介绍一下↔讲讲↔说说、请问↔想问问 等；
4. 全程去重（含与 test/val 跨集）、长度 5~80 字过滤；
5. 与 train_subset.jsonl（1200 条）合并，输出可直接用于 CPU 训练的 train_augmented.jsonl。

用法：
    python data_augment.py [--augment-size 800]
"""
import argparse
import json
import random
import re
import statistics

import config

MIN_Q_LEN = 5
MAX_Q_LEN = 80

# 句尾标点
_TAIL_PUNCT = "？?。！!~～"
# 已含敬语/请求语气时不再重复包裹
_PREFIXED = ("请问", "麻烦", "请教", "您好", "想问问", "想问下", "我想")

PREFIX_WRAPPERS = [
    "请问，{q}",
    "麻烦您解答一下，{q}",
    "想请教您一个问题，{q}",
    "您好，{q}",
    "想问问您，{q}",
    "我想了解一下，{q}",
    "老师好，{q}",
]

SUFFIX_WRAPPERS = [
    "{q}，可以详细说说吗？",
    "{q}，能解释一下吗？",
    "{q}，请您指教。",
    "{q}，谢谢！",
    "{q}，期待您的解答。",
    "{q}，麻烦了。",
]

# 同义开头替换（整词匹配，只在句首生效，避免病句）
OPENER_REWRITES = [
    ("介绍一下", "讲讲"),
    ("请介绍", "请讲讲"),
    ("请描述一下", "说说"),
    ("请简述", "说说"),
    ("请阐述", "讲讲"),
    ("请问", "想问问"),
    ("讲一下", "说说"),
    ("讲解一下", "讲讲"),
    ("谈谈", "说说"),
]


def read_jsonl(path):
    data = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                data.append(json.loads(line.strip()))
    return data


def save_jsonl(data, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for d in data:
            f.write(json.dumps(d, ensure_ascii=False) + "\n")


def body(q):
    """去掉句尾标点，便于包裹后重新加标点。"""
    return q.strip().rstrip(_TAIL_PUNCT).strip()


def add_prefix(q, rng):
    if any(q.startswith(p) for p in _PREFIXED):
        return None
    tpl = rng.choice(PREFIX_WRAPPERS)
    return tpl.format(q=body(q))


def add_suffix(q, rng):
    b = body(q)
    if any(w in b[-8:] for w in ("谢谢", "指教", "说说吗", "解释一下吗", "解答")):
        return None
    tpl = rng.choice(SUFFIX_WRAPPERS)
    return tpl.format(q=b)


def rewrite_opener(q, rng):
    b = body(q)
    for old, new in OPENER_REWRITES:
        if b.startswith(old):
            cand = new + b[len(old):]
            tail = "？" if q.rstrip().endswith(("？", "?")) else ("。" if q.rstrip().endswith("。") else "")
            return cand + tail
    return None


def cleanup(q):
    q = q.strip()
    q = q.replace("？？", "？").replace("？吗", "吗").replace("吗？？", "吗？")
    q = re.sub(r"[，,]{2,}", "，", q)
    q = re.sub(r"^[，、\s]+", "", q)
    q = re.sub(r"\s+", "", q)
    return q


def is_valid(q, seen):
    if not (MIN_Q_LEN <= len(q) <= MAX_Q_LEN):
        return False
    if q in seen:
        return False
    if "《" in q and "》" not in q:
        return False
    return True


TRANSFORMS = [add_prefix, add_suffix, rewrite_opener]


def main():
    parser = argparse.ArgumentParser(description="中医问答问题侧数据增强")
    parser.add_argument("--augment-size", type=int, default=800, help="目标新增条数")
    args = parser.parse_args()

    rng = random.Random(config.RANDOM_SEED)
    train_clean = read_jsonl(config.TRAIN_CLEAN)
    subset = read_jsonl(config.TRAIN_SUBSET)
    val = read_jsonl(config.VAL_CLEAN)
    test = read_jsonl(config.TEST_CLEAN)

    # 所有已知问题集合：增强结果不得与任何现有问题重复
    seen = {d["instruction"] for d in train_clean}
    seen.update(d["instruction"] for d in subset)
    seen.update(d["instruction"] for d in val)
    seen.update(d["instruction"] for d in test)

    # 答案索引：改写后的问题沿用原答案
    answer_map = {d["instruction"]: d["output"] for d in train_clean}

    augmented = []
    tries = 0
    max_tries = args.augment_size * 60 + 1000
    while len(augmented) < args.augment_size and tries < max_tries:
        tries += 1
        src = rng.choice(train_clean)
        q = src["instruction"]
        # 每条最多叠加 2 个变换
        n_trans = rng.choice([1, 1, 2])
        cand = q
        funcs = rng.sample(TRANSFORMS, k=min(n_trans, len(TRANSFORMS)))
        for fn in funcs:
            out = fn(cand, rng)
            if out:
                cand = cleanup(out)
        cand = cleanup(cand)
        if cand == q or not is_valid(cand, seen):
            continue
        seen.add(cand)
        augmented.append({
            "instruction": cand,
            "input": config.SYSTEM_PROMPT,
            "output": src["output"],   # 答案原样沿用，不编造
            "_source": q,
        })

    # 混合：1200 条子集 + 新增增强样本
    mix = [{k: v for k, v in d.items() if k != "_source"} for d in subset]
    for d in augmented:
        mix.append({"instruction": d["instruction"], "input": d["input"], "output": d["output"]})
    rng.shuffle(mix)
    save_jsonl(mix, config.TRAIN_AUGMENTED)

    # 统计
    aug_lens = sorted(len(d["instruction"]) for d in augmented)
    mean_len = round(statistics.mean(aug_lens), 1) if aug_lens else 0
    print("=" * 60)
    print(f"新增增强样本：{len(augmented)} 条（尝试 {tries} 次）")
    print(f"增强问题长度：mean {mean_len} / max {aug_lens[-1] if aug_lens else 0} 字符")
    print(f"训练混合集 train_augmented.jsonl：{len(mix)} 条（子集 {len(subset)} + 增强 {len(augmented)}）")
    print("\n增强样例（前 10 条）：")
    for d in augmented[:10]:
        print(f"  原: {d['_source']}")
        print(f"  新: {d['instruction']}\n")


if __name__ == "__main__":
    main()
