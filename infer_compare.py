# -*- coding: utf-8 -*-
"""阶段6：微调前后对比验证——在测试集 5 类未见问题上比较基座与微调模型的回答。

输出：
    results/compare_cases.jsonl  逐题原始结果
    results/compare_cases.md     四列对比表（问题/微调前/微调后/参考答案）

用法：
    python infer_compare.py
"""
import gc
import json

import torch
from transformers import AutoModelForCausalLM, AutoTokenizer

import config

MAX_NEW_TOKENS = 512

# 五类问题的关键词匹配规则（在 test_clean 中检索，且确保未出现在训练数据中）
CATEGORY_RULES = [
    ("中药功效", lambda q: ("功效" in q or "作用" in q) and len(q) <= 25),
    ("方剂推荐", lambda q: "方剂" in q and ("推荐" in q or "有哪些" in q)),
    ("辨证论治", lambda q: "辨证" in q or ("证" in q and "为什么" in q)),
    ("养生药膳", lambda q: "药膳" in q or "食疗" in q),
    ("病症解释", lambda q: ("是什么" in q or "解释" in q) and "？" in q),
]


def read_jsonl(path):
    data = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                data.append(json.loads(line.strip()))
    return data


def select_cases(test_data, trained_questions):
    """从测试集中按类别选取未在训练中出现过的问题。"""
    cases = []
    used = set()
    for category, rule in CATEGORY_RULES:
        picked = None
        for d in test_data:
            q = d["instruction"]
            if q in used or q in trained_questions:
                continue
            if rule(q):
                picked = d
                break
        if picked is None:
            raise RuntimeError(f"类别【{category}】未找到符合条件的测试问题，请调整匹配规则")
        used.add(picked["instruction"])
        cases.append({"category": category, **picked})
    return cases


def parse_answer(text):
    """兼容 R1 思维链：若输出含字面 ，取最后一段；否则整体作为答案。"""
    text = text.strip()
    if "</think>" in text:
        thinking, answer = text.rsplit("</think>", 1)
        return answer.strip("\n").strip(), thinking.strip("\n").strip()
    return text, ""


@torch.inference_mode()
def predict(model, tokenizer, question):
    messages = [
        {"role": "system", "content": config.SYSTEM_PROMPT},
        {"role": "user", "content": question},
    ]
    text = tokenizer.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
    inputs = tokenizer([text], return_tensors="pt").to(model.device)
    generated = model.generate(
        **inputs,
        max_new_tokens=MAX_NEW_TOKENS,
        do_sample=False,
        pad_token_id=tokenizer.pad_token_id,
    )
    out_ids = generated[0][inputs["input_ids"].shape[1]:].tolist()
    raw = tokenizer.decode(out_ids, skip_special_tokens=True)
    answer, thinking = parse_answer(raw)
    return answer, thinking, raw


def load_model(path):
    use_cuda = torch.cuda.is_available()
    model = AutoModelForCausalLM.from_pretrained(
        path, torch_dtype="auto" if use_cuda else torch.float32,
        device_map="auto" if use_cuda else None, trust_remote_code=True,
    )
    model.eval()
    tokenizer = AutoTokenizer.from_pretrained(path, trust_remote_code=True)
    if tokenizer.pad_token is None:
        tokenizer.pad_token = tokenizer.eos_token
    return model, tokenizer


def free_model(model):
    del model
    gc.collect()
    if torch.cuda.is_available():
        torch.cuda.empty_cache()


def main():
    config.RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    test_data = read_jsonl(config.TEST_CLEAN)
    trained = set()
    for p in (config.TRAIN_SUBSET, config.TRAIN_AUGMENTED, config.VAL_CLEAN):
        trained.update(d["instruction"] for d in read_jsonl(p))

    cases = select_cases(test_data, trained)
    print("选定的 5 个验证问题：")
    for c in cases:
        print(f"  [{c['category']}] {c['instruction']}")

    results = []

    print("\n>>> 加载基座模型（微调前）...")
    base_model, base_tok = load_model(str(config.BASE_MODEL_DIR))
    for c in cases:
        print(f"  基座推理：[{c['category']}] {c['instruction'][:20]}...")
        answer, thinking, raw = predict(base_model, base_tok, c["instruction"])
        c["base_answer"] = answer
        c["base_thinking"] = thinking
    free_model(base_model)
    del base_tok
    gc.collect()

    print("\n>>> 加载微调模型...")
    lora_model, lora_tok = load_model(str(config.LORA_MODEL_DIR))
    for c in cases:
        print(f"  微调推理：[{c['category']}] {c['instruction'][:20]}...")
        answer, thinking, raw = predict(lora_model, lora_tok, c["instruction"])
        c["lora_answer"] = answer
        c["lora_thinking"] = thinking
    free_model(lora_model)
    del lora_tok
    gc.collect()

    # 落盘 jsonl
    with open(config.RESULTS_DIR / "compare_cases.jsonl", "w", encoding="utf-8") as f:
        for c in cases:
            f.write(json.dumps(c, ensure_ascii=False) + "\n")

    # 客观指标
    base_lens = [len(c["base_answer"]) for c in cases]
    lora_lens = [len(c["lora_answer"]) for c in cases]
    base_think_rate = sum(1 for c in cases if c["base_thinking"]) / len(cases)
    lora_think_rate = sum(1 for c in cases if c["lora_thinking"]) / len(cases)

    # 落盘 Markdown 对比表
    lines = [
        "# 微调前后中医问答对比（5 题）",
        "",
        f"- 基座模型：{config.BASE_MODEL_ID}（{config.BASE_MODEL_DIR}）",
        f"- 微调模型：{config.LORA_MODEL_DIR}",
        f"- 解码：greedy，max_new_tokens={MAX_NEW_TOKENS}",
        f"- 客观指标：微调前平均回答 {sum(base_lens)/len(base_lens):.0f} 字 / 思维链出现率 {base_think_rate:.0%}；"
        f"微调后平均回答 {sum(lora_lens)/len(lora_lens):.0f} 字 / 思维链出现率 {lora_think_rate:.0%}",
        "",
    ]
    for c in cases:
        lines += [
            f"## 【{c['category']}】{c['instruction']}",
            "",
            f"**微调前（基座）：**",
            "",
            c["base_answer"].replace("\n", "  \n"),
            "",
            f"**微调后（LoRA）：**",
            "",
            c["lora_answer"].replace("\n", "  \n"),
            "",
            f"**数据集参考答案：**",
            "",
            c["output"].replace("\n", "  \n"),
            "",
            "---",
            "",
        ]
    with open(config.RESULTS_DIR / "compare_cases.md", "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print("\n对比完成：")
    print(f"  {config.RESULTS_DIR / 'compare_cases.jsonl'}")
    print(f"  {config.RESULTS_DIR / 'compare_cases.md'}")


if __name__ == "__main__":
    main()
