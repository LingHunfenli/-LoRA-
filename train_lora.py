# -*- coding: utf-8 -*-
"""阶段4：DeepSeek-R1-Distill-Qwen-1.5B + LoRA 中医问答指令微调（CPU/GPU 自适应）。

用法：
    python train_lora.py --smoke-test                     # 冒烟测试：50 条 1 轮
    python train_lora.py --data subset --epochs 3         # 1200 条子集正式训练
    python train_lora.py --data augmented --epochs 3      # 2000 条（子集+增强）训练
"""
import argparse
import json
import os

# Windows CPU 下规避 MKL/OpenMP 重复运行时导致的原生崩溃（0xC0000005）
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")

import torch
from datasets import Dataset
from transformers import (AutoModelForCausalLM, AutoTokenizer,
                          TrainingArguments, Trainer, DataCollatorForSeq2Seq)
from peft import LoraConfig, TaskType, get_peft_model

import config


def parse_args():
    p = argparse.ArgumentParser(description="LoRA 中医问答微调")
    p.add_argument("--data", choices=["subset", "augmented", "full"], default="subset",
                   help="subset=train_subset.jsonl(1200) / augmented=train_augmented.jsonl(2000) / full=train_clean.jsonl(4919)")
    p.add_argument("--epochs", type=float, default=3.0)
    p.add_argument("--batch-size", type=int, default=2)
    p.add_argument("--grad-accum", type=int, default=4)
    p.add_argument("--lr", type=float, default=1e-4)
    p.add_argument("--max-length", type=int, default=512)
    p.add_argument("--eval-limit", type=int, default=0, help="每轮验证最多使用的条数，0 表示用全部验证集（CPU 提速用）")
    p.add_argument("--smoke-test", action="store_true", help="50 条 1 轮快速跑通管线")
    p.add_argument("--grad-ckpt", action="store_true", help="开启梯度检查点（内存紧张时使用，更慢）")
    p.add_argument("--resume", action="store_true", help="从 output 目录最新 checkpoint 续训")
    p.add_argument("--save-steps", type=int, default=75, help="每多少优化步保存一次 checkpoint")
    p.add_argument("--output", default=str(config.OUTPUT_DIR))
    return p.parse_args()


DATA_FILES = {
    "subset": config.TRAIN_SUBSET,
    "augmented": config.TRAIN_AUGMENTED,
    "full": config.TRAIN_CLEAN,
}


def main():
    args = parse_args()
    use_cuda = torch.cuda.is_available()
    if not use_cuda:
        torch.set_num_threads(int(os.environ.get("TORCH_NUM_THREADS", "8")))
    print(f"运行设备：{'CUDA GPU' if use_cuda else 'CPU'}")

    # ---------- 加载模型与分词器 ----------
    model_dir = str(config.BASE_MODEL_DIR)
    tokenizer = AutoTokenizer.from_pretrained(model_dir, use_fast=False, trust_remote_code=True)
    if tokenizer.pad_token is None:
        tokenizer.pad_token = tokenizer.eos_token

    model = AutoModelForCausalLM.from_pretrained(
        model_dir,
        device_map="auto" if use_cuda else None,  # CPU 下禁用自动分发，避免层被 offload 到磁盘
        torch_dtype="auto" if use_cuda else torch.float32,
        trust_remote_code=True,
    )
    model.enable_input_require_grads()
    if args.grad_ckpt:
        model.gradient_checkpointing_enable()
        print("已开启梯度检查点")

    # ---------- 加载数据 ----------
    train_path = DATA_FILES[args.data]
    print(f"训练数据：{train_path}")
    train_ds = Dataset.from_json(str(train_path))
    val_ds = Dataset.from_json(str(config.VAL_CLEAN))
    if args.smoke_test:
        train_ds = train_ds.select(range(min(50, len(train_ds))))
        val_ds = val_ds.select(range(min(20, len(val_ds))))
        args.epochs = 1
        args.max_length = 256
        args.output = str(config.PROJECT_DIR / "output_smoke")
    elif args.eval_limit > 0:
        val_ds = val_ds.select(range(min(args.eval_limit, len(val_ds))))
    print(f"训练 {len(train_ds)} 条 / 验证 {len(val_ds)} 条")

    max_length = args.max_length

    def process_func(example):
        """构造 chat template 输入；prompt 段 labels=-100，仅 output+eos 参与损失。"""
        messages = [
            {"role": "system", "content": config.SYSTEM_PROMPT},
            {"role": "user", "content": example["instruction"]},
        ]
        text = tokenizer.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True
        )
        prompt_ids = tokenizer(text, add_special_tokens=False)["input_ids"]
        response_ids = tokenizer(example["output"], add_special_tokens=False)["input_ids"] + [tokenizer.eos_token_id]

        input_ids = prompt_ids + response_ids
        labels = [-100] * len(prompt_ids) + response_ids
        attention_mask = [1] * len(input_ids)
        if len(input_ids) > max_length:
            input_ids = input_ids[:max_length]
            attention_mask = attention_mask[:max_length]
            labels = labels[:max_length]
        return {"input_ids": input_ids, "attention_mask": attention_mask, "labels": labels}

    train_dataset = train_ds.map(process_func, remove_columns=["instruction", "input", "output"])
    val_dataset = val_ds.map(process_func, remove_columns=["instruction", "input", "output"])
    train_dataset = train_dataset.shuffle(seed=config.RANDOM_SEED)
    print("预处理示例：", train_dataset[0])

    # ---------- LoRA 配置 ----------
    lora_config = LoraConfig(
        task_type=TaskType.CAUSAL_LM,
        inference_mode=False,
        r=16,
        lora_alpha=32,                      # alpha/r = 2 的缩放
        lora_dropout=0.05,
        target_modules=["q_proj", "k_proj", "v_proj", "o_proj"],
    )
    model = get_peft_model(model, lora_config)
    model.print_trainable_parameters()

    # ---------- 训练参数 ----------
    # transformers 新版用 warmup_steps 替代 warmup_ratio：按总优化步数 3% 估算
    total_optim_steps = max(1, len(train_dataset) // (args.batch_size * args.grad_accum)) * args.epochs
    warmup_steps = max(2, int(total_optim_steps * 0.03))
    print(f"预计总优化步数：{int(total_optim_steps)}，warmup_steps={warmup_steps}")

    training_args = TrainingArguments(
        output_dir=args.output,
        num_train_epochs=args.epochs,
        learning_rate=args.lr,
        per_device_train_batch_size=args.batch_size,
        per_device_eval_batch_size=args.batch_size,
        gradient_accumulation_steps=args.grad_accum,
        weight_decay=0.01,
        lr_scheduler_type="cosine",
        warmup_steps=warmup_steps,
        eval_strategy="epoch",
        save_strategy="steps",   # 周期保存 checkpoint，CPU 长训练中断可续训
        save_steps=args.save_steps,
        save_total_limit=2,
        logging_strategy="steps",
        logging_steps=10,
        use_cpu=not use_cuda,
        report_to="none",
        bf16=use_cuda and torch.cuda.is_bf16_supported(),
        load_best_model_at_end=False,
    )

    trainer = Trainer(
        model=model,
        args=training_args,
        train_dataset=train_dataset,
        eval_dataset=val_dataset,
        data_collator=DataCollatorForSeq2Seq(tokenizer=tokenizer, padding=True),
    )

    # 若 output 目录已有 checkpoint 则自动续训（CPU 长训练中断恢复用）
    trainer.train(resume_from_checkpoint=bool(args.resume))

    # 保存最终 LoRA 适配器（固定 final 子目录，供合并脚本使用）
    from pathlib import Path
    final_dir = Path(args.output) / "final"
    trainer.save_model(str(final_dir))
    tokenizer.save_pretrained(str(final_dir))
    print(f"LoRA 适配器已保存：{final_dir}")

    # 保存训练日志，供报告引用
    config.RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    log_path = config.RESULTS_DIR / "training_log.json"
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(trainer.state.log_history, f, ensure_ascii=False, indent=2)
    print(f"训练日志已保存：{log_path}")


if __name__ == "__main__":
    main()
