# -*- coding: utf-8 -*-
"""阶段5：将 LoRA 适配器权重合并进基座，得到可独立推理的中医专家模型。

用法：
    python merge_model.py [--adapter output/final] [--save-dir D:\\Model\\TCM-Expert-LoRA]
"""
import argparse

import torch
from transformers import AutoModelForCausalLM, AutoTokenizer
from peft import PeftModel

import config


def main():
    parser = argparse.ArgumentParser(description="合并 LoRA 权重")
    parser.add_argument("--base", default=str(config.BASE_MODEL_DIR), help="基座模型路径")
    parser.add_argument("--adapter", default=str(config.OUTPUT_DIR / "final"), help="LoRA 适配器路径")
    parser.add_argument("--save-dir", default=str(config.LORA_MODEL_DIR), help="合并模型保存路径")
    args = parser.parse_args()

    use_cuda = torch.cuda.is_available()
    dtype = "auto" if use_cuda else torch.float32

    print(f"加载基座：{args.base}")
    tokenizer = AutoTokenizer.from_pretrained(args.base, use_fast=False, trust_remote_code=True)
    model = AutoModelForCausalLM.from_pretrained(
        args.base, device_map="auto" if use_cuda else None,
        torch_dtype=dtype, trust_remote_code=True
    )

    print(f"加载 LoRA 适配器：{args.adapter}")
    lora_model = PeftModel.from_pretrained(model, args.adapter)

    print("合并权重 ...")
    model_new = lora_model.merge_and_unload()

    print(f"保存合并模型：{args.save_dir}")
    model_new.save_pretrained(args.save_dir, safe_serialization=True)
    tokenizer.save_pretrained(args.save_dir)
    print("模型保存成功！")


if __name__ == "__main__":
    main()
