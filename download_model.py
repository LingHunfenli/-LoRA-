# -*- coding: utf-8 -*-
"""阶段1：从 ModelScope 下载 DeepSeek-R1-Distill-Qwen-1.5B 基座到 D:\\Model。

用法：
    python download_model.py
"""
import sys
from modelscope import snapshot_download, AutoTokenizer

import config


def main():
    config.MODEL_ROOT.mkdir(parents=True, exist_ok=True)
    if not config.MODEL_ROOT.exists():
        print(f"[错误] 模型根目录不可写或不存在：{config.MODEL_ROOT}")
        sys.exit(1)

    target = str(config.BASE_MODEL_DIR)
    print(f"开始下载基座模型 {config.BASE_MODEL_ID} -> {target}")
    model_dir = snapshot_download(
        config.BASE_MODEL_ID,
        local_dir=target,
    )
    print(f"下载完成：{model_dir}")

    # 空加载自检：仅加载分词器（速度快、占用小）
    print("分词器加载自检 ...")
    tokenizer = AutoTokenizer.from_pretrained(model_dir, trust_remote_code=True)
    print(f"自检通过，词表大小：{len(tokenizer)}，pad_token：{tokenizer.pad_token}，eos_token：{tokenizer.eos_token}")


if __name__ == "__main__":
    main()
