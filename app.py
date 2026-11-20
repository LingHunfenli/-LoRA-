# -*- coding: utf-8 -*-
"""阶段7（加分项）：Gradio 中医专家问答界面

功能：
  1. 双栏并排展示微调前（基座）/ 微调后（LoRA）回答；
  2. 输入问题精确命中测试集时，额外展示数据集参考答案。

GPU/CPU 混合策略：
  - 有 CUDA 且显存 ≥ 2GB → 基座 fp16 放 GPU，tokenizer CPU
  - 否则 → 纯 CPU fp32
  - LoRA 适配器跟随基座设备（peft 默认行为）
  - 基座和 LoRA 共用同一份权重，通过 disable_adapter() 切换对比

用法：
    python app.py
"""
import json
import time

import torch
import gradio as gr
from transformers import AutoModelForCausalLM, AutoTokenizer
from peft import PeftModel

import config

MAX_NEW_TOKENS_BASE = 1024  # 基座有 <think> 思维链，需要更多 token
MAX_NEW_TOKENS_LORA = 512   # LoRA 微调后倾向短答
REPETITION_PENALTY = 1.15   # 防止重复生成循环
TEMPERATURE = 0.7           # 采样温度，确定性+多样性平衡

# 显存阈值：低于此值放弃 GPU（1.5B fp16 约 3GB，留余量）
MIN_GPU_MEM_GB = 2.0


def _detect_device():
    """智能选择推理设备，返回 (dtype, device_str, info_text)。"""
    if not torch.cuda.is_available():
        return torch.float32, "cpu", "CPU 推理（未检测到 CUDA）"
    gpu_mem = torch.cuda.get_device_properties(0).total_memory / 1024**3
    gpu_name = torch.cuda.get_device_name(0)
    if gpu_mem < MIN_GPU_MEM_GB:
        return torch.float32, "cpu", f"CPU 推理（GPU 显存仅 {gpu_mem:.1f}GB，不足 {MIN_GPU_MEM_GB}GB）"
    cap = torch.cuda.get_device_capability(0)
    use_bf16 = cap[0] >= 8  # Ada/Lovelace/Hopper 架构支持 bf16
    dtype = torch.bfloat16 if use_bf16 else torch.float16
    dtype_name = "bf16" if use_bf16 else "fp16"
    return dtype, "cuda:0", f"GPU 推理（{gpu_name}, {gpu_mem:.1f}GB, {dtype_name}）"


def load_model():
    dtype, device, info = _detect_device()
    print(f"[DEVICE] {info}")
    tokenizer = AutoTokenizer.from_pretrained(str(config.BASE_MODEL_DIR), trust_remote_code=True)
    if tokenizer.pad_token is None:
        tokenizer.pad_token = tokenizer.eos_token

    print("[LOAD] 加载基座模型...")
    t0 = time.time()
    base = AutoModelForCausalLM.from_pretrained(
        str(config.BASE_MODEL_DIR),
        torch_dtype=dtype,
        device_map={"": device},  # 强制整个模型放同一个设备
        trust_remote_code=True,
    )
    model = PeftModel.from_pretrained(base, str(config.OUTPUT_DIR / "final"))
    model.eval()
    print(f"[LOAD] 完成，耗时 {time.time() - t0:.1f}s，设备: {next(model.parameters()).device}")
    return model, tokenizer, info


MODEL, TOKENIZER, DEVICE_INFO = None, None, ""


def _strip_think(text):
    """去掉 <think>...</think> 思维链，仅保留正式回答。

    R1-Distill 系模型原生倾向先输出 <think>...</think> 再给答案；
    LoRA 微调后通常直接作答。若生成被截断、标签未闭合，则保留原文。
    """
    text = (text or "").strip()
    if "<think>" in text:
        rest = text.split("<think>", 1)[1]
        if "</think>" in rest:
            return rest.split("</think>", 1)[1].strip()
        return text
    return text


@torch.inference_mode()
def _generate(model, tokenizer, question, max_new_tokens):
    """生成并返回模型原始输出文本。"""
    messages = [
        {"role": "system", "content": config.SYSTEM_PROMPT},
        {"role": "user", "content": question},
    ]
    prompt_text = tokenizer.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
    inputs = tokenizer([prompt_text], return_tensors="pt").to(model.device)
    eos_id = tokenizer.eos_token_id
    out = model.generate(
        **inputs,
        max_new_tokens=max_new_tokens,
        do_sample=True,
        temperature=TEMPERATURE,
        repetition_penalty=REPETITION_PENALTY,
        eos_token_id=eos_id,
        pad_token_id=tokenizer.pad_token_id,
    )
    return tokenizer.decode(out[0][inputs["input_ids"].shape[1]:], skip_special_tokens=True)


# ---------------- 测试集精确命中 ----------------

def _load_test_refs():
    refs = {}
    p = config.TEST_CLEAN
    if p.exists():
        with open(p, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    d = json.loads(line)
                    refs[d["instruction"]] = d["output"]
    return refs


_TEST_REF = _load_test_refs()


def _examples():
    p = config.RESULTS_DIR / "compare_cases.jsonl"
    if p.exists():
        with open(p, "r", encoding="utf-8") as f:
            return [[json.loads(l)["instruction"]] for l in f if l.strip()]
    return [
        ["黄芪有什么功效？适合什么人群？"],
        ["脾胃虚寒可以用什么方剂调理？"],
        ["风寒感冒和风热感冒在辨证上有什么区别？"],
    ]


# ---------------- 主处理函数 ----------------

def answer_both(question):
    global MODEL, TOKENIZER, DEVICE_INFO
    if MODEL is None:
        MODEL, TOKENIZER, DEVICE_INFO = load_model()
    question = (question or "").strip()
    if not question:
        return ("请输入中医相关问题。", "", "")
    # 微调前：禁用适配器，基座需要更多 token（含 <think> 思维链）
    with MODEL.disable_adapter():
        before = _strip_think(_generate(MODEL, TOKENIZER, question, MAX_NEW_TOKENS_BASE))
    # 微调后：启用 LoRA，短答够用
    after = _strip_think(_generate(MODEL, TOKENIZER, question, MAX_NEW_TOKENS_LORA))
    # 精确命中测试集时展示参考答案
    reference = _TEST_REF.get(question, "")
    return before, after, reference


def build_ui():
    with gr.Blocks(title="中医医学专家大模型 · LoRA 微调前后对比") as demo:
        gr.Markdown("# 中医医学专家大模型 · LoRA 微调前后对比")
        gr.Markdown(f"系统提示语：{config.SYSTEM_PROMPT}")
        with gr.Row():
            question = gr.Textbox(label="请输入中医问题", lines=2, scale=4)
            btn = gr.Button("提问", variant="primary", scale=1)
        with gr.Row():
            with gr.Column():
                gr.Markdown("#### 微调前（DeepSeek-R1-Distill-Qwen-1.5B 基座）")
                before = gr.Textbox(label="回答", lines=12, max_lines=50, show_copy_button=True)
            with gr.Column():
                gr.Markdown("#### 微调后（LoRA 中医专家模型）")
                after = gr.Textbox(label="回答", lines=12, max_lines=50, show_copy_button=True)
        reference = gr.Textbox(label="数据集参考答案（精确命中测试集时显示）", lines=4, max_lines=20, show_copy_button=True)
        gr.Examples(examples=_examples(), inputs=question, label="示例问题")
        btn.click(answer_both, inputs=question, outputs=[before, after, reference])
        question.submit(answer_both, inputs=question, outputs=[before, after, reference])
    return demo


if __name__ == "__main__":
    ui = build_ui()
    ui.launch(server_name="0.0.0.0", server_port=7860, inbrowser=True, pwa=True)
