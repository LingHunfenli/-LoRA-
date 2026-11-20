# -*- coding: utf-8 -*-
"""项目全局配置：路径与固定系统提示语（所有脚本统一引用）。

模型路径可通过环境变量 MODEL_ROOT 覆盖（默认 D:\\Model），
方便不同机器按各自习惯放置模型权重。
"""
import os
from pathlib import Path

# ---------- 路径配置 ----------
# 模型根目录：优先读环境变量 MODEL_ROOT，否则默认 D:\Model
# 例：set MODEL_ROOT=E:\AI_Models  或  $env:MODEL_ROOT="E:\AI_Models"
_model_root_env = os.environ.get("MODEL_ROOT", r"D:\Model")
MODEL_ROOT = Path(_model_root_env)
BASE_MODEL_DIR = MODEL_ROOT / "DeepSeek-R1-Distill-Qwen-1.5B"   # 基座模型
LORA_MODEL_DIR = MODEL_ROOT / "TCM-Expert-LoRA"                 # 合并后的微调模型

# 项目目录（src 的上一级）
PROJECT_DIR = Path(__file__).resolve().parent.parent
RAW_DATA_DIR = PROJECT_DIR / "data"                             # 原始数据（只读）
PROCESSED_DATA_DIR = PROJECT_DIR / "data_processed"             # 清洗/增强后数据
OUTPUT_DIR = PROJECT_DIR / "output"                             # checkpoint 输出
RESULTS_DIR = PROJECT_DIR / "results"                           # 对比结果

TRAIN_RAW = RAW_DATA_DIR / "chinese_medical_train.jsonl"
TEST_RAW = RAW_DATA_DIR / "chinese_medical_test.jsonl"

TRAIN_CLEAN = PROCESSED_DATA_DIR / "train_clean.jsonl"
VAL_CLEAN = PROCESSED_DATA_DIR / "val_clean.jsonl"
TEST_CLEAN = PROCESSED_DATA_DIR / "test_clean.jsonl"
TRAIN_AUGMENTED = PROCESSED_DATA_DIR / "train_augmented.jsonl"
TRAIN_SUBSET = PROCESSED_DATA_DIR / "train_subset.jsonl"
DATA_REPORT = PROCESSED_DATA_DIR / "data_report.txt"

# ModelScope 上的基座模型 ID
BASE_MODEL_ID = "deepseek-ai/DeepSeek-R1-Distill-Qwen-1.5B"

# ---------- 固定系统提示语（严格按任务图片原文） ----------
SYSTEM_PROMPT = "你是中医国医大师、中医药专家，擅长中医病症诊断与治疗。根据用户问题回答中医相关的内容。"

RANDOM_SEED = 42
