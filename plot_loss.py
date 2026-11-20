# -*- coding: utf-8 -*-
"""绘制训练 loss 曲线图（用于项目报告第六章插图）。

输入：results/training_log.json
输出：results/loss_curve.png（300 dpi）
"""
import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

# 中文字体：优先微软雅黑，其次黑体（Windows 系统自带）
for fname in ("Microsoft YaHei", "SimHei"):
    try:
        font_manager.findfont(fname, fallback_to_default=False)
        plt.rcParams["font.sans-serif"] = [fname]
        break
    except Exception:
        continue
plt.rcParams["axes.unicode_minus"] = False

BASE = Path(__file__).resolve().parent.parent
log_path = BASE / "results" / "training_log.json"
out_path = BASE / "results" / "loss_curve.png"

with open(log_path, "r", encoding="utf-8") as f:
    history = json.load(f)

train = [(e["step"], e["loss"]) for e in history if "loss" in e]
evals = [(e["step"], e["eval_loss"]) for e in history if "eval_loss" in e]
runtime = next((e.get("train_runtime") for e in history if "train_runtime" in e), None)

tx, ty = zip(*train)
ex, ey = zip(*evals)

fig, ax = plt.subplots(figsize=(8, 4.5), dpi=300)
ax.plot(tx, ty, "-o", color="#2b6cb0", markersize=4, linewidth=1.6, label="训练 loss")
ax.plot(ex, ey, "*", color="#c53030", markersize=14, label="验证 eval_loss", zorder=5)

# 标注首末点数值
ax.annotate(f"{ty[0]:.3f}", (tx[0], ty[0]), textcoords="offset points",
            xytext=(6, 8), fontsize=9, color="#2b6cb0")
ax.annotate(f"{ty[-1]:.3f}", (tx[-1], ty[-1]), textcoords="offset points",
            xytext=(-4, 10), fontsize=9, color="#2b6cb0")
for x, y in evals:
    ax.annotate(f"{y:.3f}", (x, y), textcoords="offset points",
                xytext=(0, 10), fontsize=9, color="#c53030", ha="center")

ax.set_xlabel("优化步（step）", fontsize=11)
ax.set_ylabel("损失（loss）", fontsize=11)
title = "LoRA 微调训练损失曲线（600 条 × 3 轮 = 225 步，CPU）"
if runtime:
    title += f"  总耗时 {runtime / 60:.0f} 分钟"
ax.set_title(title, fontsize=12)
ax.grid(True, linestyle="--", alpha=0.4)
ax.legend(fontsize=10, loc="upper right")
fig.tight_layout()
fig.savefig(out_path)
print(f"已保存：{out_path}")
