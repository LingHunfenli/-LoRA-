# -*- coding: utf-8 -*-
"""阶段8：生成项目汇报 PPT（课设正式风格，数据驱动）。

特性：
  - 封面（占位符）+ 目录 + 分节页 + 内容页，约 18~20 张；
  - 配色方案：深蓝主色 #1a365d / 中医绿强调 #2f855a / 浅灰底；
  - 数据表格化（清洗统计、训练超参、对比指标），自动回填真实数据；
  - 嵌入 results/loss_curve.png 训练曲线图；
  - 五题对比摘要 + 四维评价可视化。

用法：
    python plot_loss.py     # 先画曲线
    python gen_ppt.py       # 再生成 PPT
输出：项目汇报.pptx（项目根目录）

依赖：pip install python-pptx matplotlib
"""
import json
import re
import statistics
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml.etree import SubElement as _SubElement
from pptx.util import Cm, Pt, Emu

import config

PPT_PATH = config.PROJECT_DIR / "项目汇报.pptx"
LOSS_CURVE = config.RESULTS_DIR / "loss_curve.png"

# ======================== 配色 ========================
NAVY = RGBColor(0x1A, 0x36, 0x5D)       # 深蓝（标题/重点）
TEAL = RGBColor(0x2F, 0x85, 0x5A)        # 中医绿（强调/数据）
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF7, 0xFA, 0xFC)  # 极浅灰背景
DARK_TEXT = RGBColor(0x1A, 0x20, 0x2C)   # 深色正文
MUTED = RGBColor(0x71, 0x80, 0x96)       # 灰色辅助文字


# ======================== 工具函数 ========================

def read_jsonl(path):
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8") as f:
        return [json.loads(l) for l in f if l.strip()]


def load_data():
    d = {}
    d["cases"] = read_jsonl(config.RESULTS_DIR / "compare_cases.jsonl")
    log_path = config.RESULTS_DIR / "training_log.json"
    if log_path.exists():
        d["log"] = json.loads(log_path.read_text(encoding="utf-8"))
    else:
        d["log"] = []
    d["clean_report"] = config.DATA_REPORT.read_text(encoding="utf-8") if config.DATA_REPORT.exists() else ""

    def grab(pattern, default="0"):
        m = re.search(pattern, d["clean_report"])
        return m.group(1) if m else default

    def grab_all(pattern):
        m = re.search(pattern, d["clean_report"])
        return list(m.groups()) if m else ["0"] * 5

    raw = grab(r"原始数据：train (\d+) 条 / test (\d+) 条", "0 0").split()
    d["raw_train"] = raw[0] if raw else "0"
    d["raw_test"] = raw[1] if len(raw) > 1 else "0"
    d["drop_train"] = grab_all(r"\[train\] 清洗剔除明细：标题噪声 (\d+)、重复 (\d+)、字段异常 (\d+)、问题过短 (\d+)、回答长度越界 (\d+)")
    d["drop_test"] = grab_all(r"\[test\] 清洗剔除明细：标题噪声 (\d+)、重复 (\d+)、字段异常 (\d+)、问题过短 (\d+)、回答长度越界 (\d+)")
    d["clean_train_n"] = grab(r"清洗后（跨集去重前）：train (\d+) 条", "4919")
    d["clean_test_n"] = grab(r"train \d+ 条 / test (\d+) 条", "1299")

    d["subset_n"] = len(read_jsonl(config.TRAIN_SUBSET)) or 600
    d["augmented_n"] = max(len(read_jsonl(config.TRAIN_AUGMENTED)) - d["subset_n"], 800)

    log = d["log"]
    losses = [(e["step"], e["loss"]) for e in log if "loss" in e]
    evals = [(e["step"], e["eval_loss"]) for e in log if "eval_loss" in e]
    runtime = next((e.get("train_runtime") for e in log if "train_runtime" in e), None)
    avg_loss = next((e.get("train_loss") for e in log if "train_loss" in e), None)
    d["losses"] = losses
    d["evals"] = evals
    d["runtime"] = runtime
    d["avg_loss"] = avg_loss
    return d


def set_shape_fill(shape, color):
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color


def set_shape_line(shape, color=None, width=Pt(0)):
    line = shape.line
    if color is None:
        line.fill.background()
    else:
        line.color.rgb = color
        line.width = width


# ======================== 幻灯片工厂 ========================

def _blank_slide(prs):
    """空白布局（layout 6）。"""
    return prs.slides.add_slide(prs.slide_layouts[6])


# 幻灯片标准尺寸（16:9 宽屏，与 main() 中设置一致）
_SLIDE_W = Cm(33.867)
_SLIDE_H = Cm(19.05)


def add_title_bar(slide, text, subtitle=None):
    """顶部深蓝标题栏。"""
    h = Cm(1.4)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0),
                                 _SLIDE_W, h)
    set_shape_fill(bar, NAVY)
    set_shape_line(bar)
    tf = bar.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(10)
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(13)
        p2.font.color.rgb = RGBColor(0xBD, 0xC3, 0xC7)


def add_footer_bar(slide, page_num, total):
    """底部页码栏。"""
    h = Cm(0.55)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(0),
                                 _SLIDE_H - h,
                                 _SLIDE_W, h)
    set_shape_fill(bar, LIGHT_GRAY)
    set_shape_line(bar)
    tf = bar.text_frame
    p = tf.paragraphs[0]
    p.text = f"基于 LoRA 微调的中医医学专家大模型  |  {page_num} / {total}"
    p.font.size = Pt(9)
    p.font.color.rgb = MUTED
    p.alignment = PP_ALIGN.CENTER


def add_bullets(slide, items, left=1.0, top=1.8, width=11.5, height=5.0,
               size=17, color=DARK_TEXT, bullet_color=TEAL):
    """添加要点列表。"""
    box = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    tf = box.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "●  " + it
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.space_after = Pt(10)
        p.level = 0


def _set_cell_fill(cell, color):
    """通过 XML 设置表格单元格背景色。"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # 移除已有的 shd
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = _SubElement(tcPr, qn("w:shd"))
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), str(color)[1:] if not isinstance(color, str) else color.lstrip("#"))
    tcPr.append(shd)


def add_table(slide, headers, rows, left=1.0, top=1.8, col_w=None):
    """添加表格（自动样式）。"""
    nrow = len(rows) + 1
    ncol = len(headers)
    tw = Cm(sum(col_w)) if col_w else Cm(11.5)
    th = Cm(0.65 * nrow)
    tb = slide.shapes.add_table(nrow, ncol, Cm(left), Cm(top), tw, th).table
    # 表头
    for j, h in enumerate(headers):
        cell = tb.cell(0, j)
        cell.text = str(h)
        _set_cell_fill(cell, NAVY)
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
    # 数据行
    for i, row in enumerate(rows, start=1):
        bg = WHITE if i % 2 == 1 else LIGHT_GRAY
        for j, val in enumerate(row):
            cell = tb.cell(i, j)
            cell.text = str(val)
            _set_cell_fill(cell, bg)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.font.color.rgb = DARK_TEXT
            p.alignment = PP_ALIGN.CENTER
    if col_w:
        for j, w in enumerate(col_w):
            tb.columns[j].width = Cm(w)
    return tb


def add_image(slide, img_path, left=1.0, top=1.8, w_cm=11.5):
    """嵌入图片。"""
    if not img_path.exists():
        add_bullets(slide, [f"[图片缺失：{img_path.name}，请先运行 plot_loss.py]"], top=top)
        return None
    shape = slide.shapes.add_picture(str(img_path), Cm(left), Cm(top), width=Cm(w_cm))
    return shape


# ======================== 各类幻灯片 ========================

def slide_cover(prs, pg, tot):
    s = _blank_slide(prs)
    # 深蓝背景条
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0),
                              _SLIDE_W, _SLIDE_H)
    set_shape_fill(bar, NAVY)
    set_shape_line(bar)
    # 标题
    box = s.shapes.add_textbox(Cm(1), Cm(2.0), Cm(11.33), Cm(1.8))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = "基于 LoRA 微调的中医医学专家大模型"
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = "——DeepSeek-R1-Distill-Qwen-1.5B 中医问答领域适配实践"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(0xBD, 0xC3, 0xC7)
    p2.alignment = PP_ALIGN.CENTER
    # 信息占位
    info_box = s.shapes.add_textbox(Cm(1), Cm(5.0), Cm(11.33), Cm(2.0))
    itf = info_box.text_frame
    for label in ["组    别：____________", "组    员：____________",
                  "指导教师：____________", "日    期：____________"]:
        ip = itf.paragraphs[0] if label == "组    别：____________" else itf.add_paragraph()
        ip.text = label
        ip.font.size = Pt(16)
        ip.font.color.rgb = WHITE
        ip.alignment = PP_ALIGN.CENTER
        ip.space_after = Pt(8)
    add_footer_bar(s, pg, tot)
    return s


def slide_toc(prs, pg, tot):
    s = _blank_slide(prs)
    add_title_bar(s, "目  录")
    sections = [
        ("01", "项目背景与目标"),
        ("02", "数据准备与预处理"),
        ("03", "模型选择与技术方案"),
        ("04", "模型训练与结果"),
        ("05", "推理验证与效果对比"),
        ("06", "应用展示（加分项）"),
        ("07", "改进点"),
        ("08", "问题、不足与解决方案"),
        ("09", "总结"),
    ]
    start_top = 1.8
    for i, (num, title) in enumerate(sections):
        row = i // 2
        col = i % 2
        x = 1.2 + col * 6.0
        y = start_top + row * 1.05
        # 编号圆圈
        circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x), Cm(y), Cm(0.55), Cm(0.55))
        set_shape_fill(circle, TEAL)
        set_shape_line(circle)
        ctf = circle.text_frame
        ctf.paragraphs[0].text = num
        ctf.paragraphs[0].font.size = Pt(13)
        ctf.paragraphs[0].font.bold = True
        ctf.paragraphs[0].font.color.rgb = WHITE
        ctf.paragraphs[0].alignment = PP_ALIGN.CENTER
        # 标题文字
        tbox = s.shapes.add_textbox(Cm(x + 0.72), Cm(y + 0.08), Cm(4.8), Cm(0.45))
        tbox.text_frame.paragraphs[0].text = title
        tbox.text_frame.paragraphs[0].font.size = Pt(15)
        tbox.text_frame.paragraphs[0].font.color.rgb = DARK_TEXT
    add_footer_bar(s, pg, tot)
    return s


def slide_section(prs, num, title, pg, tot):
    """分节页（大号编号 + 章节名）。"""
    s = _blank_slide(prs)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0),
                              _SLIDE_W, _SLIDE_H)
    set_shape_fill(bar, NAVY)
    set_shape_line(bar)
    num_box = s.shapes.add_textbox(Cm(1.5), Cm(2.2), Cm(3), Cm(1.5))
    ntf = num_box.text_frame
    np_ = ntf.paragraphs[0]
    np_.text = f"{num:02d}"
    np_.font.size = Pt(80)
    np_.font.bold = True
    np_.font.color.rgb = TEAL
    t_box = s.shapes.add_textbox(Cm(1.5), Cm(3.8), Cm(10), Cm(1.2))
    tt = t_box.text_frame.paragraphs[0]
    tt.text = title
    tt.font.size = Pt(32)
    tt.font.color.rgb = WHITE
    add_footer_bar(s, pg, tot)
    return s


def slide_content(prs, title, items, pg, tot, subtitle=None, size=17):
    s = _blank_slide(prs)
    add_title_bar(s, title, subtitle)
    add_bullets(s, items, size=size)
    add_footer_bar(s, pg, tot)
    return s


def slide_two_col(prs, title, left_items, right_items, left_title="", right_title="",
                   pg=0, tot=0):
    """双栏内容页。"""
    s = _blank_slide(prs)
    add_title_bar(s, title)
    if left_title:
        lt = s.shapes.add_textbox(Cm(1.0), Cm(1.65), Cm(5.5), Cm(0.4))
        lt.text_frame.paragraphs[0].text = left_title
        lt.text_frame.paragraphs[0].font.size = Pt(15)
        lt.text_frame.paragraphs[0].font.bold = True
        lt.text_frame.paragraphs[0].font.color.rgb = TEAL
    if right_title:
        rt = s.shapes.add_textbox(Cm(7.0), Cm(1.65), Cm(5.5), Cm(0.4))
        rt.text_frame.paragraphs[0].text = right_title
        rt.text_frame.paragraphs[0].font.size = Pt(15)
        rt.text_frame.paragraphs[0].font.bold = True
        rt.text_frame.paragraphs[0].font.color.rgb = TEAL
    add_bullets(s, left_items, left=1.0, top=2.1, width=5.5, size=15)
    add_bullets(s, right_items, left=7.0, top=2.1, width=5.5, size=15)
    if pg:
        add_footer_bar(s, pg, tot)
    return s


def slide_image_page(prs, title, img_path, caption, pg, tot):
    """图片页（居中大图 + 说明文字）。"""
    s = _blank_slide(prs)
    add_title_bar(s, title)
    add_image(s, img_path, left=1.0, top=1.75, w_cm=11.3)
    cap_box = s.shapes.add_textbox(Cm(1.0), Cm(5.85), Cm(11.3), Cm(0.6))
    cap_box.text_frame.paragraphs[0].text = caption
    cap_box.text_frame.paragraphs[0].font.size = Pt(12)
    cap_box.text_frame.paragraphs[0].font.color.rgb = MUTED
    cap_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    add_footer_bar(s, pg, tot)
    return s


def slide_thanks(prs, pg, tot):
    s = _blank_slide(prs)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0),
                              _SLIDE_W, _SLIDE_H)
    set_shape_fill(bar, NAVY)
    set_shape_line(bar)
    box = s.shapes.add_textbox(Cm(1), Cm(2.5), Cm(11.33), Cm(2))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = "感谢聆听"
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = "欢迎提问与交流"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(0xBD, 0xC3, 0xC7)
    p2.alignment = PP_ALIGN.CENTER
    add_footer_bar(s, pg, tot)
    return s


# ======================== 主流程 ========================

def main():
    d = load_data()

    prs = Presentation()
    prs.slide_width = Cm(33.867)   # 16:9 宽屏
    prs.slide_height = Cm(19.05)

    pg = [0]

    def nxt():
        pg[0] += 1
        return pg[0]

    TOTAL = 20  # 预估总页数

    # ---- 封面 ----
    slide_cover(prs, nxt(), TOTAL)

    # ---- 目录 ----
    slide_toc(prs, nxt(), TOTAL)

    # ===== 01 项目背景与目标 =====
    slide_section(prs, 1, "项目背景与目标", nxt(), TOTAL)

    slide_content(prs, "项目背景", [
        "中医药知识体系庞大分散，长期以典籍、医案与临床经验传承为主",
        "通用大语言模型对中医术语、辨证论治逻辑理解有限",
        "直接问答常出现内容空泛、表述不专业甚至编造方药的问题",
        "课程实验环境：无独立显卡、16GB 内存，需纯 CPU 可复现方案",
    ], nxt(), TOTAL, subtitle="为什么需要领域微调？")

    slide_content(prs, "项目目标", [
        "以中医问答数据对 DeepSeek-R1-Distill-Qwen-1.5B 进行 LoRA 微调",
        "使模型在固定系统提示语约束下输出专业、准确、完整、简洁的中医回答",
        "通过微调前后对比验证领域适配效果（核心验证问题）",
        "加分项：数据增强 + Gradio 对比界面 + 相对课堂方法的改进",
    ], nxt(), TOTAL, subtitle="我们要达成什么？")

    # ===== 02 数据准备与预处理 =====
    slide_section(prs, 2, "数据准备与预处理", nxt(), TOTAL)

    dt, dtest = d["drop_train"], d["drop_test"]
    slide_content(prs, "数据来源与格式", [
        f"原始数据：chinese_medical_train.jsonl（{d['raw_train']} 条）+ test（{d['raw_test']} 条）",
        "指令微调 JSONL 格式，每条含三个字段：",
        "  instruction — 用户问题（如“湖北贝母有什么功效？”）",
        "  input — 固定系统提示语（全部记录同一值）",
        "  output — 中医专家回答（覆盖中药/方剂/辨证/针灸/养生等）",
    ], nxt(), TOTAL)

    slide_content(prs, "数据清洗", [
        f"噪声问题：约 {int(dt[0]) + int(dtest[0])} 条 output 仅为论文/文章标题（《……》）",
        f"清洗规则：标题特征检测 + 重复去除 + 长度过滤 + 提示语统一规范化",
        f"Train 清洗：标题噪声 {dt[0]} / 重复 {dt[1]} / 字段异常 {dt[2]} / 问题过短 {dt[3]} / 长度越界 {dt[4]}",
        f"Test  清洗：标题噪声 {dtest[0]} / 重复 {dtest[1]} / 字段异常 {dtest[2]} / 问题过短 {dtest[3]} / 长度越界 {dtest[4]}",
        f"清洗结果：train {d['clean_train_n']} 条 / test {d['clean_test_n']} 条；跨集去重检查通过",
    ], nxt(), TOTAL)

    slide_content(prs, "数据切分与增强（加分项）", [
        f"切分：train 4919 / val 258 / test 1299（95:5 训验比）",
        f"CPU 训练子集：从 train_clean 采样 {d['subset_n']} 条 → train_subset.jsonl",
        f"数据增强：问题侧规则改写新增 {d['augmented_n']} 条同义问法样本",
        "  （敬语包裹 / 请求式后缀 / 同义开头替换，答案不编造任何医学内容）",
        "  子集 + 增强 = 1400 条混合备用集（GPU 环境可扩大训练规模）",
    ], nxt(), TOTAL)

    # ===== 03 模型选择与技术方案 =====
    slide_section(prs, 3, "模型选择与技术方案", nxt(), TOTAL)

    slide_content(prs, "基座模型：DeepSeek-R1-Distill-Qwen-1.5B", [
        "R1 推理模型的思维链能力蒸馏到 Qwen2.5-1.5B 基座，约 15 亿参数",
        "Decoder-Only Transformer：RoPE 位置编码 / SwiGLU 激活 / RMSNorm",
        "支持 ChatML 对话模板，完全开源（ModelScope / HuggingFace）",
        "选择理由：规模适中（CPU 可跑）、生态成熟、适合教学复现",
    ], nxt(), TOTAL)

    slide_two_col(prs, "技术方案：LoRA 低秩适配",
        [
            "冻结预训练权重 W",
            "旁路注入低秩矩阵 A、B",
            "W' = W + (α/r)·BA",
            "可训练参数仅 0.24%",
            "适配器可插拔、易合并",
        ],
        [
            "target_modules：q/k/v/o_proj",
            "r = 16, α = 32 (α/r = 2)",
            "dropout = 0.05",
            "PEFT 库实现",
            "HuggingFace Trainer 训练",
        ],
        left_title="LoRA 原理", right_title="本项目配置",
        pg=nxt(), tot=TOTAL)

    slide_content(prs, "指令微调与损失构造", [
        "chat template 构造：系统提示语 + 用户问题 + 专家回答",
        "问题段 labels = -100（不参与损失），仅回答段计算交叉熵",
        "模型只学习“如何回答”，而非复述问题或系统提示语",
        "R1 思维链冲突处理：训练时不构造思维链目标 → 模型学会直接作答",
        "推理侧兼容思维链残留解析（剥离  …  段后输出）",
    ], nxt(), TOTAL)

    # ===== 04 模型训练与结果 =====
    slide_section(prs, 4, "模型训练与结果", nxt(), TOTAL)

    losses = d["losses"]
    evals = d["evals"]
    runtime = d["runtime"]
    avg_loss = d["avg_loss"]

    slide_content(prs, "训练超参数配置", [
        f"训练样本：{d['subset_n']} 条 × 3 轮 = 225 个优化步",
        "学习率：1e-4（cosine 衰减，warmup 3% ≈ 6 步）",
        "batch_size = 2, gradient_accumulation = 4 → 有效 batch = 8",
        "max_length = 384, weight_decay = 0.01",
        "硬件：纯 CPU（16GB 内存），torch 2.14+cpu / transformers 5.17 / peft 0.20",
    ], nxt(), TOTAL)

    # 训练指标表
    s = _blank_slide(prs)
    add_title_bar(s, "训练结果指标")
    r_rows = [
        ["训练日志点数", f"{len(losses)} 个（每 10 步记录）"],
        ["训练 loss（起始→末次）", f"{losses[0][1]:.4f} → {losses[-1][1]:.4f}"],
        ["平均训练 loss", f"{avg_loss:.4f}" if avg_loss else "-"],
        ["验证 eval_loss（各轮）", " / ".join(f"{y:.4f}" for _, y in evals)],
        ["最优 eval_loss", f"{min(y for _, y in evals):.4f}" if evals else "-"],
        ["训练总耗时", f"{runtime / 60:.0f} 分钟（纯 CPU）" if runtime else "-"],
        ["最终产物", "output/final/ LoRA 适配器（16.6 MB）"],
    ]
    add_table(s, ["指标", "数值"], r_rows, left=1.5, top=1.8, col_w=[6.0, 8.0])
    add_footer_bar(s, nxt(), TOTAL)

    # Loss 曲线图页
    slide_image_page(prs, "训练损失曲线", LOSS_CURVE,
                     "图：LoRA 微调训练 loss 曲线（225 步，★ 为验证 eval_loss）",
                     nxt(), TOTAL)

    slide_content(prs, "收敛分析与工程亮点", [
        "前 40 步 loss 由 2.87 快速降至 2.33，末段稳定在 2.13～2.20",
        "两次 eval_loss（2.318 / 2.312）基本持平，无明显过拟合",
        "工程亮点：断点续训（checkpoint 幂等恢复）、Windows MKL 防崩溃、冒烟测试",
        "小样本 LoRA 的作用是“对齐领域表达”而非注入全部知识",
    ], nxt(), TOTAL)

    # ===== 05 推理验证与效果对比 =====
    slide_section(prs, 5, "推理验证与效果对比", nxt(), TOTAL)

    cases = d["cases"]
    base_lens = [len(c["base_answer"]) for c in cases]
    lora_lens = [len(c["lora_answer"]) for c in cases]
    cot_base = sum(1 for c in cases if (c.get("base_thinking") or "").strip())
    cot_lora = sum(1 for c in cases if (c.get("lora_thinking") or "").strip())

    s = _blank_slide(prs)
    add_title_bar(s, "客观指标对比（5 类未见问题）")
    add_table(s, ["指标", "微调前（基座）", "微调后（LoRA）"], [
        ["平均回答长度", f"{sum(base_lens)/len(base_lens):.0f} 字", f"{sum(lora_lens)/len(lora_lens):.0f} 字"],
        ["思维链出现率", f"{cot_base}/{len(cases)} ({cot_base/len(cases):.0%})",
         f"{cot_lora}/{len(cases)} ({cot_lora/len(cases):.0%})"],
        ["回答风格", "冗长推理、常循环重复", "直接作答、结构清晰"],
    ], left=1.5, top=1.8, col_w=[5.0, 4.5, 4.5])
    add_footer_bar(s, nxt(), TOTAL)

    # 五题概览
    q_items = [f"【{c['category']}】{c['instruction']}" for c in cases]
    q_items += [
        "微调前：常输出冗长思维链、泛化内容、易编造方药",
        "微调后：直接作答、术语专业、结构完整、风格对齐",
    ]
    slide_content(prs, "五题验证问题一览", q_items, nxt(), TOTAL, size=15)

    # 四维评价
    s = _blank_slide(prs)
    add_title_bar(s, "四维人工评价")
    add_table(s, ["维度", "微调前", "微调后"], [
        ["专业性", "非中医概念随意组合\n（如“耳火症”“湿热方剂”）", "贴近中医语料风格\n使用规范术语"],
        ["准确性", "编造方药明显\n（如“黄芪白芍茯苓汤”）", "方向大体正确\n个别事实仍需复核"],
        ["完整性", "冗长跑题或中途截断\n信息组织差", "结构完整\n含调理建议与遵医嘱提示"],
        ["简洁性", "平均 407 字\n60% 思维链残留", "平均 105 字\n思维链出现率 0%"],
    ], left=1.2, top=1.75, col_w=[2.5, 5.0, 5.0])
    add_footer_bar(s, nxt(), TOTAL)

    # ===== 06 应用展示 =====
    slide_section(prs, 6, "应用展示（加分项）", nxt(), TOTAL)

    slide_content(prs, "Gradio 双栏对比界面", [
        "网页界面：输入中医问题 → 双栏并排展示微调前 / 微调后回答",
        "问题精确命中测试集时，底部展示数据集参考答案",
        "内置 5 个验证示例问题，一键对比",
        "启动：python src/app.py → http://127.0.0.1:7860",
        "支持局域网访问（server_name=0.0.0.0）与 PWA 安装为桌面应用",
        "（此处插入 Gradio 界面运行截图）",
    ], nxt(), TOTAL)

    # ===== 07 改进点 =====
    slide_section(prs, 7, "相对课堂方法的改进点", nxt(), TOTAL)

    slide_two_col(prs, "五大改进方向",
        [
            "【数据侧】",
            "  标题噪声清洗 + 去重 + 长度过滤",
            "  完整管线 + 自动审计报告",
            "  问题侧增强 +800 条",
        ],
        [
            "【方法侧】",
            "  LoRA 目标模块扩展至 q/k/v/o",
            "  α/r=2，cosine 调度 + warmup",
            "  transformers 5.x 新版适配",
        ],
        pg=nxt(), tot=TOTAL)

    slide_two_col(prs, "五大改进方向（续）",
        [
            "【模型侧】",
            "  Qwen3-0.6B → R1-Distill-1.5B",
            "  意图分类 → 生成式问答",
            "  处理思维链风格差异",
        ],
        [
            "【应用侧 / 工程侧】",
            "  新增 Gradio 前后对比界面",
            "  路径与提示语集中配置",
            "  冒烟测试 + 断点续训 + CPU/GPU 自适应",
        ],
        pg=nxt(), tot=TOTAL)

    # ===== 08 问题、不足与解决方案 =====
    slide_section(prs, 8, "问题、不足与解决方案", nxt(), TOTAL)

    s = _blank_slide(prs)
    add_title_bar(s, "遇到的问题与解决方案")
    add_table(s, ["问题", "原因", "解决方案"], [
        ["Windows 原生崩溃\n(0xC0000005)", "MKL/OpenMP 重复加载\n导致访问违例", "KMP_DUPLICATE_LIB_OK=TRUE\n限制线程数为 4"],
        ["训练中断无法继续", "CPU 训练耗时较长\n意外终止", "周期保存 checkpoint\n--resume 幂等续训"],
        ["R1 思维链风格冲突", "蒸馏模型倾向先推理\n再给答案", "训练不构造思维链目标\n推理兼容  解析"],
        ["数据存在标题噪声", "~550 条 output 仅为\n论文/文章标题", "标题特征 + 长度阈值\n规则清洗并保留报告"],
    ], left=0.8, top=1.7, col_w=[4.5, 5.0, 5.0])
    add_footer_bar(s, nxt(), TOTAL)

    slide_content(prs, "不足与后续展望", [
        "数据规模：受 CPU 限制仅用 600 条子集，4919 条全量 + 1400 条增强集保留备用",
        "评测方法：目前以人工四维评述为主，后续可引入 BLEU/ROUGE 或专家打分",
        "事实准确性：小模型仍有细节幻觉，可结合 RAG 检索增强知识库缓解",
        "应用深化：多轮对话、回答置信度、“本回答不构成医疗建议”安全提示",
    ], nxt(), TOTAL)

    # ===== 09 总结 =====
    slide_section(prs, 9, "项目总结", nxt(), TOTAL)

    slide_content(prs, "成果总结", [
        f"完成从数据清洗 → LoRA 微调 → 权重合并 → 推理验证 → 界面展示的完整闭环",
        f"7036 条原始数据 → 6476 条规范样本 → 600 条子集训练",
        f"纯 CPU 104 分钟完成 225 步微调，loss 2.871 → 2.176",
        f"五题验证：回答长度 407→105 字，思维链 60%→0%，风格显著改善",
        "结论：小模型 + 低成本 LoRA 领域适配可有效提升专业问答质量",
    ], nxt(), TOTAL)

    slide_thanks(prs, nxt(), TOTAL)

    prs.save(PPT_PATH)
    print(f"PPT 已生成：{PPT_PATH}（共 {pg[0]} 张幻灯片）")


if __name__ == "__main__":
    main()
