# -*- coding: utf-8 -*-
"""生成答辩讲稿文档（.docx），组员可直接照念。

用法：
    python gen_defense.py
输出：答辩讲稿.docx（项目根目录）
"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SONG, HEI, WESTERN = "宋体", "黑体", "Times New Roman"

REPORT_PATH = Path(__file__).resolve().parent.parent / "答辩讲稿.docx"


def set_run_font(run, cn_font=SONG, size=12, bold=False, color=None):
    run.font.name = WESTERN
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
    if color:
        run.font.color.rgb = RGBColor(*color)


def style_fonts(doc):
    normal = doc.styles["Normal"]
    normal.font.name = WESTERN
    normal.font.size = Pt(12)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), SONG)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.first_line_indent = Pt(24)
    pf.space_after = Pt(0)

    spec = {"Heading 1": (16, 14, 8), "Heading 2": (14, 10, 6), "Heading 3": (12, 8, 4)}
    for name, (size, before, after) in spec.items():
        st = doc.styles[name]
        st.font.name = WESTERN
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            from docx.oxml import OxmlElement
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        rf.set(qn("w:eastAsia"), HEI)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.first_line_indent = Pt(0)
        st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE


def para(doc, text, indent=True, bold=False, size=12, align=None):
    p = doc.add_paragraph()
    if not indent:
        p.paragraph_format.first_line_indent = Pt(0)
    if align is not None:
        p.alignment = align
    set_run_font(p.add_run(text), size=size, bold=bold)
    return p


def quote(doc, text, size=11):
    """引用/讲稿原文，浅灰底纹风格（缩进+左侧竖线模拟）。"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.left_indent = Pt(24)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    r = p.add_run("【讲稿】" + text)
    set_run_font(r, size=size)
    r.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)


def note(doc, text, size=10):
    """注意事项（灰色小字）。"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.left_indent = Pt(24)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("💡 " + text)
    set_run_font(r, size=size)
    r.font.color.rgb = RGBColor(0x71, 0x80, 0x96)


def bullet(doc, items, size=12):
    for it in items:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Pt(24)
        set_run_font(p.add_run("• " + it), size=size)


def qa(doc, question, answer, size=12):
    """预设问题 + 答案格式。"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(8)
    set_run_font(p.add_run("Q：" + question), size=size, bold=True,
                 color=(0x1A, 0x36, 0x5D))
    p2 = doc.add_paragraph()
    p2.paragraph_format.first_line_indent = Pt(0)
    p2.paragraph_format.left_indent = Pt(24)
    set_run_font(p2.add_run("A：" + answer), size=size)


# ======================== 正文 ========================

def cover(doc):
    for _ in range(4):
        doc.add_paragraph()
    para(doc, "项目答辩讲稿", indent=False, bold=True, size=28,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    para(doc, "基于 LoRA 微调的中医医学专家大模型", indent=False, size=16,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "——DeepSeek-R1-Distill-Qwen-1.5B 领域适配实践", indent=False, size=14,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(6):
        doc.add_paragraph()
    para(doc, "（本讲稿供答辩时使用，建议提前通读一遍，标注个人分工）",
         indent=False, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()


def ch_overview(doc):
    doc.add_heading("一、开场白（30 秒）", 1)
    para(doc, "（选一个组员说，自然大方即可）")
    quote(doc, "各位老师好，我们组的项目是《基于 LoRA 微调的中医医学专家大模型》。"
               "简单来说，就是让一个通用大模型学会像中医专家一样回答问题。"
               "我们从数据清洗做起，用 LoRA 方法在 CPU 上完成了微调，"
               "最后通过前后对比验证了效果。下面由我为大家汇报。")
    note(doc, "时间控制：30 秒以内说完，不要念稿痕迹太重。")


def ch_1min(doc):
    doc.add_heading("二、项目一句话概括（1 分钟版）", 1)
    quote(doc, "通用大模型虽然什么都能聊，但问到中医问题经常答非所问、甚至编造方药。"
               "我们的思路是：既然模型不懂中医，那就用中医问答数据去教它。"
               "方法上选择了 LoRA——一种只训练模型 0.24% 参数的低成本微调技术。"
               "数据上，我们从 7000 多条原始问答中清洗出规范样本，"
               "在纯 CPU 环境下完成了 3 轮训练。"
               "验证结果显示，微调后模型的回答从平均 407 字降到 105 字，"
               "不再输出冗长的思维链，回答风格也更贴合中医专家口吻。")
    note(doc, "如果老师只给 1 分钟，就用这段；时间充裕则展开后面的详细版。")


def ch_background(doc):
    doc.add_heading("三、项目背景与动机（详细版）", 1)
    quote(doc, "中医药知识体系庞大，长期以来靠典籍、医案和临床经验传承。"
               "通用大模型虽然具备广泛的语言能力，但对中医术语、辨证论治逻辑"
               "和方药配伍的理解很有限。直接问它中医问题，经常出现三种情况："
               "一是内容空泛，说了等于没说；二是表述不专业，术语乱用；"
               "三是编造方药，给出根本不存在的药方——这在医疗场景是很危险的。")
    quote(doc, "所以我们的目标是：让模型在固定系统提示语“你是中医国医大师”的约束下，"
               "输出专业、准确、完整、简洁的中医回答。"
               "考虑到我们没有 GPU，选择了 LoRA 这种参数高效微调方法，"
               "只训练 0.24% 的参数就能达到领域适配的效果。")


def ch_principle(doc):
    doc.add_heading("四、核心技术原理（老师最可能问的部分）", 1)

    doc.add_heading("（一）LoRA 是什么？", 2)
    quote(doc, "LoRA 全称 Low-Rank Adaptation，低秩适配。它的核心思想是："
               "大模型已经学了很多通用知识，我们不需要重新训练所有参数，"
               "只需要在关键位置加一个很小的“补丁”，让它学会领域特定的表达方式。")
    quote(doc, "具体来说，LoRA 冻结了预训练权重 W，在旁边注入两个低秩矩阵 A 和 B。"
               "用公式表示就是 W' = W + (α/r)·BA。其中 r 远小于隐藏层维度，"
               "所以可训练参数量只有原模型的 0.24%——大约 436 万个参数，"
               "而全模型有 15 亿参数。这样即使在没有 GPU 的电脑上也能完成训练。")
    note(doc, "如果老师追问“为什么选 q/k/v/o_proj”，回答：这些是注意力机制中的投影矩阵，"
              "对模型理解输入和生成输出最关键，课堂方法只选了 q/k/v，我们扩展到了 o_proj。")

    doc.add_heading("（二）为什么选 DeepSeek-R1-Distill-Qwen-1.5B？", 2)
    quote(doc, "三个原因：第一，规模适中，1.5B 参数在纯 CPU 上也能跑，"
               "符合我们的实验条件；第二，它是 DeepSeek 把 R1 推理模型的能力"
               "蒸馏到 Qwen2.5 小模型上的成果，本身有一定推理能力；"
               "第三，完全开源，ModelScope 上可以直接下载，生态成熟，方便复现。")

    doc.add_heading("（三）思维链问题怎么处理的？", 2)
    quote(doc, "R1-Distill 模型有个特点：它习惯先输出 <think> 思维链，再给正式答案。"
               "但我们的训练数据是“简洁直接回答”风格，两者有冲突。"
               "我们的处理策略是：训练时不构造思维链目标，只拼接“系统提示语+问题+直接回答”，"
               "让模型通过微调学会跳过思维链直接作答；"
               "推理侧则对可能残留的思维链做兼容解析，剥离 <think> 标签后再输出。")

    doc.add_heading("（四）损失函数怎么设计的？", 2)
    quote(doc, "我们采用指令微调的标准做法：用 chat template 把“系统提示语+用户问题+专家回答”"
               "拼成一个序列，然后让模型预测下一个词。"
               "关键设计是：问题段的标签设为 -100，不参与损失计算，"
               "只有回答段（包括结尾的 EOS token）计算交叉熵损失。"
               "这样模型只学习“如何回答”，而不会浪费容量去复述问题或系统提示语。")


def ch_process(doc):
    doc.add_heading("五、项目过程（按时间线讲）", 1)

    doc.add_heading("（一）数据处理", 2)
    quote(doc, "原始数据有两份：训练集 5628 条、测试集 1408 条。"
               "每条数据包含三个字段：instruction 是用户问题，input 是固定系统提示语，"
               "output 是中医专家回答。")
    quote(doc, "人工抽查发现数据里有约 550 条噪声——output 只是论文或文章标题，"
               "比如《基于XX的中医研究》，没有实质内容。"
               "我们写了清洗脚本，用“标题特征检测 + 重复去除 + 长度过滤 + 提示语统一”"
               "四步规则清洗，最终得到训练集 4919 条、验证集 258 条、测试集 1299 条。")
    quote(doc, "另外我们还做了数据增强（加分项）：对 600 条 CPU 训练子集，"
               "用敬语包裹、请求式后缀、同义开头替换等方法改写问题，新增 800 条样本。"
               "注意我们只改写了问题，答案完全沿用原始标注，没有编造任何医学内容。")

    doc.add_heading("（二）模型训练", 2)
    quote(doc, "训练配置：600 条样本 × 3 轮 = 225 个优化步，"
               "batch size 2、梯度累积 4，有效 batch 是 8，学习率 1e-4 用 cosine 衰减。"
               "LoRA 参数是 r=16、α=32、dropout=0.05，目标模块覆盖 q/k/v/o 四个投影矩阵。")
    quote(doc, "训练在纯 CPU 上进行，总共花了约 104 分钟。"
               "过程中遇到过两次 Windows 底层崩溃（MKL/OpenMP 冲突导致的访问违例），"
               "我们通过设置 KMP_DUPLICATE_LIB_OK=TRUE 和降低线程数解决，"
               "并且利用 checkpoint 机制实现了断点续训——即使中断了也能从上次保存的地方继续。")
    quote(doc, "训练 loss 从 2.871 降到 2.176，验证集 eval_loss 在 2.318 和 2.312 之间，"
               "基本持平，说明没有明显过拟合。")

    doc.add_heading("（三）验证与对比", 2)
    quote(doc, "验证方法：从测试集（没参与训练的数据）选了 5 个不同类型的问题——"
               "中药功效、方剂推荐、辨证论治、养生药膳、病症解释，"
               "用相同的系统提示语和 greedy 解码，分别调用微调前基座和微调后模型，"
               "与数据集参考答案对照。")
    quote(doc, "结果很明显：微调前平均回答 407 字，60% 的情况会输出思维链，"
               "内容冗长、经常循环重复；微调后平均回答 105 字，思维链出现率降到 0%，"
               "直接作答、结构清晰、风格贴合专家口吻。")
    quote(doc, "举个例子，问“湖北贝母有什么功效”，微调前回答“补血养血和安神助眠”，"
               "完全答错了方向；微调后回答“清热解毒、利湿化湿”，方向基本正确，"
               "虽然个别细节还有偏差，但已经符合中医表述习惯。")


def ch_results(doc):
    doc.add_heading("六、成果展示（配合 PPT 讲）", 1)
    quote(doc, "（切换到 PPT 第 17 页，指着 loss 曲线）"
               "这是训练 loss 曲线，可以看到前 40 步快速下降，后面趋于平稳，"
               "说明模型收敛良好。")
    quote(doc, "（切换到 PPT 第 20-21 页，展示对比表格）"
               "这是五题验证的客观指标对比，平均回答长度从 407 字降到 105 字，"
               "思维链出现率从 60% 降到 0%，改善非常明显。")
    quote(doc, "（切换到 PPT 第 22 页，展示四维评价表）"
               "我们从专业性、准确性、完整性、简洁性四个维度做了人工评价。"
               "微调后模型在四个维度都有显著提升，特别是简洁性——"
               "不再输出冗长的思维链，直接给出专业回答。")
    quote(doc, "（如果展示了 Gradio 界面）"
               "我们还做了一个加分项：Gradio 双栏对比界面，"
               "左边输入问题，右边并排显示微调前后的回答，"
               "方便直观对比。这个界面支持局域网访问，也可以安装为桌面应用。")


def ch_qa(doc):
    doc.add_heading("七、预设问题与回答（Q&A 环节）", 1)
    para(doc, "（老师问什么答什么，下面是高频问题，提前背熟）")

    qa(doc, "LoRA 和全量微调有什么区别？为什么选 LoRA？",
       "全量微调要更新模型所有参数，需要大量显存和算力；LoRA 只训练 0.24% 的参数，"
       "在 CPU 上就能完成。对于我们这种教学场景和资源受限环境，LoRA 是最务实的选择。"
       "而且 LoRA 的适配器可以插拔，不影响原模型，方便对比和部署。")

    qa(doc, "为什么不用更大的模型，比如 7B 或 13B？",
       "主要是硬件限制——我们没有 GPU，16GB 内存跑 7B 模型推理都很勉强，更别说训练。"
       "1.5B 是一个在效果和可行性之间的平衡点。而且我们的目标是验证"
       "“小模型+LoRA 能不能做领域适配”，1.5B 恰好能说明问题。")

    qa(doc, "训练数据只有 600 条，够吗？",
       "对于风格对齐来说够用了。我们的目标不是让模型学会中医知识（基座模型预训练时已经学过），"
       "而是让它学会用中医专家的方式表达。600 条高质量指令样本足以完成这种风格迁移。"
       "当然，如果有 GPU，我们会用全部 4919 条清洗数据 + 1400 条增强数据来训练，效果会更好。")

    qa(doc, "微调后回答还有错误，怎么解决？",
       "这是小模型的固有局限。我们的思路是：第一，扩大训练数据规模；"
       "第二，结合 RAG（检索增强生成），让模型先检索知识库再回答，减少幻觉；"
       "第三，在实际部署时加上“本回答不构成医疗建议”的提示，明确边界。")

    qa(doc, "思维链是怎么去掉的？",
       "不是“去掉”，而是让模型学会不输出。训练时我们不构造思维链目标，"
       "只给“系统提示语+问题+直接回答”的样本，模型自然就学会了直接作答。"
       "推理侧再加一层解析，剥离可能残留的 <think> 标签。")

    qa(doc, "数据清洗具体怎么做的？",
       "我们写了自动化脚本，主要四步：第一，标题特征检测——如果 output 以《开头》结尾"
       "且长度很短，判定为标题噪声；第二，跨集去重——确保训练集和测试集没有重复问题；"
       "第三，长度过滤——问题过短或回答长度越界的样本剔除；"
       "第四，提示语统一——把所有 input 字段规范为同一句系统提示语。"
       "清洗后会生成审计报告，记录每一步剔除了多少条、什么原因。")

    qa(doc, "eval_loss 为什么比 train_loss 还高？",
       "这是正常现象。train_loss 是在训练集上计算的，模型见过这些数据；"
       "eval_loss 是在没见过的验证集上计算的，自然更高。"
       "关键看趋势：两次 eval_loss 分别是 2.318 和 2.312，基本持平，"
       "说明模型没有过拟合，泛化能力可以接受。")

    qa(doc, "Gradio 界面是怎么实现的？",
       "核心思路是“一份权重，两种模式”。我们加载基座模型后，"
       "用 PEFT 的 PeftModel 包装 LoRA 适配器，"
       "需要微调前回答时用 disable_adapter() 临时禁用适配器，"
       "需要微调后回答时正常启用。这样不用加载两份模型，节省内存。"
       "界面支持双栏对比、示例问题、参考答案展示，还可以 PWA 安装为桌面应用。")

    qa(doc, "如果给你更多时间/资源，会怎么改进？",
       "第一，用 GPU 跑全量 4919 条数据 + 1400 条增强数据；"
       "第二，引入 BLEU/ROUGE 等客观评测指标，或者请中医专家打分；"
       "第三，结合 RAG 检索中医药知识库，减少事实性错误；"
       "第四，尝试更大的基座模型（如 3B、7B），对比不同规模的效果。")

    qa(doc, "这个项目和课堂上学的方法有什么区别？（改进点）",
       "五个方面：数据侧，我们建立了完整的清洗管线并做了增强；"
       "方法侧，LoRA 目标模块从 q/k/v 扩展到 q/k/v/o，α/r=2；"
       "模型侧，从 Qwen3-0.6B 意图分类迁移到 R1-Distill-1.5B 生成式问答；"
       "应用侧，新增了 Gradio 前后对比界面；"
       "工程侧，支持冒烟测试、断点续训、CPU/GPU 自适应，全程可在无 GPU 电脑复现。")


def ch_closing(doc):
    doc.add_heading("八、收尾（30 秒）", 1)
    quote(doc, "总结一下：我们完成了从数据清洗、LoRA 微调、权重合并到推理验证、"
               "界面展示的完整闭环。在纯 CPU 环境下，用 600 条数据、104 分钟训练，"
               "让 1.5B 小模型学会了中医专家的回答风格。"
               "微调后回答长度从 407 字降到 105 字，思维链出现率从 60% 降到 0%，"
               "验证了低成本领域适配的有效性。谢谢各位老师，欢迎提问。")
    note(doc, "说完“谢谢”后停顿一下，等老师提问。Q&A 环节谁负责哪个问题提前分好工。")


def ch_tips(doc):
    doc.add_heading("九、答辩注意事项", 1)
    bullet(doc, [
        "提前 10 分钟到，测试 PPT 能不能正常播放、loss 曲线图能不能显示",
        "每人负责一部分：开场+背景（1人）、技术原理（1人）、过程+成果（1人）、Q&A 分工",
        "讲的时候看老师，不要一直盯着屏幕或稿子",
        "老师打断提问时，先听完再回答，不要抢话",
        "不会的问题诚实说“这个我们还没深入研究，后续会考虑”，不要瞎编",
        "时间控制在 8-10 分钟，留 2-3 分钟给 Q&A",
        "PPT 第 22 页的 Gradio 截图提前截好，答辩前贴进去",
        "打印纸质版报告（带封面），每人一份个人总结",
    ])


def main():
    doc = Document()
    style_fonts(doc)
    cover(doc)
    ch_overview(doc)
    doc.add_page_break()
    ch_1min(doc)
    ch_background(doc)
    doc.add_page_break()
    ch_principle(doc)
    doc.add_page_break()
    ch_process(doc)
    doc.add_page_break()
    ch_results(doc)
    doc.add_page_break()
    ch_qa(doc)
    doc.add_page_break()
    ch_closing(doc)
    ch_tips(doc)
    doc.save(REPORT_PATH)
    print(f"答辩讲稿已生成：{REPORT_PATH}")


if __name__ == "__main__":
    main()
