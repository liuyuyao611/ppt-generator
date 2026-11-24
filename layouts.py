"""Content relationships choose layouts; visual variety never invents a relationship."""
import re

LAYOUT_LABELS = {
    "auto": "自动推荐",
    "focus": "重点聚焦",
    "cards": "并列主题卡片",
    "comparison": "左右对照",
    "process": "步骤流程",
    "image": "图文并排",
    "literature": "文献概览",
    "takeaway": "总结收束",
    "columns": "双栏要点",
    "bullets": "简洁列表",
}

RELATION_LABELS = {"parallel": "并列信息", "sequence": "步骤顺序",
                   "comparison": "比较关系", "statement": "重点说明"}


def recommend_layout(slide, images=None, previous=None):
    images = images or {}
    title = slide.get("title", "")
    bullets = slide.get("bullets", [])
    count = len(bullets)
    longest = max(map(len, bullets), default=0)
    relation = slide.get("relationship", "parallel")
    if re.search(r"总结|结论|回顾|启示|小结|展望", title):
        return "takeaway", "将主要结论与补充说明分层呈现"
    if slide.get("image") in images:
        if re.search(r"文献|阅读|综述|研究概览", title):
            return "literature", "并列阅读线索配原文图片，来源单独列出"
        return "image", "保留原图比例，文字与图片分区"
    if count in {2, 4} and (relation == "comparison" or re.search(r"对比|比较|差异|对照", title)):
        return "comparison", "依据已标明的比较关系并排呈现"
    if 2 <= count <= 6 and (relation == "sequence" or re.search(r"流程|步骤|实验顺序|技术路线", title)):
        return "process", "依据原文的步骤关系展示阅读顺序"
    if re.search(r"文献|阅读|综述|研究概览", title) and count <= 4:
        return "literature", "左侧独立阅读条目，右侧原文来源索引"
    if count == 1 or (count <= 3 and longest <= 180 and re.search(r"背景|问题|目标|目的|挑战", title)):
        return "focus", "突出一个焦点，其余内容作为支撑"
    if 2 <= count <= 4 and longest <= 180:
        if previous == "cards":
            return "columns", "并列内容采用双栏，避免连续重复卡片版式"
        return "cards", "多个独立要点分别呈现，不添加因果或先后关系"
    return "columns" if count >= 4 else "bullets", "按文字长度与数量安排阅读密度"


def resolve_layout(slide, images=None, previous=None):
    choice = slide.get("layout", "auto")
    if choice in LAYOUT_LABELS and choice != "auto":
        return choice, "手动选择；请核对版式表达的关系是否符合原文"
    return recommend_layout(slide, images, previous)
