"""Conservative source extraction and text-to-diagram eligibility."""
import re


_OUTCOME = re.compile(r"结果|发现|表明|结论|达到|下降|上升|相比|降低|提高|差异|局限|尚未|仅在|observed|result|limitation", re.I)
_ACTION = re.compile(r"目的|目标|采用|使用|测量|采集|验证|分析|设置|校准|对照|比较|任务|步骤|method|objective", re.I)
_FILLER = re.compile(r"^(目录|参考文献|致谢|摘要\s*$|关键词\s*[:：]|第\s*\d+\s*页|page\s+\d+)", re.I)
_FIRST = re.compile(r"^(?:首先|第一步|第一阶段|先(?=将|在|用|对|测|记录|固定|进行|校准))")
_NEXT = re.compile(r"^(?:随后|接着|然后|其次|再|最后|第二步|第三步|第四步|最终)")
_COMPARE = re.compile(r"相比|相较|高于|低于|优于|\bvs\.?\b", re.I)
_GROUP = re.compile(r"^\s*(?:方案[一二AB１２12]|[AB](?:组|方案|[：: ])|实验组|对照组|基线组)", re.I)
_TASK = re.compile(r"任务|工作|调试|整理|阅读|采集|测量|分析|复核|核查|设计|记录|校准|验证|部署|访谈|测试")
_MEASUREMENT = re.compile(r"^([^：:\n]{1,14})[：:]\s*([^\d：:\n]{1,18}?)\s*([0-9]{1,9}(?:\.[0-9]{1,6})?)"
                          r"\s*(%|mm|cm|ms|Hz|kg|g|s|μm|µm|例|次|人)[。.]?$")


def rank_excerpts(units, limit=4):
    """Select intact source excerpts, retaining their original order and source IDs."""
    candidates, seen = [], set()
    for index, (block, sentence) in enumerate(units):
        key = re.sub(r"\W", "", sentence).casefold()
        if not key or key in seen:
            continue
        seen.add(key)
        score = (4 * bool(_OUTCOME.search(sentence)) + 3 * bool(_ACTION.search(sentence))
                 + 2 * bool(re.search(r"\d+(?:\.\d+)?\s*(?:%|％|例|次|[a-zA-Zµμ]+\b)", sentence))
                 + int(14 <= len(sentence) <= 105) - 8 * bool(_FILLER.match(sentence)))
        candidates.append((score, index, block["id"], sentence))
    best = sorted(candidates, key=lambda item: (-item[0], item[1]))[:limit]
    return [(source, sentence) for _, _, source, sentence in sorted(best, key=lambda item: item[1])]


def diagram_for(page):
    """Return a diagram backed by the current text or an explicit author-selected relation."""
    bullets = diagram_points(page)
    if not 2 <= len(bullets) <= 6:
        return None
    relation = page.get("relationship", "parallel")
    if relation == "sequence" or (_FIRST.match(bullets[0]) and all(_NEXT.match(x) for x in bullets[1:])):
        return "flow"
    if len(bullets) == 2 and (relation == "comparison" or
            (all(_GROUP.match(x) for x in bullets) and any(_COMPARE.search(x) for x in bullets))):
        return "bars" if comparable_values(bullets) else "comparison"
    if relation == "parallel" and all(_TASK.search(x) for x in bullets):
        return "tasks"
    return None


def diagram_points(page):
    bullets = page.get("bullets", [])
    if len(bullets) == 1:
        parts = [part.strip() for part in re.split(r"[，,；;]", bullets[0])]
        if 2 <= len(parts) <= 6 and _FIRST.match(parts[0]) and all(_NEXT.match(x) for x in parts[1:]):
            return parts
    return bullets


def comparable_values(bullets):
    """Use bars only for a confirmed comparison with two explicitly matching measurements."""
    if len(bullets) != 2:
        return None
    matches = [_MEASUREMENT.fullmatch(line.strip()) for line in bullets]
    if not all(matches):
        return None
    pair = [tuple(part.strip() for part in match.groups()) for match in matches]
    a, b = pair
    if (a[1], a[3]) != (b[1], b[3]) or a[0] == b[0] or max(float(a[2]), float(b[2])) <= 0:
        return None
    return pair


DIAGRAM_LABELS = {"flow": "步骤流程图", "tasks": "并列任务图", "comparison": "左右对照图",
                  "bars": "同指标数值对比条"}
