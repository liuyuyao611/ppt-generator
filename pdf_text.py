"""Local PDF decoding checks and prose selection; never invent replacement text."""
from __future__ import annotations

import re
import shutil
import subprocess
import unicodedata
from collections import Counter


class PdfTextError(ValueError):
    pass


def text_problem(text):
    """Catch missing font maps and mixed-script glyph IDs, not ordinary maths.

    A single non-Latin language is valid. Random CID values interpreted as
    Unicode tend to scatter across many unrelated writing systems instead.
    """
    chars = [c for c in text if not c.isspace()]
    if not chars:
        return ""
    broken = sum(c == "\ufffd" or unicodedata.category(c) in {"Co", "Cs", "Cn"}
                 for c in chars)
    if broken >= max(2, len(chars) * .02) or re.search(r"\(cid:\d+\)", text):
        return "缺失字符映射"
    scripts = Counter()
    for c in chars:
        if not c.isalpha():
            continue
        name = unicodedata.name(c, "")
        script = name.split(" ")[0] if name else "UNKNOWN"
        if script not in {"LATIN", "GREEK", "MATHEMATICAL"}:
            # All CJK unified/compatibility ideographs share a script family.
            scripts["CJK" if "IDEOGRAPH" in name else script] += 1
    total = sum(scripts.values())
    if total >= 20 and sum(v >= 2 for v in scripts.values()) >= 5:
        if max(scripts.values()) / total < .7:
            return "字体编码混乱"
    return ""


def poppler_pages(data, count):
    executable = shutil.which("pdftotext")
    if not executable:
        raise PdfTextError("备用 PDF 解析器暂不可用")
    try:
        # stdin/stdout keep uploaded documents in memory; no shell or file paths.
        result = subprocess.run(
            [executable, "-enc", "UTF-8", "-eol", "unix", "-", "-"],
            input=data, capture_output=True, timeout=30, check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise PdfTextError("备用 PDF 解析未完成") from exc
    if result.returncode or len(result.stdout) > 4_000_000:
        raise PdfTextError("备用 PDF 解析未完成")
    try:
        pages = result.stdout.decode("utf-8", errors="strict").split("\f")
    except UnicodeError as exc:
        raise PdfTextError("备用 PDF 编码异常") from exc
    if len(pages) == count + 1 and not pages[-1].strip():
        pages.pop()
    if len(pages) != count:
        raise PdfTextError("备用解析的 PDF 页码不完整")
    return pages


def extract_pages(reader, data):
    texts, diagnostics = [], []
    for i, page in enumerate(reader.pages, 1):
        try:
            text = page.extract_text() or ""
            problem = text_problem(text)
        except Exception:
            text, problem = "", "文字层读取失败"
        texts.append(text)
        diagnostics.append({"page": i, "engine": "pypdf", "problem": problem})
    needs_fallback = [i for i, d in enumerate(diagnostics)
                      if d["problem"] or len(texts[i].strip()) < 12]
    if needs_fallback:
        try:
            alternate = poppler_pages(data, len(texts))
        except PdfTextError:
            alternate = None
        if alternate is not None:
            for i in needs_fallback:
                text = alternate[i]
                if not text_problem(text) and len(text.strip()) >= 12:
                    texts[i] = text
                    diagnostics[i].update(engine="Poppler", problem="")
    failed = [str(d["page"]) for d in diagnostics if d["problem"]]
    if failed:
        raise PdfTextError(
            f"PDF 第 {', '.join(failed[:15])} 页字体映射异常，无法可靠读取正文。"
            "请先对 PDF 做 OCR（文字识别），或上传 Word / 粘贴正文后重试。"
            "已停止生成，避免把乱码写进 PPT。")
    return texts, diagnostics


_LIGATURES = str.maketrans({"ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi",
                          "ﬄ": "ffl", "ﬅ": "st", "ﬆ": "st", "\u00ad": ""})


def is_prose(text):
    """Conservative display candidates; equations remain in sources and notes."""
    if text_problem(text):
        return False
    cjk = len(re.findall(r"[\u3400-\u9fff]", text))
    words = re.findall(r"[A-Za-z]{2,}", text)
    return cjk >= 5 or len(words) >= 4


def prose_paragraphs(text):
    """Join visual line wraps but keep equations and paragraph boundaries apart."""
    paragraphs, pending = [], ""

    def flush():
        nonlocal pending
        if pending and is_prose(pending):
            paragraphs.append(pending)
        pending = ""

    for raw in text.translate(_LIGATURES).replace("\r\n", "\n").splitlines():
        line = raw.strip()
        if not line:
            flush()
            continue
        # A short heading is a boundary, not a complete presentation bullet.
        heading = len(line) < 90 and re.match(r"^(?:\d+(?:\.\d+){0,3}|[IVX]+\.|[A-Z]\.)\s+[^\d]", line)
        formula = not is_prose(line) and (
            re.search(r"[=+∑∫⊙≤≥]|^\(?\d+\)?$", line)
            or len(re.findall(r"[A-Za-z\u3400-\u9fff]", line)) < 2)
        if heading or formula:
            flush()
            continue
        if pending:
            if re.search(r"[A-Za-z]-$", pending) and re.match(r"[a-z]", line):
                pending = pending[:-1] + line
            else:
                adjacent_cjk = bool(re.search(r"[\u3400-\u9fff]$", pending)
                                    or re.match(r"[\u3400-\u9fff]", line))
                pending += ("" if adjacent_cjk else " ") + line
        else:
            pending = line
        if re.search(r"[。！？.!?;；]$", line):
            flush()
    flush()
    return paragraphs


def running_lines(pages):
    """Identify repeated margin text for display only; retain it in raw sources."""
    counts = Counter()
    for text in pages:
        lines = [s.strip() for s in text.splitlines() if s.strip()]
        counts.update(set(margin_key(s) for s in lines[:3] + lines[-2:]))
    return {s for s, count in counts.items() if s and count >= 3}


def margin_key(line):
    return re.sub(r"[\d\s]", "", line)
