#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bookmark_pdf.py —— 给 Mac 版 Word 导出的 PDF 补上 /Outlines（书签大纲）

思路：docx 里有标题层级和标题文字，PDF 里有文字和位置。
     从 docx 按文档顺序读出 (级别, 标题)，再到 PDF 里按顺序定位到页/行，
     最后用 PyMuPDF 的 set_toc() 写入真正的 PDF 大纲树。

用法：
    python3 bookmark_pdf.py 文档.docx 文档.pdf              # 预览（不写文件）
    python3 bookmark_pdf.py 文档.docx 文档.pdf -o 输出.pdf  # 生成带书签的 PDF
    python3 bookmark_pdf.py 文档.docx 文档.pdf -o 输出.pdf --report 报告.txt

依赖：pip install pymupdf python-docx
"""
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------- 读取 docx

HEADING_NAME_RE = re.compile(r"^(?:heading|标题)\s*([1-9])$", re.IGNORECASE)


def _outline_level(p_el, style_el=None):
    """从 w:pPr/w:outlineLvl 或样式上的 outlineLvl 取大纲级别（0 基）。"""
    for el in (p_el, style_el):
        if el is None:
            continue
        try:
            pPr = el.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
        except Exception:
            pPr = None
        if pPr is None:
            continue
        ol = pPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl")
        if ol is not None:
            try:
                return int(ol.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))
            except (TypeError, ValueError):
                pass
    return None


def read_headings(docx_path: str):
    """返回 [(level(1起), text, 段落序号)]，按文档顺序。"""
    from docx import Document

    doc = Document(docx_path)
    out = []
    for idx, p in enumerate(doc.paragraphs):
        text = (p.text or "").strip()
        if not text:
            continue

        level = None
        m = HEADING_NAME_RE.match((p.style.name or "").strip())
        if m:
            level = int(m.group(1))

        if level is None:
            ol = _outline_level(p._p, p.style.element)
            if ol is not None and 0 <= ol <= 8:
                level = ol + 1

        if level is not None:
            out.append((level, text, idx))

    return out


# ---------------------------------------------------------------- 文本归一化

_PUNCT = dict.fromkeys(map(ord, "（）【】『』「」、，。；：？！·…—～\"'"))


def norm(s: str) -> str:
    """NFKC 统一全/半角，去掉所有空白和标点，便于宽松比对。"""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(_PUNCT)
    s = re.sub(r"\s+", "", s)
    return s.lower()


def is_toc_line(text: str) -> bool:
    """目录条目：带点线/省略号引导符且以页码结尾。"""
    t = text.strip()
    if not t:
        return False
    if re.search(r"[.·…]{2,}\s*\d+$", t) or re.search(r"\t\s*\d+$", t):
        return True
    if re.search(r"(?:\s{2,}|\.{3,})\s*\d+$", t):
        return True
    return False


# ---------------------------------------------------------------- PDF 行索引

@dataclass
class Line:
    page: int          # 1 基页码
    y: float
    x: float
    text: str
    size: float
    bold: bool
    nxt: "Line | None" = None      # 同一段落里紧跟的下一行（用于还原折行的标题）

    @property
    def n(self) -> str:
        return norm(self.text)

    @property
    def n_all(self) -> str:
        return self.n + (self.nxt.n if self.nxt else "")

    @property
    def full(self) -> str:
        return self.text + ((" " + self.nxt.text) if self.nxt else "")


@dataclass
class Match:
    level: int
    title: str
    page: int | None = None
    line: str = ""
    score: int = 0
    note: str = ""


def build_lines(doc, margin_ignore: float):
    """把 PDF 每一行摊平成带位置信息的列表；页眉页脚按边距带丢弃。"""
    lines: list[Line] = []
    for pno in range(doc.page_count):
        page = doc[pno]
        rect = page.rect
        top = rect.y0 + rect.height * margin_ignore
        bottom = rect.y1 - rect.height * margin_ignore
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            for ln in block.get("lines", []):
                spans = ln.get("spans") or []
                if not spans:
                    continue
                text = "".join(s.get("text", "") for s in spans).strip()
                if not text:
                    continue
                bbox = ln.get("bbox") or (0, 0, 0, 0)
                y = (bbox[1] + bbox[3]) / 2
                if y < top or y > bottom:      # 页眉 / 页脚区域
                    continue
                size = max((s.get("size") or 0) for s in spans)
                fonts = " ".join((s.get("font") or "") for s in spans)
                bold = bool(re.search(r"bold|black|heavy|semibold|demi", fonts, re.I))
                lines.append(Line(pno + 1, y, bbox[0], text, size, bold))
    lines.sort(key=lambda l: (l.page, l.y, l.x))

    # 把视觉上连续的两行连起来，标题在 PDF 里折行时也能整体匹配
    for i, ln in enumerate(lines[:-1]):
        nxt = lines[i + 1]
        if nxt.page != ln.page:
            continue
        gap = nxt.y - ln.y
        if abs(nxt.size - ln.size) > 0.6:      # 字号变了说明不是同一个块
            continue
        if 0 < gap < ln.size * 1.95 and abs(nxt.x - ln.x) < ln.size * 2.0:
            ln.nxt = nxt
    return lines


def flag_toc_pages(lines, page_count: int, headings=(), manual=None) -> set[int]:
    """识别文档自带的目录页，匹配时整页跳过。

    两路信号：
      A. 点线/制表符 + 页码结尾的行 ≥ 2 条
      B. 兜底：某一页同时“命中”≥3 个不同标题（目录条目常不带点线）
         为免误伤正文里紧挨着的多个小标题，只在文档前 40% 生效。
    """
    if manual:
        return set(manual)

    hits: dict[int, int] = {}
    for ln in lines:
        if is_toc_line(ln.full):
            hits[ln.page] = hits.get(ln.page, 0) + 1
    toc = {p for p, c in hits.items() if c >= 2}

    normed = [n for n in (norm(t) for _l, t, _i in headings) if n]
    if normed:
        per: dict[int, set[str]] = {}
        for ln in lines:
            nl = ln.n
            for n in normed:
                if n in nl and len(nl) - len(n) <= 20:
                    per.setdefault(ln.page, set()).add(n)
        front = max(2, int(page_count * 0.4))
        for p, ks in per.items():
            if len(ks) >= 3 and p <= front:
                toc.add(p)
    return toc


# ---------------------------------------------------------------- 匹配

LOOKAHEAD_PAGES = 1     # 找到首个候选后，再多看几页，取分数更高的那个


def match_headings(headings, lines, toc_pages):
    matches: list[Match] = []
    cursor = 0                      # 单调游标：页序不倒流
    body_sizes = sorted(l.size for l in lines) or [10.0]
    median_size = body_sizes[len(body_sizes) // 2]

    for level, title, _idx in headings:
        m = Match(level, title)
        nt = norm(title)
        if not nt:
            matches.append(m)
            continue

        first_page = None
        cands = []                  # [(score, index, line)]
        for i in range(cursor, len(lines)):
            ln = lines[i]
            if ln.page in toc_pages:
                continue
            if first_page is not None and ln.page > first_page + LOOKAHEAD_PAGES:
                break

            for nl, joined in ((ln.n, False), (ln.n_all, True)):
                use = ln
                if nt not in nl:
                    continue
                # 标题应当“独占一行”：行比标题长太多，多半是正文里顺带提到的
                if len(nl) - len(nt) > 20:
                    continue
                score = 40
                if nl == nt:
                    score = 100
                elif nl.endswith(nt):
                    score = 85        # 行首有自动编号（1.1 / 第三章 / 一、）
                elif nl.startswith(nt):
                    score = 70
                if use.size >= median_size * 1.08:
                    score += 10
                if use.bold:
                    score += 5
                if first_page is None:
                    first_page = use.page
                cands.append((score, i, use, joined))
                break

        if cands:
            cands.sort(key=lambda c: (-c[0], c[1]))
            score, i, ln, joined = cands[0]
            m.page = ln.page
            m.line = ln.full if joined else ln.text
            m.score = score
            cursor = i
            if score < 60:
                m.note = "低位匹配（可能误配，请核对）"
            elif len(cands) > 1 and cands[1][0] == score and cands[1][2].page != ln.page:
                m.note = "同名候选跨页，已取靠前者"
        else:
            m.page = matches[-1].page if matches else 1
            m.note = "未找到，已沿用上一页"
        matches.append(m)
    return matches


# ---------------------------------------------------------------- 写入大纲

def build_toc(matches):
    """把 [(level, title, page)] 转成 PyMuPDF 的嵌套 toc 结构。"""
    toc = []
    stack = []                      # [(level, index_in_toc)]
    for m in matches:
        lvl = max(1, m.level)
        # 修正跳级：h1 -> h3 时把 h3 提到 h2，避免阅读器里层级断裂
        while stack and stack[-1][0] >= lvl:
            stack.pop()
        actual = (stack[-1][0] + 1) if stack else 1
        entry = [actual, m.title, m.page]
        toc.append(entry)
        stack.append((actual, len(toc) - 1))
    return toc


def apply_toc(src: str, dst: str, toc):
    import pymupdf

    doc = pymupdf.open(src)
    doc.set_toc(toc)
    doc.save(dst, garbage=3, deflate=True)
    doc.close()



def dw(s: str) -> int:
    """显示宽度：中日韩全角按 2 计。"""
    w = 0
    for ch in s:
        w += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return w


def pad(s: str, width: int) -> str:
    s = s if dw(s) <= width else s[: max(1, width // 2)] + "…"
    return s + " " * max(0, width - dw(s))


# ---------------------------------------------------------------- 命令行

def main(argv=None):
    ap = argparse.ArgumentParser(description="给 PDF 补书签（大纲），依据 docx 的标题样式")
    ap.add_argument("docx", help="源 Word 文档（用来取标题层级）")
    ap.add_argument("pdf", help="Mac 版 Word 导出的 PDF")
    ap.add_argument("-o", "--out", help="输出 PDF；不给则只预览")
    ap.add_argument("--margin-ignore", type=float, default=0.07,
                    help="忽略页面上下各 x 比例区域（页眉页脚），默认 0.07")
    ap.add_argument("--report", help="把匹配明细写入文本文件")
    ap.add_argument("--max-level", type=int, default=9, help="最多收录到第几级标题")
    ap.add_argument("--toc-pages", help="手工指定要跳过的目录页，如 2,3（覆盖自动识别）")
    args = ap.parse_args(argv)

    if not Path(args.docx).exists():
        sys.exit(f"找不到 docx：{args.docx}")
    if not Path(args.pdf).exists():
        sys.exit(f"找不到 pdf：{args.pdf}")

    headings = [h for h in read_headings(args.docx) if h[0] <= args.max_level]
    if not headings:
        sys.exit("docx 里没有检测到任何标题（需要 Word 内置的“标题 1/2/3”样式或大纲级别）。")

    import pymupdf

    doc = pymupdf.open(args.pdf)
    page_count = doc.page_count
    lines = build_lines(doc, args.margin_ignore)
    manual = None
    if args.toc_pages:
        manual = {int(x) for x in re.split(r"[,\s]+", args.toc_pages.strip()) if x}
    toc_pages = flag_toc_pages(lines, page_count, headings, manual)
    matches = match_headings(headings, lines, toc_pages)
    doc.close()

    toc = build_toc(matches)

    # ---- 预览表
    print(f"共 {len(matches)} 个标题，PDF 共 {page_count} 页；"
          f"跳过的目录页：{sorted(toc_pages) or '无'}")
    print("-" * 78)
    print(f"{'页':>4}  {'级':>2}  {pad('标题', 36)}匹配到的 PDF 行")
    print("-" * 78)
    bad = []
    for m in matches:
        print(f"{m.page:>4}  H{m.level:<2}  {pad(m.title, 36)}{m.line[:46]}")
        if m.note:
            bad.append(m)
    print("-" * 78)

    if bad:
        print(f"\n⚠️  {len(bad)} 条需要人工确认：")
        for m in bad:
            print(f"   · 第{m.page}页  H{m.level}  {m.title}  ←  {m.note}")

    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            f.write(f"{'page':>5}\t{'level'}\t{'score':>3}\ttitle\tmatched_line\n")
            for m in matches:
                f.write(f"{m.page:>5}\tH{m.level}\t{m.score:>3}\t{m.title}\t{m.line}\n")
        print(f"\n明细已写入 {args.report}")

    if not args.out:
        print("\n（预览模式，未写出文件。加 -o 输出.pdf 生成。）")
        return 0

    apply_toc(args.pdf, args.out, toc)

    # ---- 回读校验
    chk = pymupdf.open(args.out)
    got = chk.get_toc(simple=True)
    chk.close()
    print(f"\n✅ 已写出 {args.out}\n   回读校验：{len(got)} 条大纲")
    for lvl, title, page in got[:8]:
        print(f"   {'  ' * (lvl - 1)}H{lvl} {title} → p{page}")
    if len(got) > 8:
        print(f"   …… 其余 {len(got) - 8} 条略")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
