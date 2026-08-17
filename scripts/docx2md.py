#!/usr/bin/env python3
"""docx2md — 将 .docx 论文转换为结构化 markdown (格式转换层: 输入阶段).

Markdown 约定 (与 md2docx.py 互逆):
  # 题目        → 论文题目 (黑体四号居中, 重建时套用)
  (无标记段落)   → 作者行 / 副标题 / 摘要 / 关键词 (原样保留)
  ## 一、xxx    → 一级标题
  ### （一）xxx → 二级标题
  #### 1. xxx   → 三级标题
  (无标记段落)   → 正文
  [^n]          → 脚注引用标记 (追加于所在段落末尾)
  [^n]: 文本    → 脚注定义 (统一置于文末, 按编号升序)

限制 (诚实声明):
  - 仅提取内容与结构 (段落/标题层级/脚注), 不保留原有排版格式;
    重建 .docx 时按规则表统一套用排版 (见 md2docx.py).
  - 不支持表格 / 图片 / 批注 / 域; 这些内容在转换中丢失.
  - 脚注引用按编号升序输出定义; 正文中脚注引用标记保留在段末.

用法:
  python docx2md.py 论文.docx -o 论文.md
  python docx2md.py 论文.docx            # 输出到同名 .md
"""

import argparse
import sys
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.oxml.ns import qn

from format_paper import (detect_heading_level, detect_title_range,
                          is_abstract_or_keywords)
from citation_formatter import extract_footnotes

# 标题层级 → markdown 前缀
_HEADING_PREFIX = {1: '## ', 2: '### ', 3: '#### ', 4: '##### '}


def _iter_paragraph_content(para) -> Tuple[str, List[int]]:
    """按阅读顺序提取段落纯文本与脚注引用 id 列表.

    遍历段落 XML 的所有后代元素: w:t 收集文本, w:footnoteReference
    收集脚注 id (id > 0 为真实脚注, 0/-1 为分隔符).

    Returns:
        (text, footnote_ids) — 脚注 id 按出现顺序.
    """
    texts = []
    refs = []
    for elem in para._p.iter():
        if elem.tag == qn('w:footnoteReference'):
            fn_id = int(elem.get(qn('w:id'), '0'))
            if fn_id > 0:
                refs.append(fn_id)
        elif elem.tag == qn('w:t') and elem.text:
            texts.append(elem.text)
    return ''.join(texts).strip(), refs


def docx_to_markdown(input_path: str) -> str:
    """将 .docx 转换为结构化 markdown 文本.

    Args:
        input_path: 输入 .docx 文件路径.

    Returns:
        markdown 文本 (UTF-8).
    """
    doc = Document(input_path)
    paragraphs = doc.paragraphs

    # —— 检测题目区域 (复用 format_paper 的规则) ——
    title_start, title_end = detect_title_range(paragraphs)
    title_indices = set(range(title_start, title_end)) if title_start >= 0 else set()

    lines: List[str] = []
    pending_fns: List[int] = []  # 当前段落待输出的脚注引用

    for i, para in enumerate(paragraphs):
        text, refs = _iter_paragraph_content(para)
        if not text:
            continue

        if i in title_indices:
            if i == title_start:
                # 论文题目 → # 一级标题
                lines.append(f'# {text}')
            else:
                # 作者行/副标题/摘要/关键词 → 原样保留 (重建时按规则识别)
                lines.append(text)
        else:
            level = detect_heading_level(text)
            if level > 0:
                lines.append(f'{_HEADING_PREFIX[level]}{text}')
            else:
                lines.append(text)

        # 脚注引用标记: 追加于段落末尾
        for fn_id in refs:
            if fn_id not in pending_fns:
                pending_fns.append(fn_id)
        if refs:
            ref_markers = ''.join(f'[^{n}]' for n in refs)
            lines[-1] = lines[-1] + ref_markers

        lines.append('')  # 段落间空行分隔

    # —— 脚注定义: 统一置于文末, 按 id 升序 ——
    footnotes = {fn['id']: fn['full_text'] for fn in extract_footnotes(doc)}
    if footnotes:
        lines.append('')  # 与正文分隔
        for fn_id in sorted(footnotes):
            # 多段脚注以空格合并为单行
            text = ' '.join(footnotes[fn_id].split('\n'))
            # 剥离脚注编号残留 "[]" (编号由 Word 自动生成, md 中由 [^n] 表示)
            if text.startswith('[]'):
                text = text[2:]
            lines.append(f'[^{fn_id}]: {text}')

    return '\n'.join(lines).rstrip() + '\n'


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        description='docx → markdown 转换 (输入阶段: 先转 md, 再修改/重建)')
    ap.add_argument('input', help='输入 .docx 文件')
    ap.add_argument('-o', '--output', default=None,
                    help='输出 .md (默认与输入同名)')
    return ap


def main():
    args = build_parser().parse_args()
    inp = Path(args.input)
    if not inp.exists():
        print(f"错误: 输入文件不存在: {inp}", file=sys.stderr)
        sys.exit(1)
    if inp.suffix.lower() != '.docx':
        print(f"错误: 不支持的文件格式: {inp.suffix} (仅支持 .docx)", file=sys.stderr)
        sys.exit(1)

    output = args.output or str(inp.with_suffix('.md'))
    md_text = docx_to_markdown(str(inp))
    Path(output).write_text(md_text, encoding='utf-8')
    print(f"已转换: {inp} → {output}")
    fn_count = md_text.count('\n[^')
    print(f"脚注: {fn_count} 条" if fn_count else "脚注: 无")


if __name__ == '__main__':
    main()
