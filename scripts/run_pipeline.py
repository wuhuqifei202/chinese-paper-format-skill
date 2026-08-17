#!/usr/bin/env python3
"""Single entry-point: runs the paper formatting pipeline.

Usage:
    # 直接格式化 .docx (原地排版)
    python scripts/run_pipeline.py --input 论文.docx --output 论文_格式化.docx

    # markdown 中转工作流: docx → md (可审查/修改) → 重建 docx (双交付)
    python scripts/run_pipeline.py --input 论文.docx --via-markdown

    # 旧版 .doc 自动转换 (需要 LibreOffice)
    python scripts/run_pipeline.py --input 论文.doc --via-markdown
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# 确保脚本目录在 sys.path 中 — 裸 import format_paper / citation_rules / rules 依赖于此
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from format_paper import format_document, check_document, backup_file

# [^n]: 脚注定义行 (markdown 层引注修复用)
_FN_DEF_RE = re.compile(r'^(\[\^(\d+)\]:\s*)(.*)$')

# LibreOffice 可执行文件候选 (Windows 常见安装路径 + PATH)
_SOFFICE_CANDIDATES = [
    'soffice', 'soffice.exe', 'libreoffice', 'libreoffice.exe',
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
]


def find_soffice() -> str:
    """查找 LibreOffice soffice 可执行文件, 找不到返回空串."""
    for cand in _SOFFICE_CANDIDATES:
        path = shutil.which(cand)
        if path:
            return path
        if Path(cand).exists():
            return cand
    return ''


def convert_doc_to_docx(input_path: str) -> str:
    """用 LibreOffice 将 .doc 转换为 .docx, 返回转换后的路径.

    转换输出到临时目录 (避免覆盖同名文件), 调用方负责清理.

    Raises:
        RuntimeError: soffice 不可用或转换失败.
    """
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            '未检测到 LibreOffice (soffice)。请安装 LibreOffice 以支持 .doc 格式,'
            '或将文件另存为 .docx 后重试。下载: https://www.libreoffice.org/')

    inp = Path(input_path)
    with tempfile.TemporaryDirectory(prefix='doc2docx_') as tmp:
        result = subprocess.run(
            [soffice, '--headless', '--convert-to', 'docx',
             '--outdir', tmp, str(inp)],
            capture_output=True, text=True, timeout=120)
        out = Path(tmp) / (inp.stem + '.docx')
        if result.returncode != 0 or not out.exists():
            raise RuntimeError(
                f'LibreOffice 转换失败: {result.stderr.strip() or "未知错误"}')
        # 复制到输入文件同目录 (临时目录随上下文清理)
        target = inp.with_suffix('.docx')
        shutil.copy2(str(out), str(target))
        return str(target)


def run_pipeline(input_path: str, output_path: str,
                 body_indent: int = 2, check_first: bool = True,
                 rules: dict = None) -> dict:
    """Run the formatting pipeline on a single document.

    Args:
        input_path: Input .docx file.
        output_path: Output .docx file.
        body_indent: First-line indent for body text in characters.
        check_first: If True, run --check before formatting.
        rules: 自定义格式规则表 (rules.py load_rules 输出); None = 默认.

    Returns:
        Statistics dict from format_document.
    """
    inp = Path(input_path)

    # Step 1: Validate
    if not inp.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if inp.suffix.lower() not in ('.docx',):
        raise ValueError(f"Unsupported format: {inp.suffix}. Use .docx files.")

    # Step 2: Check structure (optional)
    if check_first:
        result = check_document(str(inp))
        print(f"题目: {result['title'][:80] if result['title'] else '(未检测到)'}")
        print(f"标题: {len(result['headings'])} 个")
        for level, txt in result['headings']:
            prefix = {1: '一、', 2: '（一）', 3: '1. '}.get(level, '?')
            print(f"  L{level} [{prefix}] {txt}")

    # Step 3: Backup (if overwriting)
    if str(inp.resolve()) == str(Path(output_path).resolve()):
        bkp = backup_file(str(inp))
        print(f"已备份: {Path(bkp).name}")

    # Step 4: Format
    stats = format_document(str(inp), str(output_path),
                            body_indent=body_indent, rules=rules)
    print(f"输出: {output_path}")
    return stats


def fix_citations_in_markdown(md_text: str, style=None) -> tuple:
    """在 markdown 层修复脚注引注格式 ([^n]: 行文本).

    Args:
        style: 注释体例 dict; None = 默认体例.

    Returns:
        (md_text, fixed_count)
    """
    from citation_formatter import auto_fix_footnote
    lines = md_text.split('\n')
    fixed = 0
    for i, line in enumerate(lines):
        m = _FN_DEF_RE.match(line)
        if m:
            new_text, count = auto_fix_footnote(m.group(3), style=style)
            if count > 0:
                lines[i] = f'{m.group(1)}{new_text}'
                fixed += count
    return '\n'.join(lines), fixed


def run_pipeline_via_markdown(input_path: str, output_path: str,
                              body_indent: int = 2,
                              fix_citations: bool = False,
                              rules: dict = None,
                              citation_rules: dict = None) -> dict:
    """markdown 中转工作流: docx → md → (修复) → 重建 docx.

    输入支持 .doc (自动用 LibreOffice 转换) 与 .docx。
    双交付: 输出修改后的 .md 与重建的 .docx (md 与输入同名, 位于输出目录)。

    Args:
        rules: 自定义格式规则表; None = 默认规则.
        citation_rules: 注释体例 dict; None = 默认体例.

    Returns:
        统计信息 dict (md2docx 的 stats + 引注修复统计).
    """
    inp = Path(input_path)
    out = Path(output_path)
    if not inp.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Step 1: .doc → .docx (LibreOffice)
    docx_path = str(inp)
    if inp.suffix.lower() == '.doc':
        docx_path = convert_doc_to_docx(str(inp))
        print(f"已转换 (.doc → .docx): {docx_path}")

    # Step 2: docx → md
    from docx2md import docx_to_markdown
    md_text = docx_to_markdown(docx_path)
    md_path = out.with_suffix('.md')
    md_path.parent.mkdir(parents=True, exist_ok=True)
    Path(md_path).write_text(md_text, encoding='utf-8')
    print(f"已转换 (docx → md): {md_path}")

    # Step 3: md 层引注修复 (可选)
    if fix_citations:
        md_text, fixed = fix_citations_in_markdown(md_text, style=citation_rules)
        if fixed:
            Path(md_path).write_text(md_text, encoding='utf-8')
        print(f"引注修复: {fixed} 处")

    # Step 4: md → 重建 docx (排版按规则表统一套用)
    from md2docx import markdown_to_docx
    stats = markdown_to_docx(md_text, str(out), body_indent=body_indent,
                             rules=rules)
    print(f"已重建 (md → docx): {output_path}")
    stats['md_path'] = str(md_path)
    return stats


def main():
    ap = argparse.ArgumentParser(
        description='中文学术论文格式化 — 一键流水线')
    ap.add_argument('--input', required=True, help='输入 .docx 文件')
    ap.add_argument('--output', default=None, help='输出 .docx (默认覆盖输入)')
    ap.add_argument('--body-indent', type=int, default=2,
                    help='正文首行缩进字符数 (默认 2)')
    ap.add_argument('--no-check', action='store_true',
                    help='跳过格式检查步骤')
    ap.add_argument('--fix-citations', action='store_true',
                    help='自动修复脚注引注格式')
    ap.add_argument('--via-markdown', action='store_true',
                    help='markdown 中转工作流: docx→md→重建docx (md与docx双交付; '
                         '支持 .doc, 自动用 LibreOffice 转换)')
    ap.add_argument('--rules', default=None,
                    help='自定义格式规则 JSON 配置文件 '
                         '(如 {"title": {"size": "二号", "bold": true}})')
    ap.add_argument('--citation-rules', default=None,
                    help='自定义注释体例 JSON 配置文件 '
                         '(见 references/citation-rules.md)')
    ap.add_argument('--preset', default=None,
                    help='预设注释体例: 法学引注手册 / 《法学家》《中外法学》注释体例 / 《中国法学》 / 《法商研究》 / 《法学研究》')
    args = ap.parse_args()

    # 加载自定义格式规则
    rules = None
    if args.rules:
        from rules import load_rules
        try:
            rules = load_rules(args.rules)
            print(f"已加载自定义格式规则: {args.rules}")
        except (ImportError, ValueError, FileNotFoundError) as e:
            print(f"错误: {e}", file=sys.stderr)
            sys.exit(1)

    # 加载自定义注释体例 (预设或 JSON)
    citation_rules = None
    if args.citation_rules or args.preset:
        try:
            from citation_rules import load_citation_rules
            citation_rules = load_citation_rules(args.citation_rules,
                                                 preset=args.preset)
            src = args.citation_rules or f'预设 {args.preset}'
            print(f"已加载注释体例: {src}")
        except (ImportError, ValueError, FileNotFoundError) as e:
            print(f"错误: {e}", file=sys.stderr)
            sys.exit(1)

    if args.via_markdown:
        output = args.output or str(Path(args.input).with_suffix('.docx'))
        stats = run_pipeline_via_markdown(
            args.input, output,
            body_indent=args.body_indent,
            fix_citations=args.fix_citations,
            rules=rules,
            citation_rules=citation_rules)
        parts = []
        for k, label in [('title', '题目'), ('headings_l1', '一级标题'),
                         ('headings_l2', '二级标题'), ('headings_l3', '三级标题'),
                         ('body', '正文段落'), ('footnotes', '脚注')]:
            if stats.get(k):
                parts.append(f"{label}: {stats[k]}")
        parts.append(f"md: {Path(stats['md_path']).name}")
        print(f"✓ 完成 (双交付) — {' | '.join(parts)}")
        return

    output = args.output or args.input
    stats = run_pipeline(args.input, output,
                         body_indent=args.body_indent,
                         check_first=not args.no_check,
                         rules=rules)

    # Citation formatting
    if args.fix_citations:
        try:
            from citation_formatter import format_all_footnotes
            from docx import Document
            doc = Document(output)
            cstats = format_all_footnotes(doc, fix=True, style=citation_rules)
            doc.save(output)
            if cstats['fixed'] > 0:
                stats['citation_fixes'] = cstats['fixed']
                stats['citation_issues'] = cstats['issues']
        except ImportError:
            print("⚠ citation_formatter 模块未找到")

    # Print summary
    parts = []
    for k, label in [('title', '题目'), ('headings_l1', '一级标题'),
                     ('headings_l2', '二级标题'), ('headings_l3', '三级标题'),
                     ('body', '正文段落'), ('footnotes', '脚注')]:
        if stats.get(k):
            parts.append(f"{label}: {stats[k]}")
    if stats.get('citation_fixes'):
        parts.append(f"引注修复: {stats['citation_fixes']}处")
    print(f"✓ 完成 — {' | '.join(parts)}")


if __name__ == '__main__':
    main()
