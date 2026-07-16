#!/usr/bin/env python3
"""Single entry-point: runs the paper formatting pipeline.

Usage:
    python scripts/run_pipeline.py --input 论文.docx --output 论文_格式化.docx
"""

import argparse
from pathlib import Path
from format_paper import format_document, check_document, backup_file


def run_pipeline(input_path: str, output_path: str,
                 body_indent: int = 2, check_first: bool = True) -> dict:
    """Run the formatting pipeline on a single document.

    Args:
        input_path: Input .docx file.
        output_path: Output .docx file.
        body_indent: First-line indent for body text in characters.
        check_first: If True, run --check before formatting.

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
                            body_indent=body_indent)
    print(f"输出: {output_path}")
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
    args = ap.parse_args()

    output = args.output or args.input
    stats = run_pipeline(args.input, output,
                         body_indent=args.body_indent,
                         check_first=not args.no_check)

    # Citation formatting
    if args.fix_citations:
        try:
            from citation_formatter import format_all_footnotes
            from docx import Document
            doc = Document(output)
            cstats = format_all_footnotes(doc, fix=True)
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
