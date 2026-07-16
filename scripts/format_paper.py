#!/usr/bin/env python3
"""
Chinese Academic Paper Formatter — 中文学术论文格式规范化工具

按照中文学术论文排版规范，自动检测并设置 .docx 文件的格式：
  论文题目: 黑体, 四号(14pt), 不加粗, 居中
  一级标题 (一、): 宋体, 小四(12pt), 加粗, 单倍行距, 居中
  二级标题 (（一）): 楷体, 小四(12pt), 加粗, 单倍行距, 缩进2字符
  三级标题 (1. ): 宋体, 五号(10.5pt), 加粗, 单倍行距, 缩进2字符
  正文: 宋体(中文) + Times New Roman(西文), 五号(10.5pt), 单倍行距
  脚注: 宋体, 小五号(9pt), 单倍行距, 每页重新编号

Usage:
    python format_paper.py input.docx                  # 覆盖原文件
    python format_paper.py input.docx -o output.docx   # 输出到新文件
    python format_paper.py input.docx --check          # 仅检查不修改
    python format_paper.py folder/ --batch             # 批量处理
"""

import argparse
import re
import sys
import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.shared import Pt

# ---------------------------------------------------------------------------
# 字号常量 (单位: Pt)
# ---------------------------------------------------------------------------
SIZE_SIHAO = Pt(14)       # 四号    — 论文题目
SIZE_XIAOSI = Pt(12)      # 小四    — 一级标题 / 二级标题
SIZE_WUHAO = Pt(10.5)     # 五号    — 三级标题 / 正文
SIZE_XIAOWU = Pt(9)       # 小五号  — 脚注

# ---------------------------------------------------------------------------
# 字体常量
# ---------------------------------------------------------------------------
FONT_HEI = '黑体'                # 论文题目
FONT_SONG = '宋体'               # 一级/三级标题, 正文, 脚注
FONT_KAI = '楷体'                # 二级标题
FONT_TNR = 'Times New Roman'     # 西文默认

# ---------------------------------------------------------------------------
# 标题检测正则
# ---------------------------------------------------------------------------
# 一级标题: 一、 二、 … 十一、 二十、 …
_RE_LEVEL1 = re.compile(r'^[一二三四五六七八九十百千]+[、．.]')

# 二级标题: （一） （二） … （十一） …
_RE_LEVEL2 = re.compile(r'^（[一二三四五六七八九十百千]+）')

# 三级标题: 1.  2.  3.  …  或 1、 2、 … 或 1． …
_RE_LEVEL3 = re.compile(r'^\d{1,3}[\.、．]\s')

# 疑似三级标题但后面没有空格的情况 (1.xxx → 可能是序号)
_RE_LEVEL3_LOOSE = re.compile(r'^\d{1,3}[\.、．]\S')

# ---------------------------------------------------------------------------
# 最大标题长度 (字符) — 超过此长度的行不太可能是标题
# ---------------------------------------------------------------------------
MAX_HEADING_LEN = {
    1: 40,   # 一级标题最大字符数
    2: 50,   # 二级标题最大字符数
    3: 60,   # 三级标题最大字符数
}


def detect_heading_level(text: str) -> int:
    """检测段落文本的标题层级.

    Args:
        text: 段落文本 (已去除首尾空白).

    Returns:
        0 — 正文
        1 — 一级标题 (一、)
        2 — 二级标题 (（一）)
        3 — 三级标题 (1. )
    """
    if not text:
        return 0

    # 三级标题最容易误判 (普通数字开头), 用更严格的规则
    m3 = _RE_LEVEL3.match(text)
    m3_loose = _RE_LEVEL3_LOOSE.match(text)
    if m3:
        # "1. xxx" with space after dot — 高度可信
        if len(text) <= MAX_HEADING_LEN[3]:
            return 3
    elif m3_loose:
        # "1.xxx" without space — 可能是列表项, 仅在很短时判定为标题
        if len(text) <= 15:
            return 3
        return 0

    # 二级标题: （一）...
    m2 = _RE_LEVEL2.match(text)
    if m2 and len(text) <= MAX_HEADING_LEN[2]:
        return 2

    # 一级标题: 一、...
    # 增加保护: "第一，" "其一、" 等不是标题
    m1 = _RE_LEVEL1.match(text)
    if m1 and len(text) <= MAX_HEADING_LEN[1]:
        # 排除 "第一," "其一、" 等模式 (前面有"第","其"等字的不算)
        prefix = text[:m1.start()] if m1.start() > 0 else ''
        if prefix:
            return 0
        return 1

    return 0


def apply_run_font(run, cn_font: str, en_font: str, size,
                   bold: bool = False, italic: bool = False):
    """对单个 Run 设置字体属性.

    python-docx 高层 API 设置中文字体有限; 此处直接操作 XML 确保
    eastAsia (中文字体) 和 ascii/hAnsi (西文字体) 同时生效.
    """
    run.font.size = size
    run.bold = bold
    run.italic = italic

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} />')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)
    rFonts.set(qn('w:cs'), en_font)


def set_line_spacing(para, line_spacing: float = 1.0):
    """设置段落行距为单倍行距, 并清除段前/段后间距."""
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = parse_xml(f'<w:spacing {nsdecls("w")} />')
        pPr.append(spacing)

    # 单倍行距 = 240 twips (20 twips * 12pt * 1.0)
    spacing.set(qn('w:line'), str(int(line_spacing * 240)))
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:beforeLines'), '0')
    spacing.set(qn('w:afterLines'), '0')


def set_first_line_indent(para, chars: int = 2):
    """设置首行缩进 (字符数)."""
    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml(f'<w:ind {nsdecls("w")} />')
        pPr.append(ind)

    ind.set(qn('w:firstLineChars'), str(chars * 100))
    ind.set(qn('w:leftChars'), '0')
    ind.set(qn('w:hangingChars'), '0')
    # 清除绝对数值, 避免冲突
    for attr in ('w:firstLine', 'w:left', 'w:hanging'):
        val = ind.get(qn(attr))
        if val is not None:
            del ind.attrib[qn(attr)]


def remove_first_line_indent(para):
    """移除首行缩进."""
    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        for attr in ('w:firstLineChars', 'w:firstLine',
                     'w:leftChars', 'w:left',
                     'w:hangingChars', 'w:hanging'):
            val = ind.get(qn(attr))
            if val is not None:
                del ind.attrib[qn(attr)]


def set_alignment(para, alignment):
    """设置段落对齐方式."""
    para.alignment = alignment


def clear_paragraph_runs(para):
    """若段落无 run, 插入一个空 run (确保格式可写入)."""
    if not para.runs:
        # 使用 add_run 可能导致 XML 顺序问题; 直接用 lxml 插入
        r = parse_xml(f'<w:r {nsdecls("w")}><w:t xml:space="preserve"></w:t></w:r>')
        para._p.append(r)


def _get_footnotes_part(doc) -> Optional[object]:
    """获取文档的脚注 XML Part, 无脚注时返回 None."""
    FOOTNOTE_REL = ('http://schemas.openxmlformats.org/officeDocument/'
                    '2006/relationships/footnotes')
    for rel in doc.part.rels.values():
        if rel.reltype == FOOTNOTE_REL:
            return rel.target_part
    return None


def format_footnotes(doc):
    """格式化现有脚注: 宋体小五号, 单倍行距, 每页重新编号, 数字编号."""
    fn_part = _get_footnotes_part(doc)
    if fn_part is None:
        return  # 文档无脚注

    fn_xml = fn_part._element

    # 1) 设置脚注属性: 每页重新编号, 编号格式为 1,2,3...
    fn_pr = fn_xml.find(qn('w:footnotePr'))
    if fn_pr is None:
        fn_pr = parse_xml(f'<w:footnotePr {nsdecls("w")} />')
        fn_xml.insert(0, fn_pr)

    # numFmt: 数字编号
    num_fmt = fn_pr.find(qn('w:numFmt'))
    if num_fmt is None:
        num_fmt = parse_xml(f'<w:numFmt {nsdecls("w")} />')
        fn_pr.append(num_fmt)
    num_fmt.set(qn('w:val'), 'decimal')

    # numStart: 从 1 开始
    num_start = fn_pr.find(qn('w:numStart'))
    if num_start is None:
        num_start = parse_xml(f'<w:numStart {nsdecls("w")} />')
        fn_pr.append(num_start)
    num_start.set(qn('w:val'), '1')

    # numRestart: 0 = 每页重新开始 (restart per page)
    num_restart = fn_pr.find(qn('w:numRestart'))
    if num_restart is None:
        num_restart = parse_xml(f'<w:numRestart {nsdecls("w")} />')
        fn_pr.append(num_restart)
    num_restart.set(qn('w:val'), '0')

    # 2) 逐个脚注段落格式化
    for fn_elem in fn_xml.findall(qn('w:footnote')):
        fn_id_str = fn_elem.get(qn('w:id'))
        if fn_id_str is None:
            continue
        fn_id = int(fn_id_str)
        if fn_id <= 0:
            continue  # 跳过 -1 (分隔线) 和 0 (分隔线续)

        for p_elem in fn_elem.findall(qn('w:p')):
            _format_footnote_paragraph(p_elem)


def _format_footnote_paragraph(p_elem):
    """格式化单个脚注段落 XML 元素."""
    # 行距
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is None:
        pPr = parse_xml(f'<w:pPr {nsdecls("w")} />')
        p_elem.insert(0, pPr)

    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = parse_xml(f'<w:spacing {nsdecls("w")} />')
        pPr.append(spacing)
    spacing.set(qn('w:line'), '240')   # 单倍行距
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')

    # 每个 run: 宋体小五号 (9pt = 18 half-pt)
    for r_elem in p_elem.findall(qn('w:r')):
        rPr = r_elem.find(qn('w:rPr'))
        if rPr is None:
            rPr = parse_xml(f'<w:rPr {nsdecls("w")} />')
            r_elem.insert(0, rPr)

        # 字号
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = parse_xml(f'<w:sz {nsdecls("w")} />')
            rPr.append(sz)
        sz.set(qn('w:val'), '18')

        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = parse_xml(f'<w:szCs {nsdecls("w")} />')
            rPr.append(szCs)
        szCs.set(qn('w:val'), '18')

        # 字体
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} />')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), FONT_SONG)
        rFonts.set(qn('w:ascii'), FONT_TNR)
        rFonts.set(qn('w:hAnsi'), FONT_TNR)


def detect_title_range(paragraphs) -> Tuple[int, int]:
    """检测标题段落范围.

    第一个非空段落视为论文题目。若紧随的段落也为非空且非标题,
    则可能是副标题或作者信息, 一并纳入题目区域。

    Returns:
        (start_index, end_index_exclusive) — 题目区域在 paragraphs 中的范围.
    """
    start = -1
    end = 0

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            if start >= 0:
                # 空行标志着题目区域结束
                end = i
                break
            continue

        if start < 0:
            start = i
            end = i + 1
            continue

        # 题目之后: 若遇到标题关键词则结束题目区域
        if detect_heading_level(text) > 0:
            end = i
            break

        # 非空非标题可能是副标题/作者, 纳入 (但最多3段)
        end = i + 1
        if (end - start) >= 5:
            break

    return (start, end)


def format_document(input_path: str, output_path: str,
                   body_indent: int = 0) -> dict:
    """主格式化函数.

    Args:
        input_path: 输入 .docx 文件路径.
        output_path: 输出 .docx 文件路径.
        body_indent: 正文字符缩进数 (0=不缩进, 2=首行缩进2字符).

    Returns:
        统计信息 dict: {title, headings_l1, headings_l2, headings_l3,
                        body, footnotes, errors}
    """
    stats = {
        'title': 0, 'headings_l1': 0, 'headings_l2': 0,
        'headings_l3': 0, 'body': 0, 'footnotes': 0, 'errors': 0,
    }

    doc = Document(input_path)
    paragraphs = doc.paragraphs

    # —— 检测题目区域 ——
    title_start, title_end = detect_title_range(paragraphs)
    if title_start < 0:
        stats['errors'] += 1
        print("  ⚠ 未检测到任何文本内容")
        doc.save(output_path)
        return stats

    title_indices = set(range(title_start, title_end))

    # —— 逐段格式化 ——
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # 确保段落有 run
        clear_paragraph_runs(para)

        # 题目区域
        if i in title_indices:
            is_first_title = (i == title_start)
            set_alignment(para, WD_ALIGN_PARAGRAPH.CENTER)
            set_line_spacing(para)
            remove_first_line_indent(para)
            for run in para.runs:
                apply_run_font(run, FONT_HEI, FONT_TNR,
                              SIZE_SIHAO, bold=False)
            if is_first_title:
                stats['title'] += 1
            continue

        # 标题 / 正文
        level = detect_heading_level(text)

        if level == 1:
            set_alignment(para, WD_ALIGN_PARAGRAPH.CENTER)
            set_line_spacing(para)
            remove_first_line_indent(para)
            for run in para.runs:
                apply_run_font(run, FONT_SONG, FONT_TNR,
                              SIZE_XIAOSI, bold=True)
            stats['headings_l1'] += 1

        elif level == 2:
            set_alignment(para, WD_ALIGN_PARAGRAPH.LEFT)
            set_line_spacing(para)
            set_first_line_indent(para, chars=2)
            for run in para.runs:
                apply_run_font(run, FONT_KAI, FONT_TNR,
                              SIZE_XIAOSI, bold=True)
            stats['headings_l2'] += 1

        elif level == 3:
            set_alignment(para, WD_ALIGN_PARAGRAPH.LEFT)
            set_line_spacing(para)
            set_first_line_indent(para, chars=2)
            for run in para.runs:
                apply_run_font(run, FONT_SONG, FONT_TNR,
                              SIZE_WUHAO, bold=True)
            stats['headings_l3'] += 1

        else:
            # 正文
            set_alignment(para, WD_ALIGN_PARAGRAPH.LEFT)
            set_line_spacing(para)
            if body_indent > 0:
                set_first_line_indent(para, chars=body_indent)
            for run in para.runs:
                apply_run_font(run, FONT_SONG, FONT_TNR,
                              SIZE_WUHAO, bold=False)
            stats['body'] += 1

    # —— 格式化脚注 ——
    try:
        fn_part = _get_footnotes_part(doc)
        if fn_part is not None:
            fn_count = len([e for e in fn_part._element.findall(qn('w:footnote'))
                           if e.get(qn('w:id')) and int(e.get(qn('w:id'))) > 0])
            stats['footnotes'] = fn_count
            format_footnotes(doc)
    except Exception as e:
        print(f"  ⚠ 脚注处理异常: {e}")
        stats['errors'] += 1

    # —— 保存 ——
    doc.save(output_path)
    return stats


def check_document(input_path: str) -> dict:
    """检查模式: 仅读取并报告文档结构, 不修改.

    Returns:
        {paragraphs: 总段数, title: 题目文本, headings: [(层级, 文本),...], ...}
    """
    doc = Document(input_path)
    result = {
        'file': input_path,
        'total_paragraphs': len(doc.paragraphs),
        'title': '',
        'headings': [],
        'has_footnotes': False,
    }

    title_start, title_end = detect_title_range(doc.paragraphs)
    if title_start >= 0:
        titles = [doc.paragraphs[i].text.strip()
                  for i in range(title_start, title_end)
                  if doc.paragraphs[i].text.strip()]
        result['title'] = ' | '.join(titles)

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text or i < title_end:
            continue
        level = detect_heading_level(text)
        if level > 0:
            result['headings'].append((level, text[:60]))

    result['has_footnotes'] = _get_footnotes_part(doc) is not None
    return result


def backup_file(filepath: str) -> str:
    """创建备份文件 (原文件名.backup_时间戳.docx)."""
    p = Path(filepath)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = p.parent / f"{p.stem}.backup_{ts}{p.suffix}"
    shutil.copy2(filepath, str(backup))
    return str(backup)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description='中文学术论文格式规范化工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python format_paper.py 论文.docx                    # 覆盖原文件 (自动备份)
  python format_paper.py 论文.docx -o 论文_格式.docx  # 输出到新文件
  python format_paper.py 论文.docx --check             # 仅检查不修改
  python format_paper.py 论文/ --batch                 # 批量处理文件夹
  python format_paper.py 论文.docx --body-indent 2     # 正文首行缩进2字符
        """,
    )
    parser.add_argument('input', nargs='?', help='输入 .docx 文件路径, 或文件夹 (需配合 --batch)')
    parser.add_argument('-o', '--output', help='输出文件路径 (默认覆盖原文件)')
    parser.add_argument('--check', action='store_true',
                        help='仅检查文档结构, 不修改')
    parser.add_argument('--batch', action='store_true',
                        help='批量处理文件夹中所有 .docx 文件')
    parser.add_argument('--no-backup', action='store_true',
                        help='覆盖时不自动备份')
    parser.add_argument('--body-indent', type=int, default=0,
                        help='正文首行缩进字符数 (默认0, 推荐2)')
    parser.add_argument('--check-prereqs', action='store_true',
                        help='检查依赖和环境')
    parser.add_argument('--diagnostics', action='store_true',
                        help='输出技能元数据')

    args = parser.parse_args()

    # —— 诊断模式 ——
    if args.diagnostics:
        import json
        diag = {
            'skill': 'chinese-paper-format-skill',
            'version': '1.0.0',
            'harness_level': 'production',
            'commands': ['format_paper.py'],
            'harness_features': {
                'input_validation': True,
                'output_sanity': False,
                'check_prereqs': True,
                'diagnostics': True,
            },
        }
        print(json.dumps(diag, ensure_ascii=False, indent=2))
        return

    # —— 依赖检查 ——
    if args.check_prereqs:
        import json as _json
        checks = []

        # Python version
        import platform
        py_ok = sys.version_info >= (3, 8)
        checks.append({
            'check': 'python_version',
            'required': '>=3.8',
            'found': platform.python_version(),
            'ok': py_ok,
        })

        # python-docx
        try:
            import docx as _d
            checks.append({'check': 'python-docx', 'required': '>=1.0.0',
                          'found': _d.__version__, 'ok': True})
        except Exception:
            checks.append({'check': 'python-docx', 'required': '>=1.0.0',
                          'found': 'missing', 'ok': False})

        # lxml
        try:
            import lxml.etree
            checks.append({'check': 'lxml', 'required': '>=4.0',
                          'found': 'available', 'ok': True})
        except Exception:
            checks.append({'check': 'lxml', 'required': '>=4.0',
                          'found': 'missing', 'ok': False})

        ready = all(c['ok'] for c in checks)
        print(_json.dumps({'ready': ready, 'checks': checks},
                          ensure_ascii=False, indent=2))
        sys.exit(0 if ready else 1)

    # —— 输入验证 ——
    if not args.input:
        print('{"error": "请提供输入文件路径", '
              '"error_type": "validation", '
              '"hint": "用法: python format_paper.py <文件.docx>"}',
              file=sys.stderr)
        sys.exit(1)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'{{"error": "文件不存在: {args.input}", '
              f'"error_type": "validation", '
              f'"hint": "请检查文件路径是否正确"}}',
              file=sys.stderr)
        sys.exit(1)

    # —— 批量模式 ——
    if args.batch:
        if not input_path.is_dir():
            print(f'{{"error": "--batch 需要输入文件夹", '
                  f'"error_type": "validation"}}',
                  file=sys.stderr)
            sys.exit(1)

        docx_files = list(input_path.glob('*.docx'))
        if not docx_files:
            print(f'{{"error": "文件夹中未找到 .docx 文件", '
                  f'"error_type": "validation"}}',
                  file=sys.stderr)
            sys.exit(1)

        print(f"批量处理 {len(docx_files)} 个文件...\n")
        total_stats = {'files': len(docx_files), 'ok': 0, 'fail': 0}

        for f in docx_files:
            try:
                bkp = backup_file(str(f))
                print(f"  {f.name}  (备份: {Path(bkp).name})")
                stats = format_document(str(f), str(f),
                                        body_indent=args.body_indent)
                _print_stats(stats)
                total_stats['ok'] += 1
            except Exception as e:
                print(f"  ✗ {f.name}: {e}")
                total_stats['fail'] += 1

        print(f"\n完成: {total_stats['ok']} 成功, {total_stats['fail']} 失败")
        sys.exit(0 if total_stats['fail'] == 0 else 1)

    # —— 格式检查 (只读) ——
    if args.check:
        result = check_document(str(input_path))
        print(f"文件: {result['file']}")
        print(f"总段落: {result['total_paragraphs']}")
        print(f"题目: {result['title'][:80] if result['title'] else '(未检测到)'}")
        print(f"标题 ({len(result['headings'])} 个):")
        for level, txt in result['headings']:
            prefix = {1: '一、', 2: '（一）', 3: '1. '}.get(level, '?')
            print(f"  L{level} [{prefix}] {txt}")
        print(f"脚注: {'有' if result['has_footnotes'] else '无'}")
        return

    # —— 单文件格式化 ——
    # .doc 文件警告
    if input_path.suffix.lower() == '.doc':
        print("⚠ .doc 格式不支持直接处理。请先用 Word 或 LibreOffice 另存为 .docx。")
        print("  Word: 文件 → 另存为 → Word 文档 (*.docx)")
        print("  LibreOffice: soffice --headless --convert-to docx 文件.doc")
        sys.exit(1)

    if input_path.suffix.lower() not in ('.docx',):
        print(f'{{"error": "不支持的文件格式: {input_path.suffix}", '
              f'"error_type": "validation", '
              f'"hint": "请提供 .docx 文件"}}',
              file=sys.stderr)
        sys.exit(1)

    # 确定输出路径
    output_path = args.output if args.output else str(input_path)

    # 覆盖时自动备份
    if not args.output and not args.no_backup:
        bkp = backup_file(str(input_path))
        print(f"已备份: {Path(bkp).name}")

    # 执行格式化
    try:
        print(f"格式化: {input_path}")
        stats = format_document(str(input_path), output_path,
                                body_indent=args.body_indent)
        print(f"输出: {output_path}")
        _print_stats(stats)
        print("✓ 完成")
    except Exception as e:
        print(f'{{"error": "{e}", "error_type": "runtime", '
              f'"hint": "格式化过程中出现错误, 请检查文档是否损坏"}}',
              file=sys.stderr)
        sys.exit(1)


def _print_stats(stats: dict):
    """打印格式化统计."""
    parts = []
    if stats.get('title'):
        parts.append(f"题目: {stats['title']}")
    if stats.get('headings_l1'):
        parts.append(f"一级标题: {stats['headings_l1']}")
    if stats.get('headings_l2'):
        parts.append(f"二级标题: {stats['headings_l2']}")
    if stats.get('headings_l3'):
        parts.append(f"三级标题: {stats['headings_l3']}")
    if stats.get('body'):
        parts.append(f"正文段落: {stats['body']}")
    if stats.get('footnotes'):
        parts.append(f"脚注: {stats['footnotes']}")
    if stats.get('errors'):
        parts.append(f"异常: {stats['errors']}")
    if parts:
        print(f"  {' | '.join(parts)}")


if __name__ == '__main__':
    main()
