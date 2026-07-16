#!/usr/bin/env python3
"""
Citation Formatter for Chinese Legal Academic Papers
法学引注格式规范化模块

依据《法学引注手册》(2019) 的规范，检测并修复脚注中的引注格式问题。

支持的格式规则:
  - 中文标点符号规范化: ，、：。《》 vs ,.:《》"
  - 作者与文献名称之间用冒号
  - 文献名称使用书名号
  - 期刊/报纸/文集文章前加"载"
  - 页码格式: 第×页 / 第×-×页
  - 同一注释多文献用分号分隔
  - 引用符号位置 (句号后 vs 句中)
  - 同一文献多次出现时的略写格式
  - 引领词规范: 参见/见/又见/另见/转引自

Usage:
    python citation_formatter.py input.docx --check          # 仅检查
    python citation_formatter.py input.docx --fix            # 自动修复
    python citation_formatter.py input.docx -o output.docx   # 输出到新文件
"""

import re
import sys
import argparse
from pathlib import Path
from typing import List, Dict, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ---------------------------------------------------------------------------
# 中文标点映射 — 用于修复英文标点
# ---------------------------------------------------------------------------
# 注意: 不是简单替换，需要在语境中判断
_EN_TO_CN_PUNCT = {
    # 英文 → 中文 (仅当在中文语境中时替换)
    ',': '，',
    ':': '：',
    ';': '；',
    '(': '（',
    ')': '）',
}

# 需要保护的标点 (在数字、URL、外文等上下文中不替换)
_PROTECTED_CONTEXTS = [
    r'\d\.\d',           # 数字中的点号 (如 3.14)
    r'https?://',         # URL
    r'[A-Za-z]\.',        # 英文缩写
    r'\d,\d',             # 数字中的逗号
]

# ---------------------------------------------------------------------------
# 引注格式正则模式 — 用于检测和验证
# ---------------------------------------------------------------------------

# 期刊文章: 作者：《标题》，载《期刊》年份年第×期
_RE_JOURNAL = re.compile(
    r'([^，。：]+)[：:]\s*《([^》]+)》\s*[,，]\s*载《([^》]+)》\s*'
    r'(\d{4})\s*年\s*第\s*(\d+)\s*期'
)

# 图书: 作者：《书名》，出版社年份版
_RE_BOOK = re.compile(
    r'([^，。：]+)[：:]\s*《([^》]+)》\s*[,，]\s*'
    r'(.+出版社)\s*(\d{4})\s*年版'
)

# 报纸: 作者：《标题》，载《报纸》年月日
_RE_NEWSPAPER = re.compile(
    r'([^，。：]+)[：:]\s*《([^》]+)》\s*[,，]\s*载《([^》]+)》\s*'
    r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日'
)

# 文集文章: 作者：《标题》，载编者编：《文集》，出版社年份版
_RE_BOOK_ARTICLE = re.compile(
    r'([^，。：]+)[：:]\s*《([^》]+)》\s*[,，]\s*'
    r'载\s*([^：:]+)[：:]\s*《([^》]+)》'
)

# 翻译作品: [国籍]作者：《书名》，译者译，出版社年份版
_RE_TRANSLATION = re.compile(
    r'\[([^\]]+)\]\s*([^：:]+)[：:]\s*《([^》]+)》\s*[,，]\s*'
    r'([^，]+译)\s*[,，]\s*(.+)'
)

# 页码: 第×页 或 第×-×页
_RE_PAGE = re.compile(r'第\s*(\d+(?:[-–—]\d+)?)\s*页')

# 引领词
_LEADING_WORDS = ['参见', '见', '又见', '另见', '转引自', '同上注', '同前注']

# 常见错误模式
_ERROR_PATTERNS = [
    # 1. 作者名后逗号+书名号 → 冒号+书名号 (如 "王利明, 《" → "王利明：《")
    (re.compile(r'([一-鿿]{2,4})\s*[,，]\s*(《[^》]+》)'), r'\1：\2'),

    # 2. 英文页码格式 p.xx → 第xx页
    (re.compile(r'[,，]\s*[Pp]\.\s*(\d+)'), r'，第\1页'),
    (re.compile(r'([。；])\s*[Pp]\.\s*(\d+)'), r'\1第\2页'),

    # 3. 中英文标点混用
    (re.compile(r'([一-鿿）》)])\.([\s一-鿿（《])'), r'\1。\2'),   # 中文后的英文句号
    (re.compile(r'([一-鿿）》）]),(\s*)(?=[一-鿿（《])'), r'\1，\2'),  # 中文后的英文逗号 (非数字)
    (re.compile(r'([一-鿿）》）]);(\s*)(?=[一-鿿（《])'), r'\1；\2'),  # 中文后的英文分号
    (re.compile(r'([一-鿿））》]):(\s*)(?=[一-鿿（《])'), r'\1：\2'),  # 中文后的英文冒号

    # 4. "载"字缺失 (期刊/报纸文章 — 文章名后逗号+期刊名+年份)
    (re.compile(r'，《([^》]+)》(\d{4})\s*年\s*第'), r'，载《\1》\2年第'),

    # 5. 出版社格式不完整 (缺"年版")
    (re.compile(r'([一-鿿]+出版社)\s*(\d{4})\s*年\s*[,，]'), r'\1\2年版，'),
    (re.compile(r'([一-鿿]+出版社)\s*(\d{4})\s*年\s*([。；])'), r'\1\2年版\3'),

    # 6. 多余空格
    (re.compile(r'([，。：；、》）])\s+'), r'\1'),
    (re.compile(r'\s+([，。：；、《（])'), r'\1'),
]


def _get_footnotes_xml(doc) -> Optional[object]:
    """获取文档的脚注 XML 根元素."""
    FOOTNOTE_REL = ('http://schemas.openxmlformats.org/officeDocument/'
                    '2006/relationships/footnotes')
    for rel in doc.part.rels.values():
        if rel.reltype == FOOTNOTE_REL:
            from lxml import etree
            return etree.fromstring(rel.target_part.blob)
    return None


def extract_footnotes(doc: Document) -> List[Dict]:
    """从文档中提取所有脚注文本.

    Returns:
        [{id, paragraphs: [text], full_text: str}, ...]
    """
    fn_xml = _get_footnotes_xml(doc)
    if fn_xml is None:
        return []

    footnotes = []

    for fn_elem in fn_xml.findall(qn('w:footnote')):
        fn_id = int(fn_elem.get(qn('w:id'), '-1'))
        if fn_id <= 0:
            continue  # 跳过分隔脚注

        paragraphs = []
        for p_elem in fn_elem.findall(qn('w:p')):
            texts = []
            for r_elem in p_elem.findall(qn('w:r')):
                for t_elem in r_elem.findall(qn('w:t')):
                    if t_elem.text:
                        texts.append(t_elem.text)
            para_text = ''.join(texts)
            if para_text.strip():
                paragraphs.append(para_text)

        full_text = '\n'.join(paragraphs)
        footnotes.append({
            'id': fn_id,
            'paragraphs': paragraphs,
            'full_text': full_text.strip(),
        })

    return footnotes


def check_footnote(fn_text: str) -> List[Dict]:
    """检查单条脚注文本的格式问题.

    Returns:
        [{type, severity, message, suggestion, position}, ...]
    """
    issues = []

    # 1. 检查中英文标点混用
    _check_punctuation_mixing(fn_text, issues)

    # 2. 检查书名号使用
    _check_book_title_marks(fn_text, issues)

    # 3. 检查"载"字使用
    _check_zai_usage(fn_text, issues)

    # 4. 检查作者-标题分隔符
    _check_author_title_separator(fn_text, issues)

    # 5. 检查页码格式
    _check_page_format(fn_text, issues)

    # 6. 检查引注符号位置 (在正文中检查, 这里跳过)

    # 7. 检查多余空格
    _check_extra_spaces(fn_text, issues)

    # 8. 检查引领词
    _check_leading_words(fn_text, issues)

    # 9. 检查出版社格式
    _check_publisher_format(fn_text, issues)

    return issues


def _check_punctuation_mixing(text: str, issues: List[Dict]):
    """检测中文上下文中使用英文标点的问题."""
    # 中文句号后跟英文逗号
    if re.search(r'[一-鿿]\)\s*,', text):
        issues.append({
            'type': 'punctuation',
            'severity': 'medium',
            'message': '中文上下文中使用了英文标点',
            'suggestion': '将英文标点替换为中文标点（，、：；。）',
            'context': text[:80],
        })

    # 检查常见英文标点在中文后出现
    en_punct_in_cn = []
    for en_p, cn_p in _EN_TO_CN_PUNCT.items():
        if re.search(rf'[一-鿿》）]{re.escape(en_p)}\s', text):
            en_punct_in_cn.append(f'{en_p}→{cn_p}')

    if en_punct_in_cn:
        issues.append({
            'type': 'punctuation',
            'severity': 'low',
            'message': f'中文后使用了英文标点: {", ".join(en_punct_in_cn)}',
            'suggestion': '自动替换为中文标点',
            'context': text[:80],
        })


def _check_book_title_marks(text: str, issues: List[Dict]):
    """检测书名号使用是否规范."""
    # 文献标题应该用书名号
    # 检测"载"后面跟的内容是否用了书名号
    zai_match = re.search(r'载\s*([^《\s][^，。]+?)(?:\s*\d{4})', text)
    if zai_match:
        source_name = zai_match.group(1)
        if '《' not in source_name[:10]:
            issues.append({
                'type': 'title_mark',
                'severity': 'high',
                'message': f'"载"后的来源名称未使用书名号: {source_name[:30]}',
                'suggestion': f'应为: 载《{source_name.strip()}》',
                'context': text[:80],
            })

    # 检测期刊名是否用了书名号 (常见模式: "载XXX年第" 缺少书名号)
    journal_no_mark = re.search(r'载\s*([A-Za-z一-鿿（）]+?)\s*(\d{4})\s*年\s*第\s*\d+\s*期', text)
    if journal_no_mark and '《' not in journal_no_mark.group(1):
        jname = journal_no_mark.group(1)
        if jname.strip():
            issues.append({
                'type': 'title_mark',
                'severity': 'high',
                'message': f'期刊名缺少书名号: {jname}',
                'suggestion': f'应为: 《{jname}》',
                'context': text[:80],
            })


def _check_zai_usage(text: str, issues: List[Dict]):
    """检测"载"字是否按规定使用."""
    # 期刊文章: 应有"载《期刊名》"
    has_journal = re.search(r'\d{4}\s*年\s*第\s*\d+\s*期', text)
    if has_journal and '载' not in text:
        issues.append({
            'type': 'zai_missing',
            'severity': 'high',
            'message': '期刊文章缺少"载"字引导',
            'suggestion': '在文章名后、期刊名前添加"载"字',
            'context': text[:80],
        })

    # 报纸文章: 应有"载《报纸名》年月日"
    has_newspaper = re.search(r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日', text)
    if has_newspaper and '载' not in text and not re.search(r'出版社', text):
        issues.append({
            'type': 'zai_missing',
            'severity': 'high',
            'message': '报纸/新闻类文章缺少"载"字引导',
            'suggestion': '在文章名后、报纸名前添加"载"字',
            'context': text[:80],
        })


def _check_author_title_separator(text: str, issues: List[Dict]):
    """检测作者与文献名称之间的分隔符."""
    # 作者名后应该用冒号（：）而非其他符号分隔
    # 检测: 中文字符+英文冒号+书名号 → 应改为中文冒号
    wrong_sep = re.search(r'([一-鿿）\]\)])[:：]\s*(?![0-9])', text)
    # 这个太宽泛了，聚焦在作者-书名模式
    # 检测 姓名字段后跟书名号但用逗号分隔的情况
    comma_before_title = re.search(r'([一-鿿]{2,4}),\s*《', text)
    if comma_before_title:
        issues.append({
            'type': 'separator',
            'severity': 'medium',
            'message': '作者与文献名称之间使用了逗号而非冒号',
            'suggestion': f'将",《" 改为 "：《"',
            'context': comma_before_title.group()[:40],
        })


def _check_page_format(text: str, issues: List[Dict]):
    """检查页码格式."""
    # 应使用"第×页"而非"p.×"或"第×"
    if re.search(r'[Pp]\.\s*\d+', text):
        issues.append({
            'type': 'page_format',
            'severity': 'medium',
            'message': '使用了英文页码格式 (p.×)',
            'suggestion': '改为中文格式: 第×页',
            'context': text[:80],
        })

    # 页码范围应使用短横线 "-" 而非 "~" 或 "到"
    if re.search(r'第\s*\d+\s*[~～到至]\s*\d+\s*页', text):
        issues.append({
            'type': 'page_format',
            'severity': 'low',
            'message': '页码范围连接符不规范',
            'suggestion': '使用短横线: 第×-×页',
            'context': text[:80],
        })


def _check_extra_spaces(text: str, issues: List[Dict]):
    """检测不必要的空格."""
    # 中文标点前后不应有空格
    extra_space = re.findall(r'[一-鿿]\s{2,}[一-鿿]', text)
    if extra_space:
        issues.append({
            'type': 'spacing',
            'severity': 'low',
            'message': f'存在多余空格 ({len(extra_space)} 处)',
            'suggestion': '删除中文标点前后多余的空格',
            'context': text[:80],
        })


def _check_leading_words(text: str, issues: List[Dict]):
    """检查引领词使用是否规范."""
    # "转引自" 应有原始来源和转引来源
    if '转引自' in text:
        if not re.search(r'[。；]', text):
            issues.append({
                'type': 'leading_word',
                'severity': 'info',
                'message': '使用"转引自"但未标注转引来源',
                'suggestion': '应同时标注原始出处和转引出处',
                'context': text[:80],
            })

    # "参见" vs "见" 的用法提醒
    if '参见' in text and re.search(r'["""]', text):
        issues.append({
            'type': 'leading_word',
            'severity': 'info',
            'message': '直接引用原文建议使用"见"而非"参见"',
            'suggestion': '概括引用用"参见"，直接引用用"见"',
            'context': text[:80],
        })


def _check_publisher_format(text: str, issues: List[Dict]):
    """检查出版社信息格式."""
    # 出版社前不应写城市名
    city_pub = re.search(r'([一-鿿]{2,3})[：:]\s*([一-鿿]+出版社)', text)
    if city_pub:
        issues.append({
            'type': 'publisher',
            'severity': 'low',
            'message': f'出版社前有多余的城市名: {city_pub.group(1)}',
            'suggestion': f'只写"{city_pub.group(2)}"，不写城市名',
            'context': text[:80],
        })


def auto_fix_footnote(text: str) -> Tuple[str, int]:
    """自动修复脚注文本中的常见格式问题.

    Returns:
        (fixed_text, fix_count) — 修复后的文本和修复次数.
    """
    fixed = text
    count = 0

    # 应用错误模式修复
    for pattern, replacement in _ERROR_PATTERNS:
        new_text, n = pattern.subn(replacement, fixed)
        if n > 0:
            fixed = new_text
            count += n

    # 修复出版社格式: "××出版社××××年" → "××出版社××××年版"
    pub_fix = re.sub(
        r'([一-鿿]+出版社)\s*(\d{4})\s*年\s*$',
        r'\1\2年版',
        fixed
    )
    pub_fix = re.sub(
        r'([一-鿿]+出版社)\s*(\d{4})\s*年\s*[,，]',
        r'\1\2年版，',
        pub_fix
    )
    if pub_fix != fixed:
        count += 1
        fixed = pub_fix

    # 修复引注符号后的多余空格
    fixed = re.sub(r'^(\s*\d+)\s+(?=[一-鿿\[（])', r'\1 ', fixed)

    return fixed, count


def _get_footnotes_part(doc):
    """获取文档的脚注 Part 对象 (用于写回)."""
    FOOTNOTE_REL = ('http://schemas.openxmlformats.org/officeDocument/'
                    '2006/relationships/footnotes')
    for rel in doc.part.rels.values():
        if rel.reltype == FOOTNOTE_REL:
            return rel.target_part
    return None


def format_all_footnotes(doc: Document, fix: bool = True) -> Dict:
    """格式化文档中所有脚注的引注内容."""
    fn_xml = _get_footnotes_xml(doc)
    if fn_xml is None:
        return {'total': 0, 'issues': 0, 'fixed': 0, 'checked': [],
                'message': '文档中没有脚注'}

    fn_part = _get_footnotes_part(doc)
    stats = {'total': 0, 'issues': 0, 'fixed': 0, 'checked': []}

    for fn_elem in fn_xml.findall(qn('w:footnote')):
        fn_id = int(fn_elem.get(qn('w:id'), '-1'))
        if fn_id <= 0:
            continue

        stats['total'] += 1
        fn_info = {
            'id': fn_id,
            'issues': [],
            'fixed_count': 0,
        }

        for p_elem in fn_elem.findall(qn('w:p')):
            run_elements = list(p_elem.findall(qn('w:r')))
            if not run_elements:
                continue

            full_text = ''
            for r_elem in run_elements:
                for t_elem in r_elem.findall(qn('w:t')):
                    if t_elem.text:
                        full_text += t_elem.text

            if not full_text.strip():
                continue

            issues = check_footnote(full_text)
            fn_info['issues'].extend(issues)
            stats['issues'] += len(issues)

            if fix and issues:
                fixed_text, fix_count = auto_fix_footnote(full_text)
                if fix_count > 0:
                    fn_info['fixed_count'] += fix_count
                    stats['fixed'] += fix_count

                    first_t = None
                    for r_elem in run_elements:
                        t_elems = r_elem.findall(qn('w:t'))
                        if t_elems:
                            if first_t is None:
                                first_t = t_elems[0]
                                first_t.text = fixed_text
                                first_t.set(qn('xml:space'), 'preserve')
                            else:
                                for t in t_elems:
                                    t.text = ''

        stats['checked'].append(fn_info)

    # Write modified XML back to part
    if fix and stats['fixed'] > 0 and fn_part is not None:
        from lxml import etree
        fn_part._blob = etree.tostring(fn_xml, xml_declaration=True,
                                        encoding='UTF-8', standalone=True)

    return stats


def main():
    parser = argparse.ArgumentParser(
        description='法学引注格式规范化工具 — 依据《法学引注手册》(2019)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python citation_formatter.py 论文.docx --check          # 仅检查不修改
  python citation_formatter.py 论文.docx --fix             # 自动修复
  python citation_formatter.py 论文.docx -o 输出.docx      # 输出到新文件
        """,
    )
    parser.add_argument('input', help='输入 .docx 文件路径')
    parser.add_argument('-o', '--output', help='输出文件路径')
    parser.add_argument('--check', action='store_true',
                        help='仅检查引注格式, 不修改')
    parser.add_argument('--fix', action='store_true',
                        help='自动修复检测到的格式问题')

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'错误: 文件不存在: {args.input}', file=sys.stderr)
        sys.exit(1)

    if not args.check and not args.fix:
        args.check = True  # 默认检查模式

    doc = Document(str(input_path))

    # 先提取并显示所有脚注
    footnotes = extract_footnotes(doc)
    print(f'文档共有 {len(footnotes)} 条脚注\n')

    if args.check:
        print('=' * 60)
        print('引注格式检查报告')
        print('=' * 60)
        total_issues = 0

        for fn in footnotes:
            issues = check_footnote(fn['full_text'])
            total_issues += len(issues)
            if issues:
                print(f'\n[脚注 {fn["id"]}] {fn["full_text"][:80]}...')
                for iss in issues:
                    sev = {'high': '🔴', 'medium': '🟡', 'low': '🟢', 'info': 'ℹ️'}
                    print(f'  {sev.get(iss["severity"], "  ")} [{iss["type"]}] {iss["message"]}')
                    print(f'    建议: {iss["suggestion"]}')

        if total_issues == 0:
            print('\n✓ 未发现引注格式问题。')
        else:
            print(f'\n共发现 {total_issues} 个格式问题。')
            print('运行 --fix 可自动修复部分问题。')

    if args.fix:
        print('\n' + '=' * 60)
        print('自动修复引注格式')
        print('=' * 60)
        stats = format_all_footnotes(doc, fix=True)
        print(f'处理脚注: {stats["total"]} 条')
        print(f'发现问题: {stats["issues"]} 个')
        print(f'自动修复: {stats["fixed"]} 处')

        for fn_info in stats['checked']:
            if fn_info['issues']:
                print(f'\n[脚注 {fn_info["id"]}] '
                      f'{len(fn_info["issues"])} 个问题, '
                      f'{fn_info["fixed_count"]} 处已修复')

        output_path = args.output if args.output else str(input_path)
        doc.save(output_path)
        print(f'\n已保存至: {output_path}')

        unfixed = stats['issues'] - stats['fixed']
        if unfixed > 0:
            print(f'⚠ {unfixed} 个问题需手动处理 (复杂格式问题)')


if __name__ == '__main__':
    main()
