"""Unit tests for citation checking and auto-fixing."""

import re

import pytest

from citation_formatter import (check_footnote, auto_fix_footnote,
                                 _ERROR_PATTERNS, normalize_chinese_punctuation)


# ── Pattern Order Tests ─────────────────────────────────────────────────

@pytest.mark.pure
class TestPatternOrder:
    """验证修复规则的应用顺序 — 结构性修复必须在标点修复之前."""

    def test_structural_before_punctuation(self):
        """作者逗号 (rule 1, 结构) 先于中文标点 (rule 3, 标点)."""
        author_idx = None  # rule 1: 作者名 → 冒号
        punct_idx = None   # rule 3: 中文标点混用 (first sub-rule)
        for i, (pattern, _) in enumerate(_ERROR_PATTERNS):
            pstr = pattern.pattern
            if '{2,4}' in pstr and '《' in pstr:
                author_idx = i
            # rule 3: first sub-rule matches "中文)。" → "中文。."
            if r'([一-鿿）》)])\.' in pstr or \
               '一-鿿' in pstr and author_idx is not None and i > author_idx:
                if punct_idx is None:
                    punct_idx = i
        assert author_idx is not None, 'Rule 1 (author comma) not found'
        assert punct_idx is not None, 'Rule 3 (punctuation) not found'
        assert author_idx < punct_idx, \
            f'Rule 1 (idx={author_idx}) must precede rule 3 (idx={punct_idx})'

    def test_bracket_before_general_punctuation(self):
        """空格清理在最后执行."""
        # Space cleanup patterns use \\s+ to match whitespace
        # They should be at the very end of the list
        space_indices = []
        for i, (pattern, _) in enumerate(_ERROR_PATTERNS):
            pstr = pattern.pattern
            # pstr is from pattern.pattern — contains literal \s (backslash+s)
            if pstr.endswith('\\s+') or ('\\s+' in pstr and '字第' not in pstr):
                space_indices.append(i)
        assert len(space_indices) >= 2, \
            f'Space cleanup rules not found (found {len(space_indices)} at {space_indices})'
        last_idx = len(_ERROR_PATTERNS) - 1
        for si in space_indices:
            assert si >= last_idx - 2, \
                f'Space cleanup (idx={si}) should be at end (last={last_idx})'

    def test_all_patterns_compile(self):
        """所有规则的正则必须编译成功."""
        for i, (pattern, replacement) in enumerate(_ERROR_PATTERNS):
            assert isinstance(pattern, re.Pattern), \
                f'Pattern {i} is not compiled: {pattern}'
            # 验证替换字符串不含语法错误
            try:
                pattern.sub(replacement, '')
            except Exception as e:
                pytest.fail(f'Pattern {i} substitution failed: {e}')

    def test_rule_count_stable(self):
        """规则数量变化必须显式更新此测试."""
        # 当前 9 组规则, 含子规则共 22 条 regex (v1.1: 扩展了标点统一规则)
        # 修改 _ERROR_PATTERNS 时更新此数字
        expected_count = 21
        assert len(_ERROR_PATTERNS) == expected_count, \
            f'ERROR_PATTERNS count changed: {len(_ERROR_PATTERNS)} != {expected_count}.' \
            f' Update this test after intentional changes.'


# ── Helper ─────────────────────────────────────────────────────────────

def _issue_types(issues):
    """Return set of issue type strings from a list of issues."""
    return {iss['type'] for iss in issues}


def _has_issue(issues, issue_type):
    """Check if any issue matches a given type."""
    return any(iss['type'] == issue_type for iss in issues)


# ── check_footnote ─────────────────────────────────────────────────────

@pytest.mark.pure
class TestCheckFootnote:
    """Test check_footnote() — citation issue detection."""

    # ── zai_missing ──

    def test_zai_present_no_issue(self):
        """Correct format with 载 should have no issues."""
        text = '参见王利明：《侵权法》，载《中国法学》2020年第1期。'
        issues = check_footnote(text)
        assert not _has_issue(issues, 'zai_missing')

    def test_zai_missing_journal(self):
        """Journal article without 载 should be flagged."""
        text = '参见王利明：《侵权法》，《中国法学》2020年第1期。'
        issues = check_footnote(text)
        assert _has_issue(issues, 'zai_missing')

    def test_legal_article_no_zai_needed(self):
        """Law article citations don't need 载."""
        text = '《民法典》第68条规定……'
        issues = check_footnote(text)
        assert not _has_issue(issues, 'zai_missing')

    # ── legal_bracket ──

    def test_legal_bracket_wrong(self):
        """Square brackets in document number should be flagged."""
        text = '《国务院通知》，国发[2007]19号。'
        issues = check_footnote(text)
        assert _has_issue(issues, 'legal_bracket')

    def test_legal_bracket_ok(self):
        """Correct 六角括号 should pass."""
        text = '《国务院通知》，国发〔2007〕19号。'
        issues = check_footnote(text)
        assert not _has_issue(issues, 'legal_bracket')

    # ── case_bracket ──

    def test_case_bracket_wrong(self):
        """Square brackets in case number should be flagged."""
        text = '北京市海淀区人民法院[1998]海行初字第142号行政判决书。'
        issues = check_footnote(text)
        assert _has_issue(issues, 'case_bracket')

    def test_case_bracket_ok(self):
        """Parentheses in case number should pass."""
        text = '北京市海淀区人民法院（1998）海行初字第142号行政判决书。'
        issues = check_footnote(text)
        assert not _has_issue(issues, 'case_bracket')

    # ── separator (author–title) ──

    def test_author_comma_not_colon(self):
        """Comma between author and title should be flagged."""
        text = '王利明, 《侵权责任法》'
        issues = check_footnote(text)
        assert _has_issue(issues, 'separator')

    def test_author_colon_ok(self):
        """Colon between author and title should pass."""
        text = '参见王利明：《侵权责任法》'
        issues = check_footnote(text)
        assert not _has_issue(issues, 'separator')

    # ── page_format ──

    def test_page_english_format(self):
        """English page format p.xx should be flagged."""
        text = '出版社2010年版，p.89。'
        issues = check_footnote(text)
        assert _has_issue(issues, 'page_format')

    def test_page_chinese_format_ok(self):
        """Chinese page format should pass."""
        text = '出版社2010年版，第89页。'
        issues = check_footnote(text)
        assert not _has_issue(issues, 'page_format')

    # ── leading_word ──

    def test_zhuan_yin_zi_flag(self):
        """转引自 should trigger a leading word check."""
        text = '转引自王利明：《侵权法》'
        issues = check_footnote(text)
        assert _has_issue(issues, 'leading_word')

    # ── Clean text ──

    def test_clean_text_no_issues(self):
        """A properly formatted citation should have no issues."""
        text = '参见王利明：《侵权责任法》，中国人民大学出版社2010年版，第125页。'
        issues = check_footnote(text)
        assert len(issues) == 0

    # ── Compound ──

    def test_compound_issues(self):
        """Text with multiple problems should report all of them."""
        text = '王利明, 《侵权法》, 中国人民大学出版社2016年, p.125.'
        issues = check_footnote(text)
        types = _issue_types(issues)
        assert 'separator' in types
        assert 'page_format' in types
        assert len(issues) >= 2


# ── auto_fix_footnote ───────────────────────────────────────────────────

@pytest.mark.pure
class TestAutoFix:
    """Test auto_fix_footnote() — automatic citation repair."""

    def test_fix_author_comma(self):
        """Comma between author and title → colon."""
        text = '王利明, 《侵权责任法》，出版社2010年版。'
        fixed, count = auto_fix_footnote(text)
        assert '王利明：《侵权责任法》' in fixed
        assert count >= 1

    def test_fix_page_english(self):
        """p.89 → 第89页."""
        text = '出版社2010年版，p.89。'
        fixed, count = auto_fix_footnote(text)
        assert '第89页' in fixed
        assert count >= 1

    def test_fix_publisher_missing_ban(self):
        """出版社2016年 → 出版社2016年版."""
        text = '中国人民大学出版社2016年，第125页。'
        fixed, count = auto_fix_footnote(text)
        assert '2016年版' in fixed
        assert count >= 1

    def test_fix_legal_bracket(self):
        """国发[2007]19号 → 国发〔2007〕19号."""
        text = '国发[2007]19号。'
        fixed, count = auto_fix_footnote(text)
        assert '国发〔2007〕19号' in fixed
        assert count >= 1

    def test_fix_case_bracket(self):
        """[1998]海行初字 → （1998）海行初字."""
        text = '[1998]海行初字第142号行政判决书。'
        fixed, count = auto_fix_footnote(text)
        assert '（1998）海行初字第142号' in fixed
        assert count >= 1

    def test_fix_zai_missing(self):
        """, 《学术研究》2020年 → ，载《学术研究》2020年."""
        text = '参见姚辉：《探究》，《学术研究》2020年第8期。'
        fixed, count = auto_fix_footnote(text)
        assert '载《学术研究》' in fixed
        assert count >= 1

    def test_compound_fix(self):
        """Multiple fixes in one text should all apply."""
        text = '王利明, 《侵权法》, 中国人民大学出版社2016年, p.125.'
        fixed, count = auto_fix_footnote(text)
        assert count >= 3  # author comma, publisher版, p.xx
        # Verify the cleaned result
        issues = check_footnote(fixed)
        assert len(issues) == 0

    def test_no_fix_needed(self):
        """Clean text should have zero fixes."""
        text = '参见王利明：《侵权责任法》，中国人民大学出版社2010年版，第125页。'
        fixed, count = auto_fix_footnote(text)
        assert count == 0
        assert fixed == text


@pytest.mark.pure
class TestPunctuationNormalization:
    """正文标点规范化 (回归: text-biaodian C15/C17 缺陷)."""

    def test_mixed_abbrev_bracket_paired(self):
        """括号内含英文缩写 (e.g./i.e.): 整对转全角 (回归修复)."""
        assert normalize_chinese_punctuation('说明(e.g.方法论).') == '说明（e.g.方法论）。'
        assert normalize_chinese_punctuation('(i.e.域外经验).') == '（i.e.域外经验）。'

    def test_cn_bracket_paired(self):
        """括号包中文: 整对转全角."""
        assert normalize_chinese_punctuation('(即以物抵债协议)') == '（即以物抵债协议）'

    def test_url_following_comma(self):
        """URL 后接中文的英文逗号 → 中文逗号 (回归修复)."""
        got = normalize_chinese_punctuation('见https://www.example.com,该网站')
        assert 'www.example.com，该网站' in got

    def test_numeric_comma_protected(self):
        """千位分隔 1,000 应保留."""
        got = normalize_chinese_punctuation('样本量为1,000条.')
        assert '1,000条。' in got

    def test_decimal_protected(self):
        """小数 3.14 应保留."""
        got = normalize_chinese_punctuation('如3.14.')
        assert '如3.14。' in got

    def test_url_dots_protected(self):
        """URL 的点号应保留 (含行尾, 保护优先于句号转换)."""
        assert normalize_chinese_punctuation('见https://www.example.com.') == \
            '见https://www.example.com.'
        # URL 之后有中文时, 中文句号才转换
        got = normalize_chinese_punctuation('见https://www.example.com.该网站.')
        assert 'www.example.com.该网站。' in got

    def test_pure_number_paren_protected(self):
        """纯数字括号 (2020) 保持半角."""
        got = normalize_chinese_punctuation('报告(2020)发布.')
        assert '（2020）' not in got
