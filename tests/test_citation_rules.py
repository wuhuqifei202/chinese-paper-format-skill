"""Unit tests for citation checking and auto-fixing."""

import pytest

from citation_formatter import check_footnote, auto_fix_footnote


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
