"""Unit tests for heading detection and abstract/keywords detection."""

import pytest

from format_paper import detect_heading_level, is_abstract_or_keywords


# ── detect_heading_level ──────────────────────────────────────────────

@pytest.mark.pure
class TestDetectHeadingLevel:
    """Test detect_heading_level() function."""

    # ── Level 1 headings (一、) ──

    def test_level1_simple(self):
        """一、 should be detected as level 1."""
        assert detect_heading_level('一、问题的提出') == 1

    def test_level1_just_marker(self):
        """一、 alone should be level 1."""
        assert detect_heading_level('一、') == 1

    def test_level1_multi_digit(self):
        """二十一、 should be level 1 — multi-character numbering."""
        assert detect_heading_level('二十一、其他规定') == 1

    def test_level1_hundred(self):
        """一百二十三、 should be level 1."""
        assert detect_heading_level('一百二十三、附则') == 1

    def test_level1_dot_variant(self):
        """一． should work (full-width dot variant)."""
        assert detect_heading_level('一．概述') == 1

    def test_level1_too_long(self):
        """Very long text starting with heading pattern → not a heading."""
        long_text = '一、' + '这是非常长的标题' * 10  # > 40 chars
        assert detect_heading_level(long_text) == 0

    def test_false_level1_prefix(self):
        """第一， is a phrase, not a heading."""
        assert detect_heading_level('第一，从立法角度看') == 0

    def test_level1_with_prefix_char(self):
        """其、 is not a heading marker."""
        assert detect_heading_level('其一、合同无效') == 0
        assert detect_heading_level('其一') == 0  # too short but has prefix

    # ── Level 2 headings （一） ──

    def test_level2_simple(self):
        """（一） should be detected as level 2."""
        assert detect_heading_level('（一）国内研究现状') == 2

    def test_level2_multi_digit(self):
        """（十二） should be level 2."""
        assert detect_heading_level('（十二）小结') == 2

    def test_level2_too_long(self):
        """（一） with > 50 chars should not be a heading."""
        assert detect_heading_level('（一）' + '长' * 50) == 0

    # ── Level 3 headings (1. ) ──

    def test_level3_with_space(self):
        """1.  with space should be level 3."""
        assert detect_heading_level('1. 权利能力说') == 3

    def test_level3_no_space(self):
        """2.工具说 without space but short → level 3."""
        assert detect_heading_level('2.工具说') == 3

    def test_level3_no_space_long(self):
        """2.xxx without space and long → not a heading."""
        long_text = '2.' + '工' * 20
        assert detect_heading_level(long_text) == 0

    def test_level3_dot_number(self):
        """1.2 数据来源 — regex sees '1.' as L3 marker, not '1.2' as dotted number.

        Known limitation: the current regex-based approach cannot distinguish
        '1. 标题' (heading) from '1.2 数据' (dotted section number).
        This test documents current behavior; change to assert==0 if fixed.
        """
        assert detect_heading_level('1.2 数据来源') == 3  # current: matches as L3

    def test_level3_too_long(self):
        """1.  with > 60 chars should not be a heading."""
        assert detect_heading_level('1. ' + '长' * 60) == 0

    # ── Body text (level 0) ──

    def test_body_plain(self):
        """Plain text should be body (level 0)."""
        assert detect_heading_level('人工智能技术发展迅速') == 0

    def test_body_year_start(self):
        """2023年 should not match any heading pattern."""
        assert detect_heading_level('2023年，我国经济持续发展') == 0

    def test_body_empty(self):
        """Empty string should be body (level 0)."""
        assert detect_heading_level('') == 0

    def test_body_whitespace_only(self):
        """Whitespace-only should be body (level 0)."""
        assert detect_heading_level('   ') == 0


# ── is_abstract_or_keywords ────────────────────────────────────────────

@pytest.mark.pure
class TestAbstractKeywords:
    """Test is_abstract_or_keywords() function."""

    def test_abstract_bracket(self):
        """【摘要】 should be detected."""
        assert is_abstract_or_keywords('【摘要】本文研究了以物抵债') == 'abstract'

    def test_abstract_colon(self):
        """摘要： should be detected."""
        assert is_abstract_or_keywords('摘要：本文研究了') == 'abstract'

    def test_abstract_plain(self):
        """摘要 followed by text should be detected."""
        assert is_abstract_or_keywords('摘要 以物抵债是常见的交易安排') == 'abstract'

    def test_keywords_bracket(self):
        """【关键词】 should be detected."""
        assert is_abstract_or_keywords('【关键词】以物抵债；让与担保') == 'keywords'

    def test_keywords_colon(self):
        """关键词： should be detected."""
        assert is_abstract_or_keywords('关键词：以物抵债') == 'keywords'

    def test_keywords_plain(self):
        """关键词 followed by text should be detected."""
        assert is_abstract_or_keywords('关键词 以物抵债') == 'keywords'

    def test_normal_text(self):
        """Normal text should return None."""
        assert is_abstract_or_keywords('本文认为以物抵债协议') is None

    def test_normal_heading(self):
        """Heading text should return None."""
        assert is_abstract_or_keywords('一、绪论') is None


# ── Unnumbered headings ────────────────────────────────────────────────

@pytest.mark.pure
class TestUnnumberedHeadings:
    """Test detect_heading_level() for unnumbered chapter titles."""

    def test_introduction_as_level1(self):
        """引言 should be detected as level 1 heading."""
        assert detect_heading_level('引言') == 1

    def test_preface_as_level1(self):
        """前言 should be detected as level 1 heading."""
        assert detect_heading_level('前言') == 1

    def test_conclusion_as_level1(self):
        """结论 should be detected as level 1 heading."""
        assert detect_heading_level('结论') == 1

    def test_conclusion_variant_as_level1(self):
        """结语 should be detected as level 1 heading."""
        assert detect_heading_level('结语') == 1

    def test_prolegomenon_as_level1(self):
        """绪论 should be detected as level 1 heading."""
        assert detect_heading_level('绪论') == 1

    def test_not_unnumbered_heading(self):
        """Random words are not unnumbered headings."""
        assert detect_heading_level('研究背景') == 0
        assert detect_heading_level('问题提出') == 0

    def test_unnumbered_with_suffix(self):
        """引言 followed by extra text is NOT an unnumbered heading."""
        # Only exact match: '引言' not '引言：xxx' or '引言 xxx'
        assert detect_heading_level('引言：研究背景') == 0
