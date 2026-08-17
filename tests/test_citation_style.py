"""Unit tests for the customizable citation style (注释体例) feature."""

import json

import pytest

import citation_rules
from citation_rules import (load_citation_rules, describe_style, CATEGORY_ORDER)
from citation_formatter import auto_fix_footnote, _ERROR_PATTERNS


# ── load_citation_rules ────────────────────────────────────────────────

@pytest.mark.pure
class TestLoadCitationRules:
    def test_default_is_fayin_manual(self):
        """默认体例为《法学引注手册》(2019), 而非《法学家》《中外法学》注释体例."""
        style = load_citation_rules()
        assert '法学引注手册' in style['name']

    def test_default_has_category_keys(self):
        style = load_citation_rules()
        assert list(style['categories'].keys()) == list(CATEGORY_ORDER)
        assert len(CATEGORY_ORDER) == 10

    def test_default_edition_translator_off(self):
        """法学引注手册默认关闭版次归位与译者空格清理."""
        style = load_citation_rules()
        assert style['rules']['edition'] is False
        assert style['rules']['translator'] is False
        assert style['rules']['punctuation'] is True
        assert style['rules']['author_separator'] is True

    def test_default_book_examples(self):
        style = load_citation_rules()
        examples = style['categories']['book']['examples']
        assert any('王利明' in e for e in examples)

    def test_preset_ten_category(self):
        """《法学家》《中外法学》注释体例预设开启版次归位."""
        style = load_citation_rules(preset='《法学家》《中外法学》注释体例')
        assert style['rules']['edition'] is True
        assert style['rules']['translator'] is True
        assert any('韩大元' in e for e in style['categories']['book']['examples'])

    def test_preset_ten_category_english_note(self):
        style = load_citation_rules(preset='《法学家》《中外法学》注释体例')
        assert '不要缩写' in style['categories']['english']['note']

    def test_preset_zhongguo_faxue_same_as_default(self):
        """《中国法学》的引注体例与《法学引注手册》一致."""
        style = load_citation_rules(preset='《中国法学》')
        assert '《中国法学》' in style['name']
        assert style['rules']['edition'] is False
        assert style['rules']['translator'] is False
        assert style['categories']['journal']['name'] == '期刊论文'

    def test_preset_fashang_yanjiu_same_as_default(self):
        """《法商研究》的引注体例与《法学引注手册》一致."""
        style = load_citation_rules(preset='《法商研究》')
        assert '《法商研究》' in style['name']
        assert style['rules']['edition'] is False

    def test_preset_faxue_yanjiu_notes(self):
        """《法学研究》在法学引注手册基础上新增 14 条要求 + 特定类别格式."""
        style = load_citation_rules(preset='《法学研究》')
        assert '《法学研究》' in style['name']
        assert len(style.get('notes', [])) == 14
        # 文集用「载于」连接
        assert '载于' in style['categories']['collection']['template']
        # 报纸标注版面序号
        assert '第X版' in style['categories']['newspaper']['template']
        # 网络资料只注网址与访问日期
        assert style['categories']['web']['template'].startswith('作者：《标题》，网址')

    def test_unknown_preset(self):
        with pytest.raises(ValueError):
            load_citation_rules(preset='不存在的预设')

    def test_preset_and_custom_merge(self, tmp_path):
        """预设之上再叠加自定义 (只写要改的部分)."""
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({'rules': {'edition': False}},
                                  ensure_ascii=False), encoding='utf-8')
        style = load_citation_rules(str(cfg), preset='《法学家》《中外法学》注释体例')
        assert style['rules']['edition'] is False   # 覆盖预设
        assert style['rules']['translator'] is True  # 保留预设

    def test_custom_toggle_on_default(self, tmp_path):
        """在默认法学引注手册上开启版次归位."""
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({'rules': {'edition': True}},
                                  ensure_ascii=False), encoding='utf-8')
        style = load_citation_rules(str(cfg))
        assert style['rules']['edition'] is True
        assert '法学引注手册' in style['name']

    def test_custom_conventions_override(self, tmp_path):
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({'conventions': {'author_separator': '，'}},
                                  ensure_ascii=False), encoding='utf-8')
        style = load_citation_rules(str(cfg))
        assert style['conventions']['author_separator'] == '，'

    def test_custom_category_override(self, tmp_path):
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({
            'categories': {'book': {'template': '作者：《书名》，出版社YYYY年版。'}},
        }, ensure_ascii=False), encoding='utf-8')
        style = load_citation_rules(str(cfg))
        assert style['categories']['book']['template'] == \
            '作者：《书名》，出版社YYYY年版。'
        assert style['categories']['journal']['name'] == '期刊论文'

    def test_underscore_keys_ignored(self, tmp_path):
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({
            '_说明': '注释',
            'rules': {'_备注': True, 'edition': True},
        }, ensure_ascii=False), encoding='utf-8')
        style = load_citation_rules(str(cfg))
        assert style['rules']['edition'] is True


@pytest.mark.pure
class TestValidation:
    def test_unknown_rule(self, tmp_path):
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({'rules': {'nope': True}}), encoding='utf-8')
        with pytest.raises(ValueError):
            load_citation_rules(str(cfg))

    def test_non_bool_rule(self, tmp_path):
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({'rules': {'edition': 'yes'}}),
                       encoding='utf-8')
        with pytest.raises(ValueError):
            load_citation_rules(str(cfg))

    def test_unknown_category(self, tmp_path):
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({'categories': {'patent': {}}}),
                       encoding='utf-8')
        with pytest.raises(ValueError):
            load_citation_rules(str(cfg))

    def test_bad_examples_type(self, tmp_path):
        cfg = tmp_path / 'style.json'
        cfg.write_text(json.dumps({'categories': {'book': {'examples': 'x'}}}),
                       encoding='utf-8')
        with pytest.raises(ValueError):
            load_citation_rules(str(cfg))

    def test_missing_file(self):
        with pytest.raises(FileNotFoundError):
            load_citation_rules('不存在的文件.json')


# ── describe_style ─────────────────────────────────────────────────────

@pytest.mark.pure
class TestDescribeStyle:
    def test_describe_lists_categories(self):
        text = describe_style(load_citation_rules())
        assert '图书' in text
        assert '期刊论文' in text
        assert '英文文献' in text

    def test_describe_ten_category_mentions_abbrev(self):
        text = describe_style(load_citation_rules(preset='《法学家》《中外法学》注释体例'))
        assert '英文期刊名称不要缩写' in text


# ── auto_fix_footnote with style ───────────────────────────────────────

@pytest.mark.pure
class TestAutoFixStyle:
    EDITION_TEXT = '韩大元：《亚洲立宪主义研究》第2版，中国人民公安大学出版社2008年版，第277-278页。'

    def test_edition_not_fixed_by_default(self):
        """默认法学引注手册不做版次归位."""
        fixed, _ = auto_fix_footnote(self.EDITION_TEXT)
        assert '（第2版）' not in fixed
        assert '第2版' in fixed

    def test_edition_fixed_with_preset(self):
        style = load_citation_rules(preset='《法学家》《中外法学》注释体例')
        fixed, count = auto_fix_footnote(self.EDITION_TEXT, style=style)
        assert '（第2版）' in fixed
        assert count >= 1

    def test_edition_fixed_when_enabled(self):
        style = load_citation_rules()
        style['rules']['edition'] = True
        fixed, count = auto_fix_footnote(self.EDITION_TEXT, style=style)
        assert '（第2版）' in fixed
        assert count >= 1

    def test_edition_already_parenthesized_no_change(self):
        text = '王泽鉴：《不当得利》（第2版），第321页。'
        fixed, count = auto_fix_footnote(text)
        assert fixed == text
        assert count == 0

    def test_translator_not_fixed_by_default(self):
        text = '黄家镇 译，商务印书馆2020年版，第159页。'
        fixed, _ = auto_fix_footnote(text)
        assert '黄家镇 译' in fixed

    def test_translator_fixed_with_preset(self):
        style = load_citation_rules(preset='《法学家》《中外法学》注释体例')
        text = '黄家镇 译，商务印书馆2020年版，第159页。'
        fixed, count = auto_fix_footnote(text, style=style)
        assert '黄家镇译' in fixed
        assert count >= 1

    def test_clean_text_no_change(self):
        text = '参见王利明：《侵权责任法》，中国人民大学出版社2010年版，第125页。'
        fixed, count = auto_fix_footnote(text)
        assert fixed == text
        assert count == 0


@pytest.mark.pure
class TestBackwardCompat:
    def test_error_patterns_count_stable(self):
        """重构为命名规则组后, 扁平规则表数量保持不变."""
        assert len(_ERROR_PATTERNS) == 21
