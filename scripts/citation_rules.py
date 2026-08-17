#!/usr/bin/env python3
"""citation_rules — 注释体例（引注格式）配置表.

让使用者按自己的注释体例要求便捷转换脚注引注格式。**默认体例依据《法学引注
手册》(2019)**；另外内置《法学家》《中外法学》《中国法学》《法商研究》
《法学研究》等常见法学期刊的注释体例预设，可按需切换或在其上定制。

与 `rules.py` (排版格式) 对应, 本模块负责「注释/引注」这一维度的可配置化:

    {
        "_说明": "只写要改的部分, 未写出的用默认",
        "rules": {"edition": true, "translator": true},
        "conventions": {"page_prefix": "第", "page_suffix": "页"},
        "categories": {"book": {"template": "…"}}
    }

三个可配置维度:
  - rules       每个修复规则组的开关 (bool)
  - conventions 若干命名约定 (字符串, 供修复与文档使用)
  - categories  各类文献体例的模板/示例 (用于展示与人工核对)

用法:
    from citation_rules import load_citation_rules, describe_style, PRESETS
    style = load_citation_rules('我的注释体例.json')          # 默认法学引注手册 + 定制
    style = load_citation_rules(preset='《法学家》《中外法学》注释体例')     # 切换到两刊注释体例

    # 命令行校验并打印体例
    python citation_rules.py [配置.json] [--preset 《法学家》《中外法学》注释体例]
"""

import json
from copy import deepcopy
from pathlib import Path
from typing import Dict, List, Optional

# ---------------------------------------------------------------------------
# 文献类别 (键序) — 两种体例共用
# ---------------------------------------------------------------------------
CATEGORY_ORDER = (
    'book', 'journal', 'collection', 'translation', 'newspaper',
    'ancient', 'dictionary', 'web', 'english', 'foreign',
)

# ---------------------------------------------------------------------------
# 默认注释体例 — 依据《法学引注手册》(2019)
# ---------------------------------------------------------------------------
DEFAULT_CITATION_STYLE: Dict = {
    'name': '法学引注手册（2019）',
    'description': (
        '依据《法学引注手册》(2019) 的脚注引注体例（skill 默认）。'
    ),
    'categories': {
        'book': {
            'name': '图书',
            'template': '作者：《书名》，出版社YYYY年版，第N页。',
            'examples': [
                '王利明：《侵权责任法》，中国人民大学出版社2010年版，第125页。',
            ],
        },
        'journal': {
            'name': '期刊论文',
            'template': '作者：《标题》，载《期刊》YYYY年第N期，第N页。',
            'examples': [
                '张文显：《法治与国家治理现代化》，载《中国法学》2014年第4期，第25页。',
            ],
        },
        'collection': {
            'name': '文集文章',
            'template': '作者：《标题》，载编者主编：《文集》，出版社YYYY年版，第N页。',
            'examples': [
                '王轶：《诉讼时效制度三论》，载崔建远主编：《民法九人行》（第7卷），法律出版社2014年版，第98页。',
            ],
        },
        'translation': {
            'name': '译著',
            'template': '[国籍]作者：《书名》，译者译，出版社YYYY年版，第N页。',
            'examples': [
                '[德]卡尔·拉伦茨：《法学方法论》，黄家镇译，商务印书馆2020年版，第159页。',
            ],
        },
        'newspaper': {
            'name': '报纸文章',
            'template': '作者：《标题》，载《报纸》YYYY年M月D日。',
            'examples': [
                '史际春：《以法治保障供给侧结构性改革》，载《人民日报》2016年11月2日。',
            ],
        },
        'ancient': {
            'name': '古籍',
            'template': '（朝代）作者：《书名》，版本，卷N，篇名，第N页。',
            'examples': [
                '（清）方大湜：《平平言》，清光绪十八年（1892）资州官廨刊本，卷三，第43页。',
            ],
        },
        'dictionary': {
            'name': '辞书',
            'template': '《词典》，出版社YYYY年版，第N页。',
            'examples': [
                '《元照英美法词典》，法律出版社2003年版，第9页。',
            ],
        },
        'web': {
            'name': '网络资料',
            'template': '作者：《标题》，载网站名，URL，YYYY年M月D日访问。',
            'examples': [
                '郑成思：《“入世”、知识产权保护与民商法的现代化》，载中国法学网，http://www.iolaw.org.cn/showNews.asp?id=243，2007年4月29日访问。',
            ],
        },
        'english': {
            'name': '英文文献',
            'template': 'Author, Title, City: Publisher, Year, p.xx.',
            'examples': [
                'L. Fuller, The Morality of Law, revised edition, New Haven: Yale University Press, 1969, p.143.',
            ],
        },
        'foreign': {
            'name': '其他外文文献',
            'template': '西文体例比照英文，日文体例比照中文。',
            'examples': [],
            'note': '西文（法文、德文、意大利文、西班牙文等）体例比照英文；日文体例比照中文。',
        },
    },
    'conventions': {
        'author_separator': '：',          # 作者与文献名之间的分隔符
        'multiple_authors_separator': '、',  # 多位作者之间的分隔符
        'editor_suffix': '主编',           # 主编署名后缀
        'translator_suffix': '译',          # 译者署名后缀
        'journal_marker': '载',            # 期刊/报纸/文集文章前的引导字
        'publisher_year_suffix': '年版',    # 出版社年份后缀
        'edition_marker': '（第{n}版）',    # 版次标记
        'page_prefix': '第',                # 页码前缀
        'page_suffix': '页',                # 页码后缀
        'page_range_separator': '-',        # 页码范围连接符
        'english_journal_abbrev': False,    # 英文期刊名是否缩写 (False = 不缩写)
    },
    'rules': {
        'punctuation': True,       # 中文标点规范化
        'author_separator': True,  # 作者名后逗号 → 冒号
        'zai_marker': True,        # 期刊/报纸文章补 "载"
        'publisher_year': True,    # 出版社补 "年版"
        'page_format': True,       # 英文页码 p.xx → 第xx页
        'edition': False,          # 版次归入（第N版）— 默认关闭
        'translator': False,       # 译者名与 "译" 之间的空格清理 — 默认关闭
        'legal_bracket': True,     # 文号方括号 → 六角括号
        'case_bracket': True,      # 案号方括号 → 圆括号
    },
}

# ---------------------------------------------------------------------------
# 《法学家》《中外法学》注释体例 (预设, 非默认)
# ---------------------------------------------------------------------------
JOURNAL_CITATION_STYLE: Dict = {
    'name': '《法学家》《中外法学》注释体例',
    'description': (
        '《法学家》《中外法学》两刊的注释体例（文献引用格式）：著作教材、论文、'
        '文集、译著、报纸、古籍、辞书、网络资料、英文、其他外文等十类文献的'
        '脚注引注体例（预设, 需显式选择）。'
    ),
    'categories': {
        'book': {
            'name': '著作、教材类',
            'template': '作者：《书名》（第N版），出版社YYYY年版，第N页。',
            'examples': [
                '《马克思恩格斯文集》（第1卷），人民出版社2009年版，第460页。',
                '韩大元：《亚洲立宪主义研究》（第2版），中国人民公安大学出版社2008年版，第277-278页。',
                '胡鸿烈、钟期荣：《香港的婚姻与继承法》，香港南天书业公司1957年版，第115页。',
                '佟柔主编：《中国民法》，法律出版社1990年版，第67页。',
            ],
        },
        'journal': {
            'name': '论文类',
            'template': '作者：《标题》，载《期刊》YYYY年第N期，第N页。',
            'examples': [
                '张文显：《法治与国家治理现代化》，载《中国法学》2014年第4期，第25页。',
                '王利明：《隐私权的新发展》，载《人大法律评论》（2009年卷），法律出版社2009年版，第14页。',
                '[德]格哈特·瓦格纳：《当代侵权法比较研究》，高圣平、熊丙万译，载《法学家》2010年第2期，第103-127页。',
            ],
        },
        'collection': {
            'name': '文集类',
            'template': '作者：《标题》，载编者主编：《文集》（第N卷），出版社YYYY年版，第N页。',
            'examples': [
                '王轶：《诉讼时效制度三论》，载崔建远主编：《民法九人行》（第7卷），法律出版社2014年版，第98页。',
            ],
        },
        'translation': {
            'name': '译著类',
            'template': '[国籍]作者：《书名》（全本·第N版），译者译，出版社YYYY年版，第N页。',
            'examples': [
                '[德]卡尔•拉伦茨：《法学方法论》（全本·第6版），黄家镇译，商务印书馆2020年版，第159页。',
            ],
        },
        'newspaper': {
            'name': '报纸类',
            'template': '作者：《标题》，载《报纸》YYYY年M月D日。',
            'examples': [
                '史际春：《以法治保障供给侧结构性改革》，载《人民日报》2016年11月2日。',
            ],
        },
        'ancient': {
            'name': '古籍类',
            'template': '（朝代）作者：《书名》，版本，卷N，篇名，第N页。',
            'examples': [
                '（清）方大湜：《平平言》，清光绪十八年（1892）资州官廨刊本，卷三，“斗殴先下手者宜重治”，第43页a。',
            ],
        },
        'dictionary': {
            'name': '辞书类',
            'template': '《词典》，出版社YYYY年版，第N页。',
            'examples': [
                '《元照英美法词典》，法律出版社2003年版，第9页。',
            ],
        },
        'web': {
            'name': '网络资料类',
            'template': '作者：《标题》，载网站名，URL，YYYY年M月D日访问。',
            'examples': [
                '郑成思：《“入世”、知识产权保护与民商法的现代化》，载中国法学网，http://www.iolaw.org.cn/showNews.asp?id=243，2007年4月29日访问。',
            ],
        },
        'english': {
            'name': '英文类',
            'template': 'Author, Title, edition, City: Publisher, Year, p.xx.',
            'examples': [
                'L. Fuller, The Morality of Law, revised edition, New Haven: Yale University Press, 1969, p.143.',
                'See Tom Ginsburg, “East Asian Constitutionalism in Comparative Perspective”, in Albert H. Y. Chen, ed., Constitutionalism in Asia in the Early Twenty-First Century, Cambridge: Cambridge University Press, 2014, p.39.',
                'Richard A. Posner, “The Decline of Law as an Autonomous Discipline: 1962-1987”, Harvard Law Review, Vol.100, No.4 (1987), pp.761-780.',
            ],
            'note': '英文期刊名称不要缩写。',
        },
        'foreign': {
            'name': '引用英文以外的外文文种',
            'template': '西文体例比照英文，日文体例比照中文。',
            'examples': [],
            'note': '西文（法文、德文、意大利文、西班牙文等）体例比照英文；日文体例比照中文。',
        },
    },
    'conventions': {
        'author_separator': '：',
        'multiple_authors_separator': '、',
        'editor_suffix': '主编',
        'translator_suffix': '译',
        'journal_marker': '载',
        'publisher_year_suffix': '年版',
        'edition_marker': '（第{n}版）',
        'page_prefix': '第',
        'page_suffix': '页',
        'page_range_separator': '-',
        'english_journal_abbrev': False,
    },
    'rules': {
        'punctuation': True,
        'author_separator': True,
        'zai_marker': True,
        'publisher_year': True,
        'page_format': True,
        'edition': True,           # 版次归入（第N版）
        'translator': True,        # 译者名与 "译" 之间的空格清理
        'legal_bracket': True,
        'case_bracket': True,
    },
}

# ---------------------------------------------------------------------------
# 与《法学引注手册》一致的期刊预设
# ---------------------------------------------------------------------------
def _derive_fayin(name: str, description: str) -> Dict:
    """基于《法学引注手册》默认体例派生一份同名体例 (规则/约定/类别一致)."""
    style = deepcopy(DEFAULT_CITATION_STYLE)
    style['name'] = name
    style['description'] = description
    return style


ZHONGGUO_FAXUE_STYLE = _derive_fayin(
    '《中国法学》注释体例',
    '《中国法学》的引注体例与《法学引注手册》(2019) 一致。',
)
FASHANG_YANJIU_STYLE = _derive_fayin(
    '《法商研究》注释体例',
    '《法商研究》的引注体例与《法学引注手册》(2019) 一致。',
)

# ---------------------------------------------------------------------------
# 《法学研究》注释体例 — 《法学引注手册》基础上新增 14 条要求
# ---------------------------------------------------------------------------
_FAXUE_YANJIU_NOTES = [
    '引用书籍的，要标明作者、书名、出版单位、出版年份和页码。作者为两人的，均列明姓名；为三人及以上的，标注为“××（排名首位的作者）等”。作者为机构的，标注机构名。出版单位属两家（含）及以上机构的，分别列明。',
    '书籍属多人合作作品的，可视情况标注为“××主编”、“××编”。多人分章节合作撰写的编著作品，应在注释中页码后括注“××撰写”。',
    '引用译著的，应在作者前括注作者国籍，书名后增加译者。标注顺序为：国籍、作者、书名、译者、出版单位、出版年份和页码。译著本身未标明原著作者国籍，或者未翻译原著作者姓名的，遵照译著。译者为三人或三人以上的，标注为“××等译”。',
    '引用期刊论文的，要标明作者、文章标题、期刊名及期号、页码。作者为两人的，均列明姓名；为三人及以上的，标注为“××（排名首位的作者）等”。作者为机构的，标注机构名；为课题组的，标注为“××课题组”。',
    '引用文集类书刊（含集刊）中论文的，还要按第1条的要求列明该书刊的相关要素。标注顺序为：论文作者、文章标题、书刊作者、书刊名、出版单位、出版年份和页码。其中，论文与书刊之间用“载于”连接。',
    '论文为译文的，应在论文作者前括注作者国籍，文章标题后增加译者。作者国籍不明、作者名原本未译，参照第3条酌情处理。',
    '书籍再版或多次修订的，通常应以最新版次为准，但不要标注“第×版”、“修订版”等。论文被转载、摘录的，应引用最早发表的载体。',
    '对报纸的引用，一般限于信息类、数据类引用。引用报纸上的资料，应同时注重报纸及所引内容的权威性、严肃性和专业性。引用报纸文章，要注明作者、文章标题、报纸名、日期和版面序号。作者确实不明的，可免于标注。',
    '对网络资料的引用，一般限于信息类、数据类引用，对由专业机构正式发布的电子期刊或类似网络出版物的引用，不受此限。引用网络资料，要同时注重网站及所引内容的权威性、严肃性和专业性。引用网络资料，要注明作者、文章标题、网址和最新访问日期。',
    '确需引用未公开发表的作品时，需标注作者、作品名称和页码，并视情况标明“××学校博士论文（××年）”、“××机构工作论文”或“××年印行”。',
    '书名或文章标题为若干词语之并列，且词语之间以空格相间，应视情况在相应空格位置添加顿号、逗号或者中圆点。',
    '非直接引用原文的，注释前加“参见”。非引自原始出处的，注释前加“转引自”。已公开的资料，应引用原始文献，禁用转引。',
    '数个注释引自同一出处的，注释采用“前引〔×〕，××书，第×页”或者“前引〔×〕，××文，第×页”。两个注释相邻的，采用“同上书，第×页”或者“同上文，第×页”。相邻两个注释完全相同的，采用“同上”。',
    '引文出自同一资料相邻页者，只注明首页；相邻数页者，注明“第×页以下”。',
]

FAXUE_YANJIU_STYLE = _derive_fayin(
    '《法学研究》注释体例',
    '《法学研究》的引注体例在《法学引注手册》(2019) 基础上新增 14 条注释体例要求。',
)
FAXUE_YANJIU_STYLE['notes'] = _FAXUE_YANJIU_NOTES
# 《法学研究》特定类别的格式差异
FAXUE_YANJIU_STYLE['categories']['collection']['template'] = \
    '作者：《标题》，载于编者主编：《文集》，出版社YYYY年版，第N页。'
FAXUE_YANJIU_STYLE['categories']['newspaper']['template'] = \
    '作者：《标题》，载《报纸》YYYY年M月D日，第X版。'
FAXUE_YANJIU_STYLE['categories']['web']['template'] = \
    '作者：《标题》，网址，YYYY年M月D日访问。'

# 预设体例 (名称 → 体例 dict)
PRESETS: Dict[str, Dict] = {
    '法学引注手册': DEFAULT_CITATION_STYLE,
    '《法学家》《中外法学》注释体例': JOURNAL_CITATION_STYLE,
    '《中国法学》': ZHONGGUO_FAXUE_STYLE,
    '《法商研究》': FASHANG_YANJIU_STYLE,
    '《法学研究》': FAXUE_YANJIU_STYLE,
}

# 合法字段名 (校验用)
_RULE_KEYS = frozenset(DEFAULT_CITATION_STYLE['rules'])
_CONVENTION_KEYS = frozenset(DEFAULT_CITATION_STYLE['conventions'])
_CATEGORY_FIELDS = frozenset(('name', 'template', 'examples', 'note'))
_TOP_KEYS = frozenset(('name', 'description', 'categories', 'conventions',
                       'rules', 'notes'))


def _base_style(preset: Optional[str]) -> Dict:
    """根据预设名返回基础体例 (深拷贝)."""
    if preset is None:
        return deepcopy(DEFAULT_CITATION_STYLE)
    if preset not in PRESETS:
        raise ValueError(
            f'未知预设体例: {preset!r} (支持: {", ".join(PRESETS)})')
    return deepcopy(PRESETS[preset])


def load_citation_rules(config_path: Optional[str] = None,
                        preset: Optional[str] = None) -> Dict:
    """加载注释体例: 未提供配置时返回默认体例 (法学引注手册).

    Args:
        config_path: JSON 配置文件路径, None = 用预设/默认体例.
        preset: 预设体例名 ("法学引注手册" / "《法学家》《中外法学》注释体例" /
            "《中国法学》" / "《法商研究》" / "《法学研究》"); None = 默认.

    Returns:
        完整体例 dict: {name, description, categories, conventions, rules}.
        未写出的部分用预设/默认值补齐.

    Raises:
        FileNotFoundError: 配置文件不存在.
        ValueError: 配置格式/字段非法 (带友好提示).
    """
    style = _base_style(preset)
    if config_path is None:
        return style

    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f'注释体例配置文件不存在: {config_path}')
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except json.JSONDecodeError as e:
        raise ValueError(f'注释体例配置文件 JSON 解析失败: {e}')

    if not isinstance(data, dict):
        raise ValueError(f'注释体例配置顶层必须是对象, 得到: {type(data).__name__}')

    # —— 顶层键校验 ——
    for key in data:
        if key.startswith('_'):
            continue  # "_" 开头的键视为注释
        if key not in _TOP_KEYS:
            raise ValueError(
                f'未知配置项: {key!r} (支持: name, description, '
                f'categories, conventions, rules, notes)')

    # —— name / description ——
    for key in ('name', 'description'):
        if key in data:
            val = data[key]
            if not isinstance(val, str):
                raise ValueError(f'{key!r} 必须是字符串')
            style[key] = val

    # —— notes (体例要求/注意事项) ——
    if 'notes' in data:
        notes = data['notes']
        if not isinstance(notes, list) or \
                not all(isinstance(x, str) for x in notes):
            raise ValueError('notes 必须是字符串列表')
        style['notes'] = list(notes)

    # —— rules (修复规则开关) ——
    if 'rules' in data:
        cfg = data['rules']
        if not isinstance(cfg, dict):
            raise ValueError('rules 必须是对象')
        for name, enabled in cfg.items():
            if name.startswith('_'):
                continue
            if name not in _RULE_KEYS:
                raise ValueError(
                    f'未知修复规则: {name!r} (支持: {", ".join(sorted(_RULE_KEYS))})')
            if not isinstance(enabled, bool):
                raise ValueError(f'修复规则 {name!r} 必须是 true/false')
            style['rules'][name] = enabled

    # —— conventions (命名约定) ——
    if 'conventions' in data:
        cfg = data['conventions']
        if not isinstance(cfg, dict):
            raise ValueError('conventions 必须是对象')
        for name, val in cfg.items():
            if name.startswith('_'):
                continue
            if name not in _CONVENTION_KEYS:
                raise ValueError(
                    f'未知约定: {name!r} (支持: '
                    f'{", ".join(sorted(_CONVENTION_KEYS))})')
            if name == 'english_journal_abbrev':
                if not isinstance(val, bool):
                    raise ValueError('english_journal_abbrev 必须是 true/false')
            elif not isinstance(val, str):
                raise ValueError(f'约定 {name!r} 必须是字符串')
            style['conventions'][name] = val

    # —— categories (各类文献体例) ——
    if 'categories' in data:
        cfg = data['categories']
        if not isinstance(cfg, dict):
            raise ValueError('categories 必须是对象')
        for key, cat in cfg.items():
            if key.startswith('_'):
                continue
            if key not in CATEGORY_ORDER:
                raise ValueError(
                    f'未知文献类别: {key!r} (支持: {", ".join(CATEGORY_ORDER)})')
            if not isinstance(cat, dict):
                raise ValueError(f'类别 {key!r} 的配置必须是对象')
            for field, val in cat.items():
                if field not in _CATEGORY_FIELDS:
                    raise ValueError(
                        f'类别 {key!r} 未知字段: {field!r} '
                        f'(支持: name, template, examples, note)')
                if field == 'examples':
                    if not isinstance(val, list) or \
                            not all(isinstance(x, str) for x in val):
                        raise ValueError(f'类别 {key!r} 的 examples 必须是字符串列表')
                    style['categories'][key]['examples'] = list(val)
                elif field == 'note':
                    if not isinstance(val, str):
                        raise ValueError(f'类别 {key!r} 的 note 必须是字符串')
                    style['categories'][key]['note'] = val
                else:
                    if not isinstance(val, str):
                        raise ValueError(f'类别 {key!r} 的 {field!r} 必须是字符串')
                    style['categories'][key][field] = val

    return style


def describe_style(style: Dict) -> str:
    """生成体例的可读说明 (文本/markdown 片段)."""
    lines: List[str] = []
    lines.append(f"注释体例：{style.get('name', '')}")
    desc = style.get('description', '')
    if desc:
        lines.append(f"说明：{desc}")
    lines.append('')

    lines.append('## 各类文献体例')
    lines.append('')
    for idx, key in enumerate(CATEGORY_ORDER, 1):
        cat = style['categories'][key]
        lines.append(f"{idx}. {cat['name']}")
        tmpl = cat.get('template', '')
        if tmpl:
            lines.append(f"   格式：{tmpl}")
        for ex in cat.get('examples', []):
            lines.append(f"   例：{ex}")
        note = cat.get('note', '')
        if note:
            lines.append(f"   注：{note}")
        lines.append('')

    notes = style.get('notes', [])
    if notes:
        lines.append('## 注释体例要求')
        lines.append('')
        for i, n in enumerate(notes, 1):
            lines.append(f"{i}. {n}")
        lines.append('')
        lines.append('注：以上为说明性要求 (供人工核对)，当前不驱动自动修复；')
        lines.append('自动执行的范围以「启用的修复规则」为准。')
        lines.append('')

    lines.append('## 命名约定')
    lines.append('')
    for name, val in style['conventions'].items():
        lines.append(f"  {name} = {val!r}")
    lines.append('')

    lines.append('## 启用的修复规则')
    lines.append('')
    enabled = [n for n, v in style['rules'].items() if v]
    disabled = [n for n, v in style['rules'].items() if not v]
    lines.append(f"  启用：{', '.join(enabled) if enabled else '(无)'}")
    if disabled:
        lines.append(f"  关闭：{', '.join(disabled)}")
    return '\n'.join(lines)


if __name__ == '__main__':
    import sys
    import argparse
    ap = argparse.ArgumentParser(description='校验并打印注释体例')
    ap.add_argument('config', nargs='?', default=None,
                    help='注释体例 JSON 配置文件 (可选)')
    ap.add_argument('--preset', default=None,
                    help=f'预设体例: {", ".join(PRESETS)}')
    args = ap.parse_args()

    if args.config is None and args.preset is None:
        print('用法: python citation_rules.py [配置.json] [--preset 预设名]')
        print(f'  预设: {", ".join(PRESETS)}')
        print('  无参数时打印默认体例 (法学引注手册)。')
        sys.exit(1)

    style = load_citation_rules(args.config, preset=args.preset)
    print(describe_style(style))
