#!/usr/bin/env python3
"""rules — 自定义格式规则表 (JSON 配置文件加载).

让使用者按自己的格式要求便捷转换格式, 不限于 skill 预设的一种规则.

配置文件 (JSON, 键/值见下方字段说明), 只写要改的元素即可,
未写出的元素/字段使用默认规则 (DEFAULT_RULES, 与 skill 预设一致):

    {
        "title": {"size": "二号", "bold": true},
        "h1":    {"font": "宋体", "size": "四号", "bold": true},
        "body":  {"size": "小四", "line_spacing": "20磅"}
    }

支持的元素 (键):
    title     论文题目
    author    作者行 / 副标题 (题目区域非首段非元数据)
    abstract  摘要
    keywords  关键词
    h1 h2 h3 h4   一/二/三/四级标题
    body      正文
    footnote  脚注

字段 (值):
    font          中文字体名, 如 "黑体" "宋体" "楷体" "仿宋"
    size          中文字号名 ("二号" "小四" "五号"…) 或磅值数字 (22) 或 "22pt"
    bold          true / false
    align         "center" "left" "right" "justify" (或 "居中" "左对齐" …)
    indent        首行缩进字符数 (0 = 无缩进)
    line_spacing  行距: 数字 = 倍数 (1.0 / 1.5), "20磅" / "20pt" = 固定磅值,
                  "单倍" = 1.0

用法:
    from rules import DEFAULT_RULES, load_rules
    rules = load_rules('我的格式.json')   # 完整规则 dict, 字段齐全
"""

import json
import re
from pathlib import Path
from typing import Dict, Optional

# ---------------------------------------------------------------------------
# 中文字号 → 磅值 (Pt)
# 标准中文字号表 (来源: 中文排版字号规范)
# ---------------------------------------------------------------------------
CN_SIZE_PT = {
    '初号': 42, '小初': 36,
    '一号': 26, '小一': 24,
    '二号': 22, '小二': 18,
    '三号': 16, '小三': 15,
    '四号': 14, '小四': 12,
    '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5,
    '七号': 5.5, '八号': 5,
}

# 对齐方式: 配置值 (英文/中文) → 内部值
ALIGN_ALIASES = {
    'center': 'center', '居中': 'center',
    'left': 'left', '左对齐': 'left', '靠左': 'left',
    'right': 'right', '右对齐': 'right',
    'justify': 'justify', '两端对齐': 'justify',
}

# 可配置的元素
ELEMENTS = ('title', 'author', 'abstract', 'keywords',
            'h1', 'h2', 'h3', 'h4', 'body', 'footnote')

# 默认规则 (与 skill 预设格式一致: format_paper.py 常量)
DEFAULT_RULES: Dict[str, dict] = {
    'title':    {'font': '黑体', 'size': 14,   'bold': False,
                 'align': 'center', 'indent': 0, 'line_spacing': 1.0},
    'author':   {'font': '黑体', 'size': 14,   'bold': False,
                 'align': 'center', 'indent': 0, 'line_spacing': 1.0},
    'abstract': {'font': '楷体', 'size': 12,   'bold': False,
                 'align': 'left', 'indent': 0, 'line_spacing': 1.0},
    'keywords': {'font': '楷体', 'size': 12,   'bold': False,
                 'align': 'left', 'indent': 0, 'line_spacing': 1.0},
    'h1':       {'font': '宋体', 'size': 12,   'bold': True,
                 'align': 'center', 'indent': 0, 'line_spacing': 1.0},
    'h2':       {'font': '楷体', 'size': 12,   'bold': True,
                 'align': 'left', 'indent': 2, 'line_spacing': 1.0},
    'h3':       {'font': '宋体', 'size': 10.5, 'bold': True,
                 'align': 'left', 'indent': 2, 'line_spacing': 1.0},
    'h4':       {'font': '宋体', 'size': 10.5, 'bold': True,
                 'align': 'left', 'indent': 2, 'line_spacing': 1.0},
    'body':     {'font': '宋体', 'size': 10.5, 'bold': False,
                 'align': 'left', 'indent': None, 'line_spacing': 1.0},
    'footnote': {'font': '宋体', 'size': 9,    'bold': False,
                 'align': 'left', 'indent': 0, 'line_spacing': 1.0},
}

# body.indent: None = 由运行时 --body-indent 参数决定 (默认 2)

_RE_PT = re.compile(r'^([0-9]+(?:\.[0-9]+)?)\s*(?:pt|磅|磅值)$', re.I)


def parse_size(value) -> float:
    """解析字号值 → 磅值. 支持: 中文字号名 / 数字 / "22pt" "22磅"."""
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        raise ValueError(f'字号必须为数字或中文字号名, 得到: {value!r}')
    s = value.strip()
    if s in CN_SIZE_PT:
        return CN_SIZE_PT[s]
    m = _RE_PT.match(s)
    if m:
        return float(m.group(1))
    raise ValueError(f'未知字号: {value!r} (支持中文字号名如 "二号"/"小四", '
                     f'或磅值如 22 / "22pt")')


def parse_line_spacing(value):
    """解析行距值.
    数字 → 倍数 (float); "20磅"/"20pt" → {'mode': 'exact', 'value': 20}.
    """
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        raise ValueError(f'行距必须为数字(倍数)或 "N磅", 得到: {value!r}')
    s = value.strip()
    if s == '单倍':
        return 1.0
    m = _RE_PT.match(s)
    if m:
        return {'mode': 'exact', 'value': float(m.group(1))}
    try:
        return float(s)  # 允许 "1.5" 字符串
    except ValueError:
        raise ValueError(f'未知行距: {value!r} (数字=倍数, "20磅"/"20pt"=固定磅值)')


def load_rules(config_path: Optional[str] = None) -> Dict[str, dict]:
    """加载格式规则: 未提供配置时返回默认规则.

    Args:
        config_path: JSON 配置文件路径, None = 默认规则.

    Returns:
        完整规则 dict: {元素: {font, size, bold, align, indent, line_spacing}}.
        所有字段齐全 (未写出的用默认值).

    Raises:
        FileNotFoundError: 配置文件不存在.
        ValueError: 配置文件格式/字段非法.
    """
    if config_path is None:
        return {k: dict(v) for k, v in DEFAULT_RULES.items()}

    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f'规则配置文件不存在: {config_path}')
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except json.JSONDecodeError as e:
        raise ValueError(f'规则配置文件 JSON 解析失败: {e}')

    if not isinstance(data, dict):
        raise ValueError(f'规则配置文件顶层必须是对象, 得到: {type(data).__name__}')

    rules = {k: dict(v) for k, v in DEFAULT_RULES.items()}
    for elem, cfg in data.items():
        if elem.startswith('_'):  # "_" 开头的键视为注释, 允许写在配置里
            continue
        if elem not in ELEMENTS:
            raise ValueError(
                f'未知元素: {elem!r} (支持: {", ".join(ELEMENTS)})')
        if not isinstance(cfg, dict):
            raise ValueError(f'元素 {elem!r} 的配置必须是对象')
        for field, value in cfg.items():
            if field not in DEFAULT_RULES[elem]:
                raise ValueError(
                    f'元素 {elem!r} 未知字段: {field!r} '
                    f'(支持: font, size, bold, align, indent, line_spacing)')
            rules[elem][field] = value

    # —— 字段解析与校验 ——
    for elem, rule in rules.items():
        try:
            rule['size'] = parse_size(rule['size'])
            rule['line_spacing'] = parse_line_spacing(rule['line_spacing'])
        except ValueError as e:
            raise ValueError(f'元素 {elem!r}: {e}')
        align = rule['align']
        if align not in ALIGN_ALIASES:
            raise ValueError(f'元素 {elem!r}: 未知对齐方式 {align!r} '
                             f'(支持: center/left/right/justify 或中文别名)')
        rule['align'] = ALIGN_ALIASES[align]
        if not isinstance(rule['bold'], bool):
            raise ValueError(f'元素 {elem!r}: bold 必须为 true/false')
        if rule['indent'] is not None and not isinstance(rule['indent'], int):
            raise ValueError(f'元素 {elem!r}: indent 必须为整数 (字符数)')

    return rules


if __name__ == '__main__':
    import sys
    # CLI: 校验并打印完整规则
    if len(sys.argv) < 2:
        print('用法: python rules.py 配置.json [--show]')
        sys.exit(1)
    rules = load_rules(sys.argv[1])
    for elem, rule in rules.items():
        print(f'{elem:10s} {rule}')
