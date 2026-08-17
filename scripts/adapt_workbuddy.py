# -*- coding: utf-8 -*-
"""为 workbuddy 平台生成适配版 skill (SKILL.md 格式转换 + _skillhub_meta.json).

workbuddy (~/.workbuddy) 的 skill frontmatter 与 Claude Agent Skills
标准不同: 需要 description_zh/description_en、顶层 version、
metadata.clawdbot (emoji/requires/install)、display_name/visibility 等
字段, 并配套 _skillhub_meta.json 元数据。
参考本机已安装的 workbuddy skill 格式 (obsidian/wechat-article-search)。

用法:
    python adapt_workbuddy.py <source_dir> <target_dir>

source_dir: skill 根目录 (含标准 SKILL.md)
target_dir: 目标 skills 目录 (如 ~/.workbuddy/skills/chinese-paper-format-skill)

生成:
  - SKILL.md: frontmatter 转为 workbuddy 格式, 主体内容保留
  - _skillhub_meta.json: 安装元数据
其他文件 (scripts/references/assets/tests/README/AGENTS/install.sh) 原样复制。
"""

import json
import re
import shutil
import sys
import time
from pathlib import Path

# workbuddy 平台展示信息
DISPLAY_NAME_ZH = '中文学术论文格式规范化'
DISPLAY_NAME_EN = 'chinese-paper-format-skill'
EMOJI = '📄'
DESC_ZH = ('中文学术论文格式规范化: docx 排版(题目/一~四级标题/正文/脚注)、'
           '引注修复、自定义格式规则、markdown 中转')
DESC_EN = 'Format Chinese academic papers (docx) per academic conventions'
EXAMPLES_ZH = [
    '把这篇论文格式化成学术规范',
    '统一调整论文排版',
    '批量格式化文件夹里的论文',
]
EXAMPLES_EN = [
    'Format this paper to academic style',
    'Unify paper formatting',
    'Batch format papers in a folder',
]


def read_skill_md(skill_md: Path) -> dict:
    """轻量解析标准 SKILL.md frontmatter, 提取 name/description/version.

    不依赖 PyYAML: 只处理本脚本需要的字段。
    - "key: value" 单行值
    - "key: >-" 多行折叠块 (description, 缩进行以空格拼接)
    - "key:" 嵌套块起始 (metadata/provenance), 其缩进子键被收集为
      "key 内容" 字符串, 供 version 等子字段正则提取
    """
    raw = skill_md.read_text(encoding='utf-8')
    if not raw.startswith('---'):
        raise ValueError(f'{skill_md} 必须以 --- 开头')
    lines = raw.splitlines()
    end = next(i for i, l in enumerate(lines[1:], start=1)
               if l.strip() == '---')
    fm_lines = lines[1:end]
    body = '\n'.join(lines[end + 1:])

    fields = {}
    key = None
    for line in fm_lines:
        m = re.match(r'^([A-Za-z_-]+):\s*(.*)$', line)
        if m:
            key, value = m.group(1), m.group(2).strip()
            if value in ('>-', '|'):
                fields[key] = []
            elif value:
                fields[key] = value
            else:
                fields[key] = []  # 嵌套块起始: 收集缩进子键行
            continue
        if key and line.startswith('  ') and isinstance(fields.get(key), list):
            fields[key].append(line.strip())

    desc = ''
    if isinstance(fields.get('description'), list):
        desc = ' '.join(fields['description']).strip()
    elif isinstance(fields.get('description'), str):
        desc = fields['description'].strip()

    # 嵌套块内提取 version (metadata 块中的 "version: x.y.z")
    version = '1.0.0'
    for k in ('metadata', 'provenance'):
        block = fields.get(k, [])
        if isinstance(block, list):
            m = re.search(r'^version:\s*([\d.]+)', '\n'.join(block), re.M)
            if m:
                version = m.group(1)
                break
    fields['description_text'] = desc or '将中文论文 doc/docx 按学术排版规范自动格式化'
    fields['version'] = version
    fields['body'] = body
    return fields


def build_workbuddy_skill_md(fields: dict) -> str:
    """构造 workbuddy 版 SKILL.md 完整内容 (frontmatter 转换 + 主体保留)."""
    name = fields.get('name', 'chinese-paper-format-skill')
    desc = fields['description_text'].replace('"', "'")
    fm = (
        '---\n'
        f'name: {name}\n'
        f'description: "{desc}"\n'
        f'description_zh: "{DESC_ZH}"\n'
        f'description_en: "{DESC_EN}"\n'
        f'version: {fields["version"]}\n'
        'allowed-tools: Bash,Read\n'
        'metadata:\n'
        '  clawdbot:\n'
        f'    emoji: "{EMOJI}"\n'
        f'display_name: "{DISPLAY_NAME_ZH}"\n'
        f'display_name_en: "{DISPLAY_NAME_EN}"\n'
        'visibility: "public"\n'
        '---\n\n'
    )
    return fm + fields['body'].lstrip('\n')


def _copy2_ignore_locked(src, dst):
    """复制单个文件, 被占用/锁定文件跳过 (目标已有相同内容时无碍)."""
    try:
        shutil.copy2(src, dst)
    except (PermissionError, OSError):
        pass


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    src = Path(sys.argv[1]).resolve()
    dst = Path(sys.argv[2]).resolve()
    dst.mkdir(parents=True, exist_ok=True)

    skill_md = src / 'SKILL.md'
    if not skill_md.exists():
        print(f'错误: {skill_md} 不存在', file=sys.stderr)
        sys.exit(1)

    fields = read_skill_md(skill_md)
    name = fields.get('name', 'chinese-paper-format-skill')
    version = fields['version']

    # 1. workbuddy 版 SKILL.md
    (dst / 'SKILL.md').write_text(
        build_workbuddy_skill_md(fields), encoding='utf-8')
    print(f'已生成: {dst / "SKILL.md"} (workbuddy 格式, v{version})')

    # 2. _skillhub_meta.json
    meta = {
        'name': name,
        'installedAt': int(time.time() * 1000),
        'source': 'local',
        'version': version,
        'skillId': f'local_{name}',
        'examples_zh': EXAMPLES_ZH,
        'examples_en': EXAMPLES_EN,
        'description_zh': DESC_ZH,
        'description_en': DESC_EN,
    }
    (dst / '_skillhub_meta.json').write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'已生成: {dst / "_skillhub_meta.json"}')

    # 3. 其余文件 (scripts/references/assets/tests/README/AGENTS/install.sh)
    for item in src.iterdir():
        if item.name == 'SKILL.md':
            continue
        target = dst / item.name
        if item.is_dir():
            shutil.copytree(item, target, dirs_exist_ok=True,
                            copy_function=_copy2_ignore_locked)
        else:
            _copy2_ignore_locked(item, target)
    print(f'其余文件已复制到 {dst}')


if __name__ == '__main__':
    main()
