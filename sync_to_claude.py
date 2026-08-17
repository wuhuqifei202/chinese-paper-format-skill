#!/usr/bin/env python3
"""单向同步 skill 到 Claude Code 安装目录 (主仓库 → 部署目录).

主仓库: 本脚本所在目录 (~/.workbuddy/skills/chinese-paper-format-skill)
部署:   ~/.claude/skills/chinese-paper-format-skill

用法:
    python sync_to_claude.py            # dry-run: 打印同步计划, 不改动
    python sync_to_claude.py --apply    # 执行同步
    python sync_to_claude.py --apply --quiet   # 无输出 (供 post-commit hook 调用)

设计原则:
    - 单向: 主仓库是唯一事实源, 部署目录只被覆盖. 禁止在部署目录直接修改.
    - 删除策略: 部署目录中源没有的文件会被删除 (排除项除外),
      保证部署目录与主仓库完全一致.
    - 排除项: .git*, __pycache__, .pytest_cache, *.old, *.pyc, logs/, 本脚本自身.
"""

import argparse
import filecmp
import os
import shutil
import sys
from pathlib import Path

# 主仓库根目录 = 本脚本所在目录
SOURCE = Path(os.path.dirname(os.path.abspath(__file__)))
TARGET = Path.home() / '.claude' / 'skills' / SOURCE.name

# 排除规则: 目录名 / 文件名 / 文件名后缀 (注意 .gitignore 不处理, 这里显式列出)
EXCLUDED_DIR_NAMES = {'.git', '__pycache__', '.pytest_cache', 'logs'}
EXCLUDED_NAME_PREFIXES = {'.git'}
EXCLUDED_SUFFIXES = {'.old', '.pyc', '.bak'}
EXCLUDED_FILES = {'sync_to_claude.py'}


def is_excluded(rel: str, is_dir: bool) -> bool:
    parts = rel.split(os.sep)
    for p in parts:
        if is_dir and p in EXCLUDED_DIR_NAMES:
            return True
        if p.startswith(tuple(EXCLUDED_NAME_PREFIXES)):
            return True
    name = parts[-1]
    if name in EXCLUDED_FILES:
        return True
    if name.endswith(tuple(EXCLUDED_SUFFIXES)):
        return True
    return False


def collect_files(root: Path):
    """返回源树相对路径列表 (已排除), 仅文件."""
    files = []
    for dirpath, dirnames, filenames in os.walk(root):
        rel_dir = os.path.relpath(dirpath, root)
        # 原地过滤目录 (os.walk 剪枝, 避免遍历排除目录)
        dirnames[:] = [d for d in dirnames
                       if not is_excluded(os.path.join(rel_dir, d) if rel_dir != '.' else d, True)]
        for f in filenames:
            rel = os.path.join(rel_dir, f) if rel_dir != '.' else f
            if not is_excluded(rel, False):
                files.append(rel)
    return sorted(files)


def sync(dry_run: bool, quiet: bool):
    if not TARGET.exists():
        print(f'错误: 部署目录不存在: {TARGET}', file=sys.stderr)
        print(f'提示: 请先运行 install.sh 安装到 .claude, 或手动创建该目录。',
              file=sys.stderr)
        sys.exit(1)

    source_files = collect_files(SOURCE)
    target_files = collect_files(TARGET)

    plan = []  # (action, rel)
    # 复制/更新
    for rel in source_files:
        src = SOURCE / rel
        dst = TARGET / rel
        if not dst.exists():
            plan.append(('+', rel))
        elif not filecmp.cmp(src, dst, shallow=False):
            plan.append(('~', rel))
    # 删除目标中多余文件
    for rel in target_files:
        if rel not in source_files:
            plan.append(('-', rel))

    if dry_run:
        print(f'同步计划 (dry-run): {len(plan)} 项')
        for action, rel in plan:
            print(f'  {action} {rel}')
        print(f'\n源:     {SOURCE}')
        print(f'目标:   {TARGET}')
        return

    if not plan:
        if not quiet:
            print(f'已是最新: {TARGET}')
        return

    n = 0
    for action, rel in plan:
        src = SOURCE / rel
        dst = TARGET / rel
        if action == '-':
            if not quiet:
                print(f'删除 {rel}')
            if dst.is_dir():
                shutil.rmtree(dst)
            else:
                dst.unlink()
        else:
            if not quiet:
                print(f'{"复制" if action == "+" else "更新"} {rel}')
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
        n += 1
    if not quiet:
        print(f'同步完成: {n} 项 → {TARGET}')


def main():
    ap = argparse.ArgumentParser(description='同步 skill 主仓库 → Claude 部署目录')
    ap.add_argument('--apply', action='store_true',
                    help='执行同步 (默认仅打印计划)')
    ap.add_argument('--quiet', action='store_true',
                    help='无输出 (与 --apply 组合, 供 post-commit hook)')
    args = ap.parse_args()
    sync(dry_run=not args.apply, quiet=args.quiet)


if __name__ == '__main__':
    main()
