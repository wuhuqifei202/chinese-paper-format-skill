#!/usr/bin/env python3
"""单向同步 skill 到 WorkBuddy 备用目录 (主仓库 → 实验沙盒).

主仓库: 本脚本所在目录 (~/.claude/skills/chinese-paper-format-skill, git, 唯一事实源)
备用:   ~/.workbuddy/skills/chinese-paper-format-skill (无 git, 实验沙盒)

用法:
    python sync_to_workbuddy.py            # dry-run: 打印同步计划, 不改动
    python sync_to_workbuddy.py --apply    # 执行同步 (保护本地修改/新增)
    python sync_to_workbuddy.py --force    # 忽略保护, 完全对齐 (与 --apply 组合)
    python sync_to_workbuddy.py --apply --quiet   # 无输出 (供 post-commit hook 调用)

冲突保护 (实验沙盒语义):
    备用目录允许单独修改与测试, 不直接影响主仓库 (同步永远是主仓库 → 备用).
    同步时:
      - 备用中「比主仓库新」的文件 (本地修改过) → 跳过, 警告, 不覆盖
      - 备用中「主仓库没有」的文件 (本地新增, 实验产物) → 跳过, 警告, 不删除
      - 其余差异 (主仓库后改/新增) → 正常覆盖/复制
    判定依据: mtime (shutil.copy2 保留 mtime, 故「后改」即「本地改」).
    需要完全对齐时用 --force (覆盖/删除全部差异).
"""

import argparse
import filecmp
import os
import shutil
import sys
from pathlib import Path

# 主仓库根目录 = 本脚本所在目录
SOURCE = Path(os.path.dirname(os.path.abspath(__file__)))
TARGET = Path.home() / '.workbuddy' / 'skills' / SOURCE.name

# mtime 比较容差 (秒): 两文件 mtime 差在此范围内视为同时
_MTIME_EPS = 0.5

# 排除规则: 目录名 / 文件名 / 文件名后缀
EXCLUDED_DIR_NAMES = {'.git', '__pycache__', '.pytest_cache', 'logs'}
EXCLUDED_NAME_PREFIXES = {'.git'}
EXCLUDED_SUFFIXES = {'.old', '.pyc', '.bak'}
EXCLUDED_FILES = {'sync_to_workbuddy.py'}


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
    """返回目录树相对路径列表 (已排除), 仅文件."""
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


def is_locally_modified(src: Path, dst: Path) -> bool:
    """备用目录文件是否被本地修改过 (备用 mtime 比主仓库新)."""
    return dst.stat().st_mtime > src.stat().st_mtime + _MTIME_EPS


def sync(dry_run: bool, quiet: bool, force: bool):
    if not TARGET.exists():
        print(f'错误: 备用目录不存在: {TARGET}', file=sys.stderr)
        print(f'提示: 请先运行 install.sh 安装到 workbuddy, 或手动创建该目录。',
              file=sys.stderr)
        sys.exit(1)

    source_files = collect_files(SOURCE)
    target_files = collect_files(TARGET)

    plan = []  # (action, rel)  action: + 复制, ~ 覆盖, - 删除, P~ 保护(跳过覆盖), P- 保护(跳过删除)
    # 复制/更新
    for rel in source_files:
        src = SOURCE / rel
        dst = TARGET / rel
        if not dst.exists():
            plan.append(('+', rel))
        elif not filecmp.cmp(src, dst, shallow=False):
            if not force and is_locally_modified(src, dst):
                plan.append(('P~', rel))  # 备用后改 → 本地实验, 保护
            else:
                plan.append(('~', rel))
    # 删除目标中多余文件 (本地新增的实验产物默认保护)
    for rel in target_files:
        if rel not in source_files:
            if force:
                plan.append(('-', rel))
            else:
                plan.append(('P-', rel))

    if dry_run:
        print(f'同步计划 (dry-run): {len(plan)} 项'
              + (', --force 完全对齐' if force else ', 本地修改/新增已保护'))
        for action, rel in plan:
            if action == 'P~':
                print(f'  ⚠ 保护(本地修改, 跳过覆盖) {rel}')
            elif action == 'P-':
                print(f'  ⚠ 保护(本地新增, 跳过删除) {rel}')
            else:
                print(f'  {action} {rel}')
        print(f'\n源:     {SOURCE} (主仓库)')
        print(f'目标:   {TARGET} (实验沙盒)')
        if not force:
            print('提示: 用 --force 忽略保护, 完全对齐')
        return

    if not plan:
        if not quiet:
            print(f'已是最新: {TARGET}')
        return

    n = 0
    protected = 0
    for action, rel in plan:
        src = SOURCE / rel
        dst = TARGET / rel
        if action in ('P~', 'P-'):
            protected += 1
            if not quiet:
                hint = '本地修改, 跳过覆盖' if action == 'P~' else '本地新增, 跳过删除'
                print(f'跳过({hint}): {rel}')
            continue
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
        tail = f' (保护 {protected} 项本地改动)' if protected else ''
        print(f'同步完成: {n} 项 → {TARGET}{tail}')


def main():
    ap = argparse.ArgumentParser(description='同步 skill 主仓库 → WorkBuddy 实验沙盒')
    ap.add_argument('--apply', action='store_true',
                    help='执行同步 (默认仅打印计划)')
    ap.add_argument('--force', action='store_true',
                    help='忽略冲突保护, 完全对齐 (含本地修改/新增)')
    ap.add_argument('--quiet', action='store_true',
                    help='无输出 (与 --apply 组合, 供 post-commit hook)')
    args = ap.parse_args()
    sync(dry_run=not args.apply, quiet=args.quiet, force=args.force)


if __name__ == '__main__':
    main()
