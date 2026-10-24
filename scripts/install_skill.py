# -*- coding: utf-8 -*-
"""
install_skill.py — 把本仓库的技能安装到本机各个 harness 的技能目录。

用法:
    python scripts/install_skill.py                 # 安装到检测到的所有 harness
    python scripts/install_skill.py --list          # 只看会装到哪里
    python scripts/install_skill.py --targets claude,dsh
    python scripts/install_skill.py --targets all

目标目录（用户级）:
    workbuddy  ~/.workbuddy/skills/<name>/      WorkBuddy
    claude     ~/.claude/skills/<name>/         Claude Code
    dsh        ~/.dsh/skills/<name>/            DeepSeek Harness

只复制 SKILL.md、scripts/、references/、assets/，不带 README、.gitignore 与 __pycache__。
"""
import argparse
import os
import shutil
import sys
from pathlib import Path

SKILL_NAME = "gszy-thesis-format"
REPO = Path(__file__).resolve().parent.parent

TARGETS = {
    "workbuddy": Path.home() / ".workbuddy" / "skills",
    "claude": Path.home() / ".claude" / "skills",
    "dsh": Path.home() / ".dsh" / "skills",
}

# 各 harness 对 frontmatter 的要求不同：
# WorkBuddy 需要 agent_created: true 才允许后续修改/删除；Claude Code 与 DSH 只认 name/description，
# 保留额外键在 Claude Code 里会被静默忽略，但上传到 claude.ai / Skills API 时会硬报错，所以剥掉。
STRIP_KEYS = {"claude": ["agent_created"], "dsh": ["agent_created"], "workbuddy": []}

COPY_ITEMS = ["SKILL.md", "scripts", "references", "assets"]
IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store")


def render_skill_md(target: str) -> str:
    text = (REPO / "SKILL.md").read_text(encoding="utf-8")
    drop = STRIP_KEYS.get(target, [])
    if not drop:
        return text
    lines = text.splitlines(keepends=True)
    if not lines or lines[0].strip() != "---":
        return text
    end = next((i for i in range(1, len(lines)) if lines[i].strip() == "---"), None)
    if end is None:
        return text
    head = [lines[0]]
    for line in lines[1:end]:
        key = line.split(":", 1)[0].strip() if ":" in line else ""
        if key in drop:
            continue
        head.append(line)
    return "".join(head + lines[end:])


def install(target: str, base: Path, dry_run: bool = False) -> None:
    dst = base / SKILL_NAME
    print(f"[{target}] {dst}")
    if dry_run:
        return
    dst.mkdir(parents=True, exist_ok=True)
    for item in COPY_ITEMS:
        src = REPO / item
        if not src.exists():
            print(f"  ! 源缺失 {item}")
            continue
        d = dst / item
        if src.is_dir():
            shutil.rmtree(d, ignore_errors=True)
            shutil.copytree(src, d, ignore=IGNORE)
        else:
            shutil.copy2(src, d)
    (dst / "SKILL.md").write_text(render_skill_md(target), encoding="utf-8")
    n = sum(1 for _ in dst.rglob("*") if _.is_file())
    print(f"  ✓ {n} 个文件")


def main():
    ap = argparse.ArgumentParser(description="安装 gszy-thesis-format 技能到本机 harness")
    ap.add_argument("--targets", default="auto",
                    help="auto（默认，仅已存在的 harness）| all | workbuddy,claude,dsh")
    ap.add_argument("--list", action="store_true", help="只打印安装位置")
    args = ap.parse_args()

    if not (REPO / "SKILL.md").exists():
        sys.exit(f"找不到 {REPO / 'SKILL.md'}，请在仓库根目录内运行")

    if args.targets in ("auto", "all"):
        chosen = list(TARGETS)
        if args.targets == "auto":
            chosen = [t for t, p in TARGETS.items() if p.parent.exists()]
    else:
        chosen = [t.strip() for t in args.targets.split(",") if t.strip()]

    unknown = [t for t in chosen if t not in TARGETS]
    if unknown:
        sys.exit(f"未知目标：{unknown}，可选 {list(TARGETS)}")
    if not chosen:
        sys.exit("没有检测到任何 harness 目录，用 --targets all 强制安装")

    for t in chosen:
        install(t, TARGETS[t], dry_run=args.list)


if __name__ == "__main__":
    main()
