#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tree_print.py — הדפסת מבנה פרויקט (תיקיות וקבצים) בסגנון tree
שימוש:
  python tree_print.py [נתיב] [--max-depth N] [--exclude "*.pyc" --exclude ".git*"] [--dirs-first] [--show-sizes] [--hidden]
"""
from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path

BRANCH_MID = "├── "
BRANCH_END = "└── "
PIPE       = "│   "
INDENT     = "    "

def human_size(n: int) -> str:
    for unit in ("B","KB","MB","GB","TB","PB"):
        if n < 1024:
            return f"{n}{unit}"
        n //= 1024
    return f"{n}PB"

def should_exclude(name: str, patterns: list[str]) -> bool:
    return any(fnmatch.fnmatch(name, pat) for pat in patterns)

def list_children(p: Path, show_hidden: bool, excludes: list[str], dirs_first: bool):
    items = [c for c in p.iterdir() if show_hidden or not c.name.startswith(".")]
    if excludes:
        items = [c for c in items if not should_exclude(c.name, excludes)]
    if dirs_first:
        items.sort(key=lambda x: (x.is_file(), x.name.lower()))
    else:
        items.sort(key=lambda x: x.name.lower())
    return items

def print_tree(root: Path,
               prefix: str = "",
               depth: int | None = None,
               excludes: list[str] | None = None,
               dirs_first: bool = True,
               show_sizes: bool = False,
               show_hidden: bool = False,
               counters: dict | None = None):
    if counters is None:
        counters = {"dirs": 0, "files": 0}

    children = list_children(root, show_hidden, excludes or [], dirs_first)
    for i, child in enumerate(children):
        connector = BRANCH_END if i == len(children)-1 else BRANCH_MID
        label = child.name
        if show_sizes and child.is_file():
            try:
                size = human_size(child.stat().st_size)
                label += f" ({size})"
            except OSError:
                pass
        print(prefix + connector + label)
        if child.is_dir():
            counters["dirs"] += 1
            if depth is None or depth > 1:
                new_prefix = prefix + (INDENT if i == len(children)-1 else PIPE)
                print_tree(child,
                           prefix=new_prefix,
                           depth=None if depth is None else depth-1,
                           excludes=excludes,
                           dirs_first=dirs_first,
                           show_sizes=show_sizes,
                           show_hidden=show_hidden,
                           counters=counters)
        else:
            counters["files"] += 1
    return counters

def main():
    ap = argparse.ArgumentParser(description="הדפסת מבנה תיקיות וקבצים בסגנון tree")
    ap.add_argument("path", nargs="?", default=".", help="נתיב הבסיס (ברירת מחדל: הנוכחי)")
    ap.add_argument("--max-depth", type=int, default=None, help="עומק מקסימלי לסריקה (ללא הגבלה כברירת מחדל)")
    ap.add_argument("--exclude", action="append", default=[], help="תבנית לסינון (ניתן לציין מספר פעמים), לדוגמה: --exclude '*.pyc' --exclude '.git'")
    ap.add_argument("--dirs-first", action="store_true", help="סדר תיקיות לפני קבצים")
    ap.add_argument("--show-sizes", action="store_true", help="הצגת גדלי קבצים")
    ap.add_argument("--hidden", action="store_true", help="כולל קבצים/תיקיות חבויים (מתחילים בנקודה)")
    args = ap.parse_args()

    root = Path(args.path).resolve()
    if not root.exists():
        print(f"שגיאה: הנתיב לא קיים: {root}")
        return

    print(root.name + "/")
    counters = print_tree(root,
                          prefix="",
                          depth=args.max_depth,
                          excludes=args.exclude,
                          dirs_first=args.dirs_first,
                          show_sizes=args.show_sizes,
                          show_hidden=args.hidden)
    print(f"\n{counters['dirs']} directories, {counters['files']} files")

if __name__ == "__main__":
    main()
