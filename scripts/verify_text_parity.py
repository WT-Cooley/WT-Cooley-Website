#!/usr/bin/env python3
"""Verify that a converted/hand-assembled page preserves the source's wording.

Extracts normalized plain text from the raw Word-export source and from the
final target page's <article> (or <main>, or a custom selector), then diffs
them word-by-word. Any non-empty diff is a hard stop -- investigate before
treating the conversion as done.

Known-intentional drops (old-site nav chrome that leaked into a Word export,
stray "Return to Home Page" links, etc.) must be passed explicitly via
--exclude so they're visible in the run's output, not silently absorbed.

Usage:
    python3 verify_text_parity.py "path/to/Source.html" writings/target.html \
        [--encoding utf-16le] [--target-selector article] \
        [--exclude "Return to Home Page"]
"""
import argparse
import difflib
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).parent))
from convert_word_html import (  # noqa: E402
    decompose_bullet_glyphs,
    find_root,
    load_source,
    strip_conditional_blocks,
)


def normalize(text: str) -> list[str]:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text.split(" ")


def source_words(path: Path, encoding: str, excludes: list[str]) -> list[str]:
    text = load_source(path, encoding)
    text = strip_conditional_blocks(text)
    soup = BeautifulSoup(text, "html.parser")
    root = find_root(soup)
    decompose_bullet_glyphs(root)
    raw = root.get_text()
    for pattern in excludes:
        raw = re.sub(pattern, " ", raw)
    return normalize(raw)


def target_words(path: Path, selector: str) -> list[str]:
    soup = BeautifulSoup(path.read_text(encoding="utf-8"), "html.parser")
    node = soup.find(selector)
    if node is None:
        node = soup.find("main")
    if node is None:
        raise SystemExit(f"could not find <{selector}> or <main> in {path}")
    h1 = soup.find("h1")
    raw = (h1.get_text() + " " if h1 else "") + node.get_text()
    return normalize(raw)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source", type=Path)
    ap.add_argument("target", type=Path)
    ap.add_argument("--encoding", default="utf-16le")
    ap.add_argument("--target-selector", default="article")
    ap.add_argument(
        "--exclude",
        action="append",
        default=[],
        help="regex matched against the raw source text and removed before diffing "
        "(e.g. leftover old-site nav chrome). May be repeated.",
    )
    args = ap.parse_args()

    src = source_words(args.source, args.encoding, args.exclude)
    tgt = target_words(args.target, args.target_selector)

    if args.exclude:
        print(f"# excluded from source before diff: {args.exclude}", file=sys.stderr)

    if src == tgt:
        print(f"OK: {args.target} matches {args.source} word-for-word ({len(tgt)} words)")
        return

    diff = list(difflib.unified_diff(src, tgt, fromfile=str(args.source), tofile=str(args.target), lineterm=""))
    print(f"MISMATCH: {args.target} vs {args.source}", file=sys.stderr)
    print(f"  source words: {len(src)}  target words: {len(tgt)}", file=sys.stderr)
    for line in diff[:200]:
        print(line, file=sys.stderr)
    raise SystemExit(1)


if __name__ == "__main__":
    main()
