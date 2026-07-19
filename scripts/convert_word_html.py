#!/usr/bin/env python3
"""Strip Microsoft Word "Filtered HTML" export cruft down to a clean content fragment.

This is a mechanical first pass only. It does not write a full page -- it prints
a cleaned HTML fragment (the content that belongs inside <article>, or the list
sections for a landing page) for a human to hand-place into the site template
and verify against the source with verify_text_parity.py. No wording is changed;
only presentational markup is stripped or restructured (Word bullet paragraphs
become real <ul><li> lists).

Usage:
    python3 convert_word_html.py "path/to/Source.html" [--encoding utf-16le] [--out fragment.html]
"""
import argparse
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup, Tag

MSO_LIST_RE = re.compile(r"mso-list:\s*(l\d+)\s+level(\d+)", re.I)


def load_source(path: Path, encoding: str) -> str:
    raw = path.read_bytes()
    text = raw.decode(encoding)
    return text.replace("﻿", "")


def strip_conditional_blocks(html: str) -> str:
    # HTML-comment-style conditionals, e.g. <!--[if gte vml 1]>...VML...<![endif]-->
    # These wrap content we never want (VML shape defs, mso XML property blocks) -- drop whole block.
    html = re.sub(r"<!--\[if[^\]]*\]>.*?<!\[endif\]-->", "", html, flags=re.S)
    # Bare downlevel-revealed conditionals, e.g. <![if !vml]><img .../><![endif]>
    # These wrap content we DO want (the real fallback <img>) -- strip only the marker text.
    html = re.sub(r"<!\[if[^\]]*\]>", "", html)
    html = re.sub(r"<!\[endif\]>", "", html)
    return html


def find_root(soup: BeautifulSoup) -> Tag:
    root = soup.find("div", class_="WordSection1")
    if root is not None:
        return root
    body = soup.find("body")
    return body if body is not None else soup


def decompose_bullet_glyphs(root: Tag) -> None:
    for span in root.find_all("span"):
        if span.attrs is None:
            continue  # already decomposed as a descendant of a span removed earlier in this loop
        style = (span.get("style") or "").lower()
        if "mso-list:ignore" in style:
            span.decompose()


def capture_list_grouping(root: Tag) -> int:
    """Stamp data-mso-list-id/-level onto Word list paragraphs before styles are stripped.
    Returns count of list paragraphs found."""
    count = 0
    for p in root.find_all("p"):
        classes = p.get("class") or []
        if any(c.startswith("MsoListParagraph") for c in classes):
            style = p.get("style") or ""
            m = MSO_LIST_RE.search(style)
            if m:
                p["data-mso-list-id"] = m.group(1)
                p["data-mso-list-level"] = m.group(2)
                count += 1
            else:
                print(
                    f"WARNING: list paragraph without parseable mso-list style: "
                    f"{p.get_text(strip=True)[:60]!r}",
                    file=sys.stderr,
                )
    return count


def unwrap_presentational_spans(root: Tag) -> None:
    # Word never uses <span> for anything semantic -- every span in this corpus is a
    # font-run, spellcheck (SpellE/GramE), or spacer wrapper. Unwrap all of them,
    # keeping only their text/child content.
    for span in root.find_all("span"):
        if span.attrs is None:
            continue  # already decomposed as a descendant of a span removed earlier
        span.unwrap()


def handle_images(root: Tag) -> list[str]:
    srcs = []
    for img in root.find_all("img"):
        src = img.get("src", "")
        srcs.append(src)
        img["alt"] = "TODO: write real alt text (source had Word auto-generated description)"
        for attr in ("width", "height", "v:shapes", "border"):
            if img.has_attr(attr):
                del img[attr]
    return srcs


def strip_presentational_attrs(root: Tag) -> None:
    for tag in root.find_all(True):
        if tag.name in ("o:p", "v:shapetype", "v:shape", "v:stroke", "v:formulas", "v:f", "v:path", "o:lock"):
            continue
        if tag.has_attr("class"):
            kept = [c for c in tag["class"] if not c.lower().startswith("mso")]
            if kept:
                tag["class"] = kept
            else:
                del tag["class"]
        for attr in ("style", "lang", "align", "v:shapes"):
            if tag.has_attr(attr):
                del tag[attr]


def remove_o_p_tags(root: Tag) -> None:
    for tag in root.find_all("o:p"):
        tag.decompose()


def remove_stray_vml(root: Tag) -> None:
    for name in ("v:shapetype", "v:shape", "v:stroke", "v:formulas", "v:f", "v:path", "o:lock"):
        for tag in root.find_all(name):
            tag.decompose()


def is_blank_paragraph(p: Tag) -> bool:
    text = p.get_text(strip=True)
    return text in ("", "\xa0") and not p.find("img")


def drop_spacer_paragraphs(root: Tag) -> int:
    count = 0
    for p in root.find_all("p"):
        if is_blank_paragraph(p):
            p.decompose()
            count += 1
    return count


def group_lists(root: Tag) -> int:
    """Convert consecutive Word list paragraphs sharing a list id into <ul class="page-list"><li>."""
    count = 0
    p_tags = [p for p in root.find_all("p") if p.has_attr("data-mso-list-id")]
    seen = set()
    for p in p_tags:
        if id(p) in seen:
            continue
        list_id = p["data-mso-list-id"]
        group = [p]
        seen.add(id(p))
        nxt = p.find_next_sibling()
        while nxt is not None and getattr(nxt, "has_attr", None) and nxt.has_attr("data-mso-list-id") and nxt["data-mso-list-id"] == list_id:
            group.append(nxt)
            seen.add(id(nxt))
            nxt = nxt.find_next_sibling()
        ul = root.new_tag("ul")
        ul["class"] = "page-list"
        p.insert_before(ul)
        for item in group:
            del item["data-mso-list-id"]
            del item["data-mso-list-level"]
            item.name = "li"
            ul.append(item)
            count += 1
    return count


def prettify_fragment(root: Tag) -> str:
    inner = "".join(str(c) for c in root.contents)
    soup = BeautifulSoup(inner, "html.parser")
    out = soup.prettify()
    lines = [line for line in out.splitlines() if line.strip()]
    return "\n".join(lines)


def convert(path: Path, encoding: str) -> tuple[str, list[str], int]:
    text = load_source(path, encoding)
    text = strip_conditional_blocks(text)
    soup = BeautifulSoup(text, "html.parser")
    root = find_root(soup)

    decompose_bullet_glyphs(root)
    list_para_count = capture_list_grouping(root)
    unwrap_presentational_spans(root)
    image_srcs = handle_images(root)
    strip_presentational_attrs(root)
    remove_o_p_tags(root)
    remove_stray_vml(root)
    spacer_count = drop_spacer_paragraphs(root)
    li_count = group_lists(root)

    fragment = prettify_fragment(root)
    print(
        f"# {path.name}: {list_para_count} list paragraphs -> {li_count} <li>, "
        f"{spacer_count} spacer paragraphs dropped, {len(image_srcs)} image(s) found",
        file=sys.stderr,
    )
    for src in image_srcs:
        print(f"#   image src: {src}", file=sys.stderr)
    return fragment, image_srcs, li_count


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source", type=Path)
    ap.add_argument("--encoding", default="utf-16le")
    ap.add_argument("--out", type=Path, default=None)
    args = ap.parse_args()

    fragment, _, _ = convert(args.source, args.encoding)

    if args.out:
        args.out.write_text(fragment, encoding="utf-8")
        print(f"# wrote {args.out}", file=sys.stderr)
    else:
        print(fragment)


if __name__ == "__main__":
    main()
