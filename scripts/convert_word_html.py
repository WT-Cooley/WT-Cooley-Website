#!/usr/bin/env python3
"""Strip Microsoft Word "Filtered HTML" export cruft down to a clean content fragment.

This is a mechanical first pass only. It does not write a full page -- it prints
a cleaned HTML fragment (the content that belongs inside <article>, or the list
sections for a landing page) for a human to hand-place into the site template
and verify against the source with verify_text_parity.py. No wording is changed;
only presentational markup is stripped or restructured (Word bullet paragraphs
become real <ul><li> lists).

Usage:
    python3 convert_word_html.py "path/to/Source.html" [--encoding auto|utf-16le|utf-8|cp1252] [--out fragment.html]

Encoding defaults to "auto": these files come from three eras of export tooling
(classic Word "Filtered HTML" is UTF-16LE; a batch of older files across the
corpus are actually Windows-1252 despite `file` reporting them as plain ASCII,
because they only contain a handful of non-ASCII bytes -- smart quotes, en-dashes,
middle dots; Google Docs exports and hand-authored files are UTF-8). Auto-detection
tries UTF-16LE first (via BOM or null-byte density), then UTF-8, then falls back
to cp1252, which never raises on arbitrary bytes.
"""
from __future__ import annotations

import argparse
import html as html_module
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup, Tag

MSO_LIST_RE = re.compile(r"mso-list:\s*(l\d+)\s+level(\d+)", re.I)


def detect_encoding(raw: bytes) -> str:
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return "utf-16"
    # Word's UTF-16LE exports have no BOM but are dense with 0x00 bytes
    # (every ASCII codepoint is followed by a null byte).
    sample = raw[:4000]
    if sample.count(b"\x00") > len(sample) * 0.2:
        return "utf-16le"
    try:
        raw.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "cp1252"


def load_source(path: Path, encoding: str) -> str:
    raw = path.read_bytes()
    if encoding == "auto":
        encoding = detect_encoding(raw)
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
    # font-run, spellcheck (SpellE/GramE), or spacer wrapper -- EXCEPT Word's real
    # highlighter tool (mso-highlight:<color>), which at least one course (Fantasy
    # & Religion) uses to mark passages an explicit in-document NOTE calls
    # "optional" reading. Promote those runs to a real <mark> instead of stripping
    # them, since dropping the highlight would leave that NOTE's own claim
    # meaningless. Everything else gets unwrapped as before.
    for span in root.find_all("span"):
        if span.attrs is None:
            continue  # already decomposed as a descendant of a span removed earlier
        style = (span.get("style") or "").lower()
        if "mso-highlight" in style:
            span.name = "mark"
            continue
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
        if tag.name in ("table", "tr", "td", "th"):
            # Word tables carry fixed pixel widths and HTML4-era presentational
            # attrs -- strip them (CSS handles table styling) but keep rowspan/
            # colspan, which are structural, not cosmetic.
            for attr in ("width", "border", "cellpadding", "cellspacing", "valign", "bgcolor", "height"):
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


def normalize_whitespace(root: Tag) -> None:
    """Collapse whitespace runs (including the literal newlines Word inserts at its
    own arbitrary line-wrap points) to single spaces within each text node. This does
    not change any word -- it only undoes Word's cosmetic mid-sentence line wrapping,
    the same normalization a browser applies visually to any HTML whitespace anyway."""
    for node in root.find_all(string=True):
        collapsed = re.sub(r"\s+", " ", str(node))
        if collapsed != str(node):
            node.replace_with(collapsed)


def prettify_fragment(root: Tag) -> str:
    # One top-level element (p/ul/h2/...) per line, serialized as-is -- NOT via
    # BeautifulSoup's prettify(), which inserts a newline (whitespace) between
    # every adjacent tag pair for readability, even when the source had zero
    # characters between them (e.g. a footnote marker "word[a]" glued to a word).
    # That's fine when this fragment is only read by a human, but this fragment
    # sometimes gets copied verbatim into the final page, so an inserted
    # separator there would be a real, published text alteration, not just
    # cosmetic. Serializing each top-level node with str() preserves the exact
    # inter-tag spacing (already normalized to single spaces/none by
    # normalize_whitespace) with nothing added.
    lines = []
    for child in root.contents:
        if getattr(child, "name", None):
            lines.append(str(child))
        else:
            # NavigableString.__str__ does not re-escape &/</> the way a
            # parent tag's str() does -- escape stray top-level text nodes
            # by hand so "&" doesn't silently become invalid raw "&" in output.
            text = html_module.escape(str(child)).strip()
            if text:
                lines.append(text)
    return "\n".join(lines)


def find_endnote_list(soup: BeautifulSoup) -> Tag | None:
    return soup.find("div", style=lambda s: bool(s) and "mso-element:endnote-list" in s)


def convert_endnotes(path: Path, encoding: str) -> list[tuple[str, str]]:
    """Word's real endnote *text* (References > Insert Endnote) lives in a
    <div style='mso-element:endnote-list'> that is a SIBLING of WordSection1,
    not inside it -- find_root()/convert() never sees it, silently dropping
    real footnote content on any document that uses this feature (mainly
    academic papers, unlike the informal inline citations used elsewhere).
    Returns [(marker, cleaned_html), ...] in document order, e.g. ("i", "...").
    """
    text = load_source(path, encoding)
    text = strip_conditional_blocks(text)
    soup = BeautifulSoup(text, "html.parser")
    endnote_list = find_endnote_list(soup)
    if endnote_list is None:
        return []

    results = []
    for div in endnote_list.find_all("div", id=re.compile(r"^edn\d+$")):
        num = re.search(r"\d+", div["id"]).group()
        # Drop the endnote's own back-reference marker link (e.g. "[i]") --
        # it's plumbing to jump back to the body, not part of the note's text.
        for a in div.find_all("a", attrs={"name": re.compile(r"^_edn\d+$")}):
            a.decompose()
        unwrap_presentational_spans(div)
        strip_presentational_attrs(div)
        remove_o_p_tags(div)
        remove_stray_vml(div)
        drop_spacer_paragraphs(div)
        normalize_whitespace(div)
        html = prettify_fragment(div)
        results.append((num, html))
    return results


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
    normalize_whitespace(root)

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
    ap.add_argument("--encoding", default="auto")
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
