"""Turn OneNote page HTML into compact Markdown for the model to read."""

from __future__ import annotations

import re
from dataclasses import dataclass

from bs4 import BeautifulSoup, NavigableString, Tag
from markdownify import MarkdownConverter

# OneNote note tags (data-tag) that we render as a short textual marker.
TAG_LABELS = {
    "important": "Important",
    "question": "Question",
    "critical": "Critical",
    "idea": "Idea",
    "definition": "Definition",
    "highlight": "Highlight",
    "contact": "Contact",
    "address": "Address",
    "phone-number": "Phone",
    "remember-for-later": "Remember for later",
    "project-a": "Project A",
    "project-b": "Project B",
}

EMPTY_PAGE_NOTE = (
    "_No text content found. The page may contain only handwriting/ink, drawings or "
    "images, which Microsoft Graph does not return as text._"
)


class _OneNoteConverter(MarkdownConverter):
    def convert_img(self, el, text, parent_tags=None, **kwargs):
        alt = (el.get("alt") or "").strip()
        return f" [Image: {alt}] " if alt else " [Image] "


def _tag_prefix(value: str) -> str:
    parts = [p.strip() for p in value.split(",") if p.strip()]
    out = []
    for p in parts:
        if p == "to-do":
            out.append("[ ]")
        elif p == "to-do:completed":
            out.append("[x]")
        else:
            name, _, state = p.partition(":")
            label = TAG_LABELS.get(name, name.replace("-", " ").capitalize())
            out.append(f"[{label}{' (done)' if state == 'completed' else ''}]")
    return " ".join(out)


def _preprocess(soup: BeautifulSoup) -> None:
    for el in soup.find_all(["script", "style"]):
        el.decompose()

    for el in soup.find_all(attrs={"data-tag": True}):
        prefix = _tag_prefix(el["data-tag"])
        if prefix:
            el.insert(0, NavigableString(prefix + " "))

    for el in soup.find_all("object"):
        name = el.get("data-attachment") or el.get("type") or "file"
        el.replace_with(NavigableString(f"\n\n[Attachment: {name}]\n\n"))

    # OneNote expresses bold/italic as inline styles on spans.
    for el in soup.find_all("span", style=True):
        style = el["style"].replace(" ", "").lower()
        if "font-weight:bold" in style:
            el.name = "strong"
        elif "font-style:italic" in style:
            el.name = "em"

    for el in soup.find_all("iframe"):
        src = el.get("data-original-src") or el.get("src") or ""
        el.replace_with(NavigableString(f"[Embedded content: {src}]" if src else "[Embedded content]"))

    # Each absolutely-positioned outline div is a separate block on the page canvas.
    for el in soup.find_all("div"):
        el.insert_after(NavigableString("\n\n"))


def html_to_markdown(html: str) -> tuple[str, str | None, str | None]:
    """Convert page HTML. Returns (markdown_body, title, created)."""
    soup = BeautifulSoup(html, "html.parser")
    title_el = soup.find("title")
    title = title_el.get_text(strip=True) if title_el else None
    created_el = soup.find("meta", attrs={"name": "created"})
    created = created_el.get("content") if isinstance(created_el, Tag) else None

    body = soup.body or soup
    if soup.head:
        soup.head.decompose()
    _preprocess(soup)

    md = _OneNoteConverter(heading_style="ATX", bullets="-", table_infer_header=True, strip=["span", "font"]).convert_soup(body)
    md = md.replace(" ", " ")
    md = re.sub(r"[ \t]+\n", "\n", md)
    md = re.sub(r"^ (\[Image)", r"\1", md, flags=re.MULTILINE)
    md = re.sub(r"\n{3,}", "\n\n", md).strip()
    return md, title, created


def render_page(html: str, page_id: str, max_chars: int) -> str:
    body, title, created = html_to_markdown(html)
    lines = [f"# {title or '(untitled page)'}", "", f"Page ID: {page_id}"]
    if created:
        lines.append(f"Created: {created}")
    lines.append("")
    if not body:
        lines.append(EMPTY_PAGE_NOTE)
    elif len(body) > max_chars:
        cut = body[:max_chars].rsplit("\n", 1)[0] or body[:max_chars]
        lines.append(cut)
        lines.append("")
        lines.append(
            f"_[Truncated: showing {len(cut):,} of {len(body):,} characters. "
            "Call get_page_content again with a larger max_chars to see more.]_"
        )
    else:
        lines.append(body)
    return "\n".join(lines)


# -- page structure (element IDs for targeted edits) -----------------------

STRUCTURE_TAGS = {"div", "p", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "table", "img", "object", "iframe"}


@dataclass
class Element:
    target: str
    tag: str
    depth: int
    preview: str


def _preview(el: Tag) -> str:
    if el.name == "img":
        return f"[Image: {el.get('alt', '')}]".replace(": ]", "]")
    if el.name == "object":
        return f"[Attachment: {el.get('data-attachment', 'file')}]"
    text = " ".join(el.get_text(" ", strip=True).split())
    return text[:80] + ("…" if len(text) > 80 else "")


def page_structure(html: str) -> list[Element]:
    """Elements that edits can target, from page HTML fetched with includeIDs=true."""
    soup = BeautifulSoup(html, "html.parser")
    body = soup.body or soup
    out: list[Element] = []
    for el in body.find_all(STRUCTURE_TAGS):
        target = el.get("id") or (f"#{el['data-id']}" if el.get("data-id") else None)
        if not target:
            continue
        depth = sum(
            1 for parent in el.parents
            if isinstance(parent, Tag) and parent.name in STRUCTURE_TAGS
            and (parent.get("id") or parent.get("data-id"))
        )
        out.append(Element(target, el.name, depth, _preview(el)))
    return out


def render_structure(html: str, page_id: str) -> str:
    _, title, _ = html_to_markdown(html)
    elements = page_structure(html)
    lines = [
        f"# {title or '(untitled page)'}",
        f"Page ID: {page_id}",
        "",
        "Targets for edits (use the value in backticks as target_id; `title` is the page title):",
    ]
    for i, e in enumerate(elements, 1):
        lines.append(f"{'  ' * e.depth}{i}. {e.tag} `{e.target}` {e.preview}".rstrip())
    if not elements:
        lines.append("(no targetable elements; use position=end to add content)")
    return "\n".join(lines)
