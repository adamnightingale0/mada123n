"""Build OneNote page HTML (and binary parts) from Markdown written by Claude."""

from __future__ import annotations

import html as htmllib
import mimetypes
from pathlib import Path
from urllib.parse import unquote

from bs4 import BeautifulSoup, NavigableString
from markdown_it import MarkdownIt

from .graph import Part

# Inline HTML is allowed for formatting Markdown can't express (colour, highlight,
# underline...), then cut down to an allowlist of tags, attributes and styles.
_md = MarkdownIt("commonmark", {"html": True}).enable(["table", "strikethrough"])

ALLOWED_TAGS = {
    "p", "br", "span", "b", "strong", "i", "em", "u", "s", "del", "sub", "sup", "code", "pre",
    "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "table", "thead", "tbody", "tr",
    "th", "td", "a", "img", "cite", "blockquote",
}
DROP_WITH_CONTENT = {"script", "style", "iframe", "object", "embed", "form", "input", "textarea", "head", "title"}
ALLOWED_ATTRS = {"a": {"href"}, "img": {"src", "alt", "width", "height"}, "p": {"data-tag"}, "li": {"data-tag"}}
ALLOWED_STYLES = {
    "color", "background-color", "font-size", "font-family", "font-weight", "font-style", "text-decoration",
}

IMAGE_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
}
MAX_PART_BYTES = 25 * 1024 * 1024


class ContentError(Exception):
    """The requested content can't be built. The message is for the user/model."""


class Parts:
    """Collects images/files to upload with a request and names them uniquely."""

    def __init__(self) -> None:
        self.items: list[Part] = []

    def add(self, data: bytes, content_type: str, prefix: str = "part") -> str:
        name = f"{prefix}{len(self.items) + 1}"
        self.items.append(Part(name, data, content_type))
        return name

    def add_local_file(self, path_str: str, images_only: bool) -> tuple[str, str, str]:
        """Read a local file into a part. Returns (part name, file name, content type)."""
        path = Path(path_str.strip().strip('"')).expanduser()
        if not path.is_file():
            raise ContentError(f"File not found: {path}")
        size = path.stat().st_size
        if size > MAX_PART_BYTES:
            raise ContentError(f"{path.name} is {size / 1e6:.1f} MB; the limit is 25 MB per file.")
        ext = path.suffix.lower()
        if images_only:
            if ext not in IMAGE_TYPES:
                raise ContentError(
                    f"{path.name} isn't a supported image type ({', '.join(sorted(IMAGE_TYPES))}). "
                    "Use attach_file for other files."
                )
            ctype = IMAGE_TYPES[ext]
        else:
            ctype = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        name = self.add(path.read_bytes(), ctype, "img" if images_only else "file")
        return name, path.name, ctype


def _is_url(src: str) -> bool:
    return src.lower().startswith(("http://", "https://"))


def image_tag(src: str, parts: Parts, alt: str = "", width: int | None = None) -> str:
    """<img> for a URL (fetched by OneNote) or a local path (uploaded as a part)."""
    if _is_url(src):
        ref = src
    else:
        name, _, _ = parts.add_local_file(src, images_only=True)
        ref = f"name:{name}"
    attrs = f'src="{htmllib.escape(ref)}" alt="{htmllib.escape(alt)}"'
    if width:
        attrs += f' width="{int(width)}"'
    return f"<img {attrs} />"


def attachment_tag(path: str, parts: Parts) -> str:
    name, filename, ctype = parts.add_local_file(path, images_only=False)
    return (
        f'<object data-attachment="{htmllib.escape(filename)}" data="name:{name}" '
        f'type="{htmllib.escape(ctype)}"></object>'
    )


def _task_lists(soup: BeautifulSoup) -> None:
    """GitHub-style "- [ ] item" / "- [x] item" become OneNote to-do tags."""
    for li in soup.find_all("li"):
        first = next((t for t in li.descendants if isinstance(t, NavigableString) and t.strip()), None)
        if first is None:
            continue
        text = first.lstrip()
        for marker, tag in (("[ ] ", "to-do"), ("[x] ", "to-do:completed"), ("[X] ", "to-do:completed")):
            if text.startswith(marker):
                first.replace_with(NavigableString(text[len(marker):]))
                li["data-tag"] = tag
                break


def _sanitize(soup: BeautifulSoup) -> None:
    for el in soup.find_all(DROP_WITH_CONTENT):
        el.decompose()
    for el in soup.find_all(True):
        if el.name not in ALLOWED_TAGS:
            el.unwrap()
            continue
        keep = ALLOWED_ATTRS.get(el.name, set())
        style = el.get("style")
        el.attrs = {k: v for k, v in el.attrs.items() if k in keep}
        if style:
            decls = []
            for decl in str(style).split(";"):
                prop, _, value = decl.partition(":")
                prop, value = prop.strip().lower(), value.strip()
                if prop in ALLOWED_STYLES and value and not any(c in value for c in "<>\"(){}"):
                    decls.append(f"{prop}:{value}")
            if decls:
                el["style"] = ";".join(decls)
        href = el.get("href")
        if href and not str(href).lower().startswith(("http://", "https://", "mailto:", "onenote:")):
            del el["href"]


def markdown_to_html(markdown: str, parts: Parts) -> str:
    """Render Markdown to OneNote-compatible HTML.

    Images written as ![alt](path-or-url) are uploaded from local paths or
    fetched by OneNote from URLs.
    """
    soup = BeautifulSoup(_md.render(markdown), "html.parser")
    _sanitize(soup)
    _task_lists(soup)
    for img in soup.find_all("img"):
        src = img.get("src", "")
        if not _is_url(src):
            name, _, _ = parts.add_local_file(unquote(src), images_only=True)
            img["src"] = f"name:{name}"
    for hr in soup.find_all("hr"):  # OneNote has no horizontal rule
        hr.decompose()
    return str(soup).strip()


def page_html(title: str, body_html: str) -> str:
    return (
        "<!DOCTYPE html><html><head>"
        f"<title>{htmllib.escape(title)}</title>"
        '<meta charset="utf-8" />'
        f"</head><body>{body_html}</body></html>"
    )
