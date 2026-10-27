"""In-memory stand-in for GraphClient's OneNote calls, for editing tests."""

import asyncio
import itertools

from onenote_mcp.graph import GraphError

GRAPH = "https://graph.microsoft.com/v1.0"
PNG = b"\x89PNG\r\n\x1a\n" + b"\x00" * 8


def page_doc(title, paragraphs, prefix):
    body = "".join(
        t if t.startswith("<") else f'<p id="p:{{{prefix}}}{{{i}}}">{t}</p>'
        for i, t in enumerate(paragraphs, 1)
    )
    return f'<html><head><title>{title}</title></head><body><div id="div:{{{prefix}}}{{0}}">{body}</div></body></html>'


class FakeOneNote:
    def __init__(self, copy_supported=True, copy_delay=0.0, fail_downloads=()):
        self.copy_supported = copy_supported
        self.copy_delay = copy_delay
        self.fail_downloads = set(fail_downloads)
        self.downloads = []
        self._ids = itertools.count(1)
        self.pages = {}  # id -> {"title", "section", "paragraphs"}
        self.updates = []  # (page_id, commands, parts)
        self.created = []  # (section_id, html, parts)
        self.deleted = []

    def add_page(self, title, paragraphs, section="sec1"):
        pid = f"page{next(self._ids)}"
        self.pages[pid] = {"title": title, "section": section, "paragraphs": list(paragraphs)}
        return pid

    def _get(self, page_id):
        if page_id not in self.pages:
            raise GraphError("Not found (HTTP 404).", 404)
        return self.pages[page_id]

    async def page(self, page_id):
        p = self._get(page_id)
        return {"id": page_id, "title": p["title"], "parentSection": {"id": p["section"]},
                "links": {"oneNoteClientUrl": {"href": f"onenote:https://x/{page_id}"}}}

    async def page_content(self, page_id, include_ids=False):
        p = self._get(page_id)
        html = page_doc(p["title"], p["paragraphs"], page_id)
        if not include_ids:
            html = html.replace(' id="', ' data-x="')
        extra = (
            f'<object data-attachment="Topic_2.pdf" type="application/pdf" '
            f'data="{GRAPH}/me/onenote/resources/r2/$value"></object>'
            f'<img src="{GRAPH}/me/onenote/resources/r1/$value" alt="pic">'
        )
        return html.replace("</div></body>", extra + "</div></body>")

    async def copy_page(self, page_id, section_id, timeout=90.0, on_wait=None):
        if not self.copy_supported:
            raise GraphError("Copy not supported (HTTP 400).", 400)
        if self.copy_delay:
            if on_wait:
                on_wait(0)
            await asyncio.sleep(self.copy_delay)
        p = self._get(page_id)
        return self.add_page(p["title"], p["paragraphs"], section_id)

    async def create_page(self, section_id, html, parts):
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html, "html.parser")
        title = soup.title.get_text() if soup.title else ""
        paras = [p.get_text() for p in soup.find_all("p")]
        self.created.append((section_id, html, parts))
        pid = self.add_page(title, paras, section_id)
        return {"id": pid}

    async def update_page(self, page_id, commands, parts):
        from bs4 import BeautifulSoup

        p = self._get(page_id)
        for c in commands:
            target = c["target"]
            if c["action"] == "replace" and target.startswith("#"):
                # Like OneNote: paragraphs can only be replaced by their generated id.
                raise GraphError("Bad request (HTTP 400). Details: invalid target", 400)
        self.updates.append((page_id, commands, parts))
        for c in commands:
            target = c["target"]
            if target == "title" and c["action"] == "replace":
                p["title"] = c["content"]
            elif c["action"] == "replace" and target.startswith(f"p:{{{page_id}}}"):
                k = int(target.rsplit("{", 1)[1].rstrip("}"))
                soup = BeautifulSoup(c["content"], "html.parser")
                for el in soup.find_all(["img", "object"]):  # stored as a page resource
                    attr = "src" if el.name == "img" else "data"
                    el[attr] = f"{GRAPH}/me/onenote/resources/{page_id}-{el[attr][5:]}/$value"
                p["paragraphs"][k - 1] = str(soup) if soup.find(["img", "object"]) else soup.get_text()

    async def get_bytes(self, url):
        self.downloads.append(url)
        if any(url.endswith(f"{r}/$value") for r in self.fail_downloads):
            raise GraphError("Not found (HTTP 404).", 404)
        if url.endswith("r2/$value") or "-file" in url:
            return b"%PDF-1.7", "application/octet-stream"
        return PNG, "application/octet-stream"

    async def delete_page(self, page_id):
        self._get(page_id)
        del self.pages[page_id]
        self.deleted.append(page_id)

    async def create_section(self, notebook_id, name):
        return {"id": "sec-new", "displayName": name}
