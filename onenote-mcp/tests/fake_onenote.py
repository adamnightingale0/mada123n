"""In-memory stand-in for GraphClient's OneNote calls, for editing tests."""

import itertools

from onenote_mcp.graph import GraphError

GRAPH = "https://graph.microsoft.com/v1.0"


def page_doc(title, paragraphs, prefix):
    body = "".join(f'<p id="p:{{{prefix}}}{{{i}}}">{t}</p>' for i, t in enumerate(paragraphs, 1))
    return f'<html><head><title>{title}</title></head><body><div id="div:{{{prefix}}}{{0}}">{body}</div></body></html>'


class FakeOneNote:
    def __init__(self, copy_supported=True):
        self.copy_supported = copy_supported
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
        return html + (
            f'<object data-attachment="Topic_2.pdf" type="application/pdf" '
            f'data="{GRAPH}/me/onenote/resources/r2/$value"></object>'
            f'<img src="{GRAPH}/me/onenote/resources/r1/$value" alt="pic">'
        )

    async def copy_page(self, page_id, section_id):
        if not self.copy_supported:
            raise GraphError("Copy not supported (HTTP 400).", 400)
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
        p = self._get(page_id)
        self.updates.append((page_id, commands, parts))
        for c in commands:
            if c["target"] == "title" and c["action"] == "replace":
                p["title"] = c["content"]

    async def get_bytes(self, url):
        if url.endswith("r2/$value"):
            return b"%PDF", "application/pdf"
        return b"\x89PNG", "image/png"

    async def delete_page(self, page_id):
        self._get(page_id)
        del self.pages[page_id]
        self.deleted.append(page_id)

    async def create_section(self, notebook_id, name):
        return {"id": "sec-new", "displayName": name}
