import pytest

from onenote_mcp.compose import ContentError, Parts, attachment_tag, image_tag, markdown_to_html, page_html
from onenote_mcp.convert import page_structure, render_structure


def test_markdown_to_onenote_html(tmp_path):
    img = tmp_path / "my pic.png"
    img.write_bytes(b"\x89PNG")
    parts = Parts()
    html = markdown_to_html(
        f"# Plan\n\nSome **bold**.\n\n- [ ] milk\n- [x] call\n- plain\n\n"
        f"| a | b |\n|---|---|\n| 1 | 2 |\n\n![cat](<{img}>)\n\n![web](https://example.com/x.png)\n\n"
        "<script>alert(1)</script>\n\n---\n",
        parts,
    )
    assert "<h1>Plan</h1>" in html and "<strong>bold</strong>" in html
    assert '<li data-tag="to-do">milk</li>' in html
    assert '<li data-tag="to-do:completed">call</li>' in html
    assert "<li>plain</li>" in html and "<table>" in html
    assert 'src="name:img1"' in html and 'src="https://example.com/x.png"' in html
    assert "<script>" not in html and "alert" not in html and "<hr" not in html
    assert [(p.name, p.content_type) for p in parts.items] == [("img1", "image/png")]


def test_local_file_checks(tmp_path):
    parts = Parts()
    with pytest.raises(ContentError, match="not found"):
        image_tag(str(tmp_path / "missing.png"), parts)
    doc = tmp_path / "notes.pdf"
    doc.write_bytes(b"%PDF")
    with pytest.raises(ContentError, match="attach_file"):
        image_tag(str(doc), parts)
    tag = attachment_tag(str(doc), parts)
    assert 'data-attachment="notes.pdf"' in tag and 'data="name:file1"' in tag
    assert 'type="application/pdf"' in tag


def test_image_tag_url_and_width():
    tag = image_tag("https://example.com/a.png", Parts(), alt='a "b"', width=300)
    assert tag == '<img src="https://example.com/a.png" alt="a &quot;b&quot;" width="300" />'


def test_page_html_escapes_title():
    assert "<title>A &amp; B</title>" in page_html("A & B", "<p>x</p>")


def test_structure_lists_targets():
    html = (
        '<html><head><title>T</title></head><body><div id="div:{a}{1}">'
        '<h1 id="h1:{a}{2}">Heading</h1><p id="p:{a}{3}">Hello <span>world</span></p>'
        '<ul id="ul:{a}{4}"><li id="li:{a}{5}">item</li></ul>'
        '<img id="img:{a}{6}" alt="chart" src="x"/><p data-id="mine">custom</p><p>no id</p>'
        "</div></body></html>"
    )
    els = page_structure(html)
    assert [(e.target, e.depth) for e in els] == [
        ("div:{a}{1}", 0), ("h1:{a}{2}", 1), ("p:{a}{3}", 1), ("ul:{a}{4}", 1),
        ("li:{a}{5}", 2), ("img:{a}{6}", 1), ("#mine", 1),
    ]
    assert els[2].preview == "Hello world" and els[5].preview == "[Image: chart]"
    out = render_structure(html, "pg")
    assert "  2. h1 `h1:{a}{2}` Heading" in out and out.startswith("# T")


def test_inline_html_is_allowlisted():
    html = markdown_to_html(
        'A <span style="color:red; background-color: yellow; position:absolute">hot</span> '
        '<u>word</u> <span style="color:url(x)">bad</span> <a href="javascript:x()">link</a> '
        '<a href="https://ok.example" onclick="x()">ok</a> <font color="red">old</font>\n\n'
        '<div onclick="x()"><p>inside</p></div><iframe src="https://evil"></iframe>',
        Parts(),
    )
    assert '<span style="color:red;background-color:yellow">hot</span>' in html
    assert "<u>word</u>" in html and "<span>bad</span>" in html
    assert "<a>link</a>" in html and '<a href="https://ok.example">ok</a>' in html
    assert "old" in html and "<font" not in html
    assert "onclick" not in html and "<div" not in html and "iframe" not in html
    assert "<p>inside</p>" in html
