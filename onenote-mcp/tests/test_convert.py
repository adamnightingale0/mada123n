from onenote_mcp.convert import EMPTY_PAGE_NOTE, html_to_markdown, render_page

PAGE = """<html lang="en-US"><head><title>Trip &amp; plans</title>
<meta name="created" content="2024-01-02T09:00:00.0000000" /></head>
<body data-absolute-enabled="true" style="font-family:Calibri">
<div style="position:absolute;left:48px;top:120px">
<h1>Weekend</h1>
<p>Some <span style="font-weight:bold">bold</span> and <span style="font-style:italic">italic</span>&nbsp;text</p>
<p data-tag="to-do">Buy milk</p>
<p data-tag="to-do:completed,important">Call home</p>
<ul><li>one</li><li data-tag="question">two?</li></ul>
<img alt="A cat" src="https://graph.microsoft.com/x/$value" />
<object data-attachment="Plan.pdf" type="application/pdf" data="https://x" />
<table><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2</td></tr></table>
</div>
<div style="position:absolute;left:400px;top:120px"><p>Second outline</p></div>
</body></html>"""


def test_converts_onenote_markup():
    md, title, created = html_to_markdown(PAGE)
    assert title == "Trip & plans"
    assert created.startswith("2024-01-02")
    assert "# Weekend" in md
    assert "Some **bold** and *italic* text" in md
    assert "[ ] Buy milk" in md
    assert "[x] [Important] Call home" in md
    assert "- [Question] two?" in md
    assert "[Image 1: A cat]" in md
    assert "[Attachment: Plan.pdf]" in md
    assert "| a | b |\n| --- | --- |\n| 1 | 2 |" in md
    assert "Second outline" in md
    assert "position:absolute" not in md and " " not in md


def test_render_includes_header_and_truncates():
    out = render_page(PAGE.replace("Second outline", "x" * 5000), "0-abc!1", max_chars=1000)
    assert out.startswith("# Trip & plans\n\nPage ID: 0-abc!1")
    assert "[Truncated:" in out
    assert len(out) < 1500


def test_empty_page_explains_ink():
    out = render_page("<html><head><title>Ink</title></head><body><div></div></body></html>", "p", 1000)
    assert EMPTY_PAGE_NOTE in out


def test_long_image_text_moves_after_typed_notes():
    slide = "TOPIC 2 Monoclonal Antibodies " * 10
    html = (
        "<html><head><title>L2</title></head><body><div>"
        f'<object data-attachment="Topic_2.pdf"></object><img alt="{slide}"/><img alt="Logo"/>'
        "<ul><li>Rituximab targets CD20</li></ul></div></body></html>"
    )
    md, _, _ = html_to_markdown(html)
    assert "[Attachment: Topic_2.pdf]" in md
    assert "[Image 1, text below]" in md and "[Image 2: Logo]" in md
    assert md.index("Rituximab") < md.index("## Text found in images (OCR)") < md.index("**Image 1:** TOPIC 2")

    short, _, _ = html_to_markdown(html, image_text=False)
    assert "TOPIC 2" not in short and "Text from 1 image(s) left out" in short
    assert "Rituximab" in short
