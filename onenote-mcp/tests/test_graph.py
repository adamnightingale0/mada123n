import httpx
import pytest
import respx

from onenote_mcp.graph import GRAPH_BASE, GraphClient, GraphError, segment


def make_client(tokens=None):
    tokens = tokens if tokens is not None else []
    sleeps = []

    async def token(force_refresh: bool) -> str:
        tokens.append(force_refresh)
        return "fresh" if force_refresh else "tok"

    async def sleep(seconds: float) -> None:
        sleeps.append(seconds)

    return GraphClient(token, httpx.AsyncClient(), sleep=sleep), tokens, sleeps


def test_segment_quotes_ids():
    assert segment("0-ABC!1-DEF") == "0-ABC!1-DEF"
    assert segment("a/../b") == "a%2F..%2Fb"
    with pytest.raises(GraphError):
        segment("  ")


@respx.mock
async def test_follows_next_link():
    respx.get(f"{GRAPH_BASE}/me/onenote/notebooks").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "1"}], "@odata.nextLink": f"{GRAPH_BASE}/next?page=2"}
        )
    )
    respx.get(f"{GRAPH_BASE}/next").mock(return_value=httpx.Response(200, json={"value": [{"id": "2"}]}))
    client, _, _ = make_client()
    assert [n["id"] for n in await client.notebooks()] == ["1", "2"]


@respx.mock
async def test_retries_429_using_retry_after():
    route = respx.get(f"{GRAPH_BASE}/me/onenote/pages/p1/content").mock(
        side_effect=[
            httpx.Response(429, headers={"Retry-After": "7"}),
            httpx.Response(200, text="<html></html>"),
        ]
    )
    client, _, sleeps = make_client()
    assert await client.page_content("p1") == "<html></html>"
    assert route.call_count == 2 and sleeps == [7.0]


@respx.mock
async def test_gives_up_on_persistent_throttling():
    respx.get(f"{GRAPH_BASE}/me/onenote/pages/p1/content").mock(
        return_value=httpx.Response(429, headers={"Retry-After": "120"},
                                    json={"error": {"code": "20166", "message": "Too many requests"}})
    )
    client, _, sleeps = make_client()
    with pytest.raises(GraphError) as exc:
        await client.page_content("p1")
    assert exc.value.status == 429
    assert "throttling" in str(exc.value) and "20166" in str(exc.value)
    assert sleeps == [30.0, 30.0, 30.0]  # capped


@respx.mock
async def test_refreshes_token_once_on_401():
    route = respx.get(f"{GRAPH_BASE}/me/onenote/notebooks").mock(
        side_effect=[httpx.Response(401), httpx.Response(200, json={"value": []})]
    )
    client, tokens, _ = make_client()
    assert await client.notebooks() == []
    assert tokens == [False, True]
    assert route.calls[1].request.headers["Authorization"] == "Bearer fresh"


@respx.mock
async def test_404_has_helpful_message():
    respx.get(f"{GRAPH_BASE}/me/onenote/sections/bad/pages").mock(
        return_value=httpx.Response(404, json={"error": {"code": "20102", "message": "nope"}})
    )
    client, _, _ = make_client()
    with pytest.raises(GraphError, match="list_\\* tools"):
        await client.section_pages("bad", 10)


@respx.mock
async def test_sections_include_section_groups():
    def coll(*items):
        return httpx.Response(200, json={"value": list(items)})

    respx.get(f"{GRAPH_BASE}/me/onenote/notebooks/nb/sections").mock(return_value=coll({"id": "s1", "displayName": "Top"}))
    respx.get(f"{GRAPH_BASE}/me/onenote/notebooks/nb/sectionGroups").mock(return_value=coll({"id": "g1", "displayName": "Work"}))
    respx.get(f"{GRAPH_BASE}/me/onenote/sectionGroups/g1/sections").mock(return_value=coll({"id": "s2", "displayName": "Meetings"}))
    respx.get(f"{GRAPH_BASE}/me/onenote/sectionGroups/g1/sectionGroups").mock(return_value=coll({"id": "g2", "displayName": "Old"}))
    respx.get(f"{GRAPH_BASE}/me/onenote/sectionGroups/g2/sections").mock(return_value=coll({"id": "s3", "displayName": "2019"}))
    respx.get(f"{GRAPH_BASE}/me/onenote/sectionGroups/g2/sectionGroups").mock(return_value=coll())
    client, _, _ = make_client()
    paths = [s["path"] for s in await client.notebook_sections("nb")]
    assert paths == ["Top", "Work / Meetings", "Work / Old / 2019"]


@respx.mock
async def test_max_items_reports_truncation():
    respx.get(f"{GRAPH_BASE}/me/onenote/sections/s/pages").mock(
        return_value=httpx.Response(200, json={"value": [{"id": str(i)} for i in range(5)],
                                               "@odata.nextLink": f"{GRAPH_BASE}/more"})
    )
    client, _, _ = make_client()
    pages, more = await client.section_pages("s", 3)
    assert len(pages) == 3 and more


# -- writes ----------------------------------------------------------------

from onenote_mcp.graph import Part  # noqa: E402


@respx.mock
async def test_copy_page_polls_operation():
    respx.post(f"{GRAPH_BASE}/me/onenote/pages/p1/copyToSection").mock(
        return_value=httpx.Response(202, headers={"Operation-Location": f"{GRAPH_BASE}/me/onenote/operations/op1"},
                                    json={"id": "op1", "status": "NotStarted"})
    )
    respx.get(f"{GRAPH_BASE}/me/onenote/operations/op1").mock(side_effect=[
        httpx.Response(200, json={"status": "Running"}),
        httpx.Response(200, json={"status": "Completed", "resourceId": "p2"}),
    ])
    client, _, _ = make_client()
    assert await client.copy_page("p1", "s1") == "p2"
    sent = respx.calls[0].request
    assert sent.method == "POST" and b'"id":"s1"' in sent.content.replace(b" ", b"")


@respx.mock
async def test_copy_page_failure_is_reported():
    respx.post(f"{GRAPH_BASE}/me/onenote/pages/p1/copyToSection").mock(
        return_value=httpx.Response(202, json={"id": "op1"})
    )
    respx.get(f"{GRAPH_BASE}/me/onenote/operations/op1").mock(
        return_value=httpx.Response(200, json={"status": "Failed", "error": {"message": "nope"}})
    )
    client, _, _ = make_client()
    with pytest.raises(GraphError, match="nope"):
        await client.copy_page("p1", "s1")


@respx.mock
async def test_writes_retry_429_but_not_5xx():
    route = respx.patch(f"{GRAPH_BASE}/me/onenote/pages/p1/content").mock(side_effect=[
        httpx.Response(429, headers={"Retry-After": "1"}),
        httpx.Response(503),
    ])
    client, _, sleeps = make_client()
    with pytest.raises(GraphError) as exc:
        await client.update_page("p1", [{"target": "body", "action": "append", "content": "<p>x</p>"}], [])
    assert exc.value.status == 503 and route.call_count == 2 and sleeps == [1.0]


@respx.mock
async def test_create_page_multipart_with_image():
    route = respx.post(f"{GRAPH_BASE}/me/onenote/sections/s1/pages").mock(
        return_value=httpx.Response(201, json={"id": "new"})
    )
    client, _, _ = make_client()
    page = await client.create_page("s1", "<html><body><img src='name:img1'/></body></html>",
                                    [Part("img1", b"\x89PNG", "image/png")])
    assert page["id"] == "new"
    req = route.calls[0].request
    assert req.headers["Content-Type"].startswith("multipart/form-data")
    assert b'name="Presentation"' in req.content and b'name="img1"' in req.content


@respx.mock
async def test_create_page_plain_html_and_patch_json():
    create = respx.post(f"{GRAPH_BASE}/me/onenote/sections/s1/pages").mock(
        return_value=httpx.Response(201, json={"id": "new"})
    )
    patch = respx.patch(f"{GRAPH_BASE}/me/onenote/pages/new/content").mock(return_value=httpx.Response(204))
    client, _, _ = make_client()
    await client.create_page("s1", "<html></html>", [])
    await client.update_page("new", [{"target": "title", "action": "replace", "content": "T"}], [])
    assert create.calls[0].request.headers["Content-Type"] == "text/html"
    assert patch.calls[0].request.headers["Content-Type"] == "application/json"


@respx.mock
async def test_copy_page_gives_up_with_408():
    respx.post(f"{GRAPH_BASE}/me/onenote/pages/p1/copyToSection").mock(
        return_value=httpx.Response(202, json={"id": "op1"})
    )
    respx.get(f"{GRAPH_BASE}/me/onenote/operations/op1").mock(
        return_value=httpx.Response(200, json={"status": "Running"})
    )
    client, _, _ = make_client()
    waits = []
    with pytest.raises(GraphError) as exc:
        await client.copy_page("p1", "s1", timeout=10, on_wait=waits.append)
    assert exc.value.status == 408 and waits and waits[-1] >= 10
