# onenote-mcp

A local [MCP](https://modelcontextprotocol.io) server that lets Claude read, and optionally
edit, your Microsoft OneNote notebooks through Microsoft Graph. It runs on your machine over
stdio, for Claude Desktop and Claude Code. It is **read-only unless you turn editing on**.

- Signs in with the MSAL **device code flow** using delegated permissions (`Notes.Read`, or
  `Notes.ReadWrite` when editing is on), so it can only reach the notebooks of the account
  that signs in.
- Caches the token locally (owner-only file permissions), so you sign in once.
- Converts OneNote page HTML to clean Markdown.
- Logs to stderr only; stdout is reserved for MCP JSON-RPC.

> **Privacy:** the server talks only to Microsoft. When Claude reads a page in a chat, that
> page's text is sent to Anthropic as part of the conversation.

## Tools

| Tool | What it does |
| --- | --- |
| `list_notebooks` | Your notebooks, with IDs |
| `list_sections(notebook_id)` | Every section, including those inside section groups (shown as `Group / Section`) |
| `list_pages(section_id, limit=100)` | Page titles, IDs and dates, newest first |
| `get_page_content(page_id, max_chars=40000, image_text=true)` | One page as Markdown. Typed text comes first; text recognised in images (e.g. lecture-slide printouts) is listed at the end, or left out with `image_text=false` |
| `search_pages(query, limit=25)` | Pages whose **title** contains all the words (checks the 1,000 most recently modified pages) |
| `get_page_structure(page_id)` | A page's paragraphs, headings, lists, tables and images with their element IDs |

These tools are always available and only read. The editing tools are below.

## Editing (optional)

Turn editing on with `ONENOTE_ALLOW_WRITE=1` (or `"allow_write": true` in `config.json`).

| Tool | What it does |
| --- | --- |
| `copy_page(page_id, target_section_id, new_title, allow_duplicate)` | Duplicate a page with everything on it (slide printouts, images, attachments, ink). The first copy of a page becomes its working copy. Big pages copy in the background |
| `get_copy_status(job_id, wait_seconds)` | Progress of a background copy or restore, then the new page |
| `create_page(section_id, title, content)` | New page from Markdown |
| `create_section(notebook_id, name)` | New section |
| `add_content(page_id, content, position, target_id)` | Add Markdown at the end or start of the page, or before/after an element |
| `replace_content(page_id, target_id, content)` | Rewrite or reformat one paragraph, heading, list, table or image |
| `set_page_title(page_id, title)` | Rename a page |
| `insert_image(page_id, source, alt_text, width, position, target_id)` | Image from a local file (png/jpg/gif/bmp/tiff, ≤ 25 MB) or a URL |
| `attach_file(page_id, path, position, target_id)` | Attach a local file (≤ 25 MB) |
| `restore_previous_version(page_id)` | Undo the most recent edit (see below) |
| `delete_page(page_id)` | Only when `ONENOTE_ALLOW_DELETE=1` is also set |

Content is Markdown: headings, bold/italic/strikethrough, lists, `- [ ]` / `- [x]` to-dos
(they become OneNote to-do tags), tables, links, and `![alt](path or URL)` images. For
formatting Markdown can't express, a little inline HTML is allowed: `<u>`, `<sub>`, `<sup>`
and `<span style="…">` with colour, highlight (`background-color`), font, size, weight, style
or decoration. Any other HTML is stripped before it reaches OneNote.

**How your originals are protected**

- **Originals are never edited.** Claude is told to call `copy_page` before working on an
  existing page, and then edits the copy. If it edits an original's ID anyway, the server
  first copies the page into the same section as **"<title> [edited]"**, and that edit and
  every later one go to the copy. Your original stays exactly as it was, next to the copy, for comparison. The copy is
  made by OneNote itself (`copyToSection`), so ink, drawings and attachments come along. If
  that feature isn't available to your account, the server rebuilds the copy from the page's
  HTML instead, re-uploading every image and attached file (including PDF printouts), and
  says so; that version loses ink.
- **Copies run as background jobs.** `copy_page` returns within about 8 s; if the copy
  isn't done by then it gives a job ID, and `get_copy_status` reports progress (e.g.
  "copying images and files (16/59)"). Without a job ID, `get_copy_status` lists recent jobs.
  The job keeps going if the chat is stopped, and calling `copy_page` again for the same page
  joins the running job or returns the existing copy instead of making a duplicate.
- **How a copy is made.** OneNote's exact copy is tried first (90 s limit). Personal
  (consumer) notebooks don't support it (Graph answers HTTP 501); the server remembers that
  and goes straight to rebuilding: it creates the copy at once with "[Copying image 3 of
  59…]" placeholders, then downloads images and files six at a time and uploads them in
  batches of eight, replacing the placeholders. An image that genuinely can't be copied is
  marked on the page.
  Each file's type is taken from its own bytes (and OneNote's records), not from the
  download's generic header, and anything that isn't real image data isn't uploaded. When
  the copy is done, the server downloads sample images back from it and checks for the
  attached files; if something's wrong, the result starts with a ⚠️ warning.
- **Unfinished copies are resumable.** If the connection drops, Claude Desktop restarts or
  the job is otherwise cut off, the copy keeps its remaining placeholders and is flagged as
  UNFINISHED in `list_pages`. Calling `copy_page` for the same original (or
  `get_copy_status` for the job) finishes it from where it stopped, by matching the
  placeholders to the original's images. Edits wait until a copy is complete.
- **The original→copy link is remembered** in `edit_state.json` in the app directory, so
  later chats keep editing the same copy. `list_pages` marks originals, working copies and
  pages Claude created. If you delete the copy in OneNote, the next edit
  makes a fresh one. Pages Claude creates itself are edited directly; they need no copy.
- **One snapshot per page.** Before each edit, the server saves the page's HTML to
  `snapshots/` in the app directory, overwriting the previous snapshot for that page. Only
  text and markup are saved, not image data, so the files are small. `restore_previous_version`
  rebuilds the page as it was before the last edit, as a new page next to the broken one
  (delete that one yourself once you've checked). It's a one-step undo, not a history.
- Claude Desktop asks before each tool call that changes something, unless you've chosen
  "always allow".

**Turning it on**

1. In Azure → your app → **API permissions** → Add a permission → Microsoft Graph →
   Delegated → **`Notes.ReadWrite`**. (You can then remove `Notes.Read`.)
2. Sign in again so you consent to the new permission:
   ```powershell
   $env:ONENOTE_CLIENT_ID="<id>"; $env:ONENOTE_ALLOW_WRITE="1"
   uv run onenote-mcp logout
   uv run onenote-mcp login
   uv run onenote-mcp status      # Mode: read/write (scope Notes.ReadWrite)
   ```
3. Run `uv run onenote-mcp desktop-config --write` in the same window (it copies
   `ONENOTE_ALLOW_WRITE` into the config) and restart Claude Desktop.

**Editing limits**

- Handwriting/ink can't be edited, and positions on the page can't be changed.
- Claude can't see what you've selected in OneNote; point at text by page and wording, e.g.
  "rewrite the paragraph starting 'Budget for…' on my Trip page".
- Images pasted into the Claude chat can't be passed to OneNote; give a file path or URL.
- Edits show up after OneNote syncs, usually within a minute.
- Sections and notebooks can't be renamed or deleted, and subpages can't be created.

## 1. Register an app in Azure (one-off, manual)

1. Go to the [Azure portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** →
   **New registration**.
   - Name: anything, e.g. `OneNote MCP (local)`.
   - **Supported account types:** *Personal Microsoft accounts only* (or *Accounts in any
     organizational directory and personal Microsoft accounts* if you also want work/school).
   - Redirect URI: leave empty (device code flow doesn't need one).
2. **Authentication** → *Advanced settings* → **Allow public client flows: Yes** → Save.
3. **API permissions** → Add a permission → Microsoft Graph → **Delegated** → `Notes.Read` → Add.
   (Remove `User.Read` if you like; it isn't used.) You consent to it yourself at first sign-in.
4. Copy the **Application (client) ID** from the Overview page.

If the portal says you can't register apps outside a directory, sign in with an Azure account
that has a (free) default directory; the app can still be used by your personal account.

### Which authority?

For a personal Microsoft account (outlook.com, hotmail.com, live.com…) use **`consumers`**,
which is the default. It works with both account-type options above. `common` also accepts
personal accounts, but only if the registration includes work/school accounts. It fails for
personal-only registrations, so stick with `consumers` unless you need both. Set
`ONENOTE_AUTHORITY` to override (`consumers`, `common`, `organizations`, or a tenant ID).

## 2. Install

Requires Python 3.11+ and [uv](https://docs.astral.sh/uv/).

```sh
cd onenote-mcp
uv sync
```

## 3. Configure the client ID

Either set an environment variable:

```sh
export ONENOTE_CLIENT_ID=<your-application-client-id>
```

or create `config.json` in the app directory (see below):

```json
{ "client_id": "<your-application-client-id>" }
```

| Setting | Env var | `config.json` key | Default |
| --- | --- | --- | --- |
| Client ID | `ONENOTE_CLIENT_ID` | `client_id` | *(required)* |
| Authority | `ONENOTE_AUTHORITY` | `authority` | `consumers` |
| Token cache file | `ONENOTE_TOKEN_CACHE` | `token_cache` | `<app dir>/token_cache.json` |
| App directory | `ONENOTE_MCP_HOME` | – | see below |
| Log level | `ONENOTE_LOG_LEVEL` | – | `INFO` |
| Allow editing | `ONENOTE_ALLOW_WRITE` | `allow_write` | off |
| Allow page deletion | `ONENOTE_ALLOW_DELETE` | `allow_delete` | off (needs editing on) |

App directory: macOS `~/Library/Application Support/onenote-mcp/`, Windows
`%LOCALAPPDATA%\onenote-mcp\`, Linux `~/.config/onenote-mcp/`. It's outside the repo, so the
token cache can't be committed by accident (`.gitignore` also excludes it). On macOS and Linux
the cache file is `0600` in a `0700` directory. On Windows it relies on your user profile's
default permissions.

## 4. Sign in

```sh
uv run onenote-mcp login     # prints a URL and code; finish in your browser
uv run onenote-mcp status    # shows config and who is signed in
uv run onenote-mcp logout    # deletes the cached tokens
```

You can skip `login`: if the server isn't signed in, the first tool call returns the URL and
code to Claude, which shows them to you. Sign in, tell Claude you're done, and it retries.

## 5. Add to Claude Desktop

```sh
uv run onenote-mcp desktop-config           # print the JSON entry
uv run onenote-mcp desktop-config --write   # merge it into claude_desktop_config.json (backs up first)
```

The config file is at `~/Library/Application Support/Claude/claude_desktop_config.json`
(macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows). The entry looks like this:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "/path/to/onenote-mcp/.venv/bin/python",
      "args": ["-m", "onenote_mcp"],
      "env": { "ONENOTE_CLIENT_ID": "<id>", "ONENOTE_AUTHORITY": "consumers" }
    }
  }
}
```

It points straight at the project's virtualenv Python, so Claude Desktop doesn't need `uv` on
its `PATH`. Run `desktop-config` with `ONENOTE_CLIENT_ID` set (or `config.json` present), then
**restart Claude Desktop**. Server logs: `~/Library/Logs/Claude/mcp-server-onenote.log` (macOS)
or `%APPDATA%\Claude\logs\` (Windows).

## Use from Claude Code

```sh
claude mcp add --transport stdio --env ONENOTE_CLIENT_ID=<id> onenote -- uv --directory /path/to/onenote-mcp run onenote-mcp
```

Then run `/mcp` in Claude Code to check that it's connected, and ask something like "list my
OneNote notebooks". You can also use the MCP Inspector:
`npx @modelcontextprotocol/inspector uv --directory /path/to/onenote-mcp run onenote-mcp`.

## Development

```sh
uv run pytest
```

The tests mock Graph and MSAL, so they need no account or network.

Layout: `config.py` (settings), `auth.py` (MSAL + token cache), `graph.py` (HTTP, paging,
retries, writes), `convert.py` (HTML → Markdown, page structure), `compose.py` (Markdown →
OneNote HTML, images/files), `editing.py` (copy on first edit, snapshots, restore),
`state.py` (local edit state), `server.py` (MCP tools), `cli.py` (commands).

The server uses MCP Python SDK 2.x. FastMCP was renamed `MCPServer` in that release
(`from mcp.server.mcpserver import MCPServer`), with the same decorator-based API.

## Logs

Every Graph request is logged with its status and timing (writes and anything slower than
3 s at INFO), along with each copy job's steps, including every image's size and detected
type. The server writes its own log file, `server.log` in the app directory
(`%LOCALAPPDATA%\onenote-mcp\server.log` on Windows,
`~/Library/Application Support/onenote-mcp/server.log` on macOS; `onenote-mcp status`
prints the path), and also logs to stderr, which Claude Desktop may keep in its own logs
folder. Logs contain page IDs and titles,
not page content. Set `ONENOTE_LOG_LEVEL=DEBUG` to log every request.

## Errors and throttling

- **429 / 502 / 503 / 504:** the server retries reads up to 3 times. It honours
  `Retry-After`, with each wait capped at 30 s. Edits are retried only on 429, since a server
  error during an edit may have half-applied it and a retry could duplicate content. If Graph is still throttling after that, Claude gets a message
  telling it to wait and slow down.
- **401:** the server forces one token refresh and retries. If that fails, run
  `onenote-mcp logout` and sign in again.
- **403 / 404:** Claude gets messages that point at missing permissions or stale IDs. Graph's
  own error code is included.

## Known limitations

- **Handwriting, ink and drawings** aren't returned as text by Graph. Pages that contain only
  ink say so instead of coming back blank.
- **Images and attachments** appear as `[Image: alt text]` and `[Attachment: name]`
  placeholders. Their content isn't downloaded.
- Formatting is simplified. Headings, lists, tables, links, bold/italic and to-do checkboxes
  come through; colours and the positions of items on the page don't.
- OneNote's Graph API is slow and throttles. Read the pages you need rather than everything.
- `search_pages` checks titles only, because Graph has no full-text OneNote search for
  personal accounts.
- This server is local only, so it doesn't work on claude.ai web or mobile.

## Possible later additions

- A hosted remote connector (HTTPS + proper OAuth) for claude.ai web/mobile.
