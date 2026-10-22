# onenote-mcp

A local, **read-only** [MCP](https://modelcontextprotocol.io) server that lets Claude read your
Microsoft OneNote notebooks through Microsoft Graph. It runs on your machine over stdio, for
Claude Desktop and Claude Code.

- Signs in with the MSAL **device code flow** using the delegated `Notes.Read` scope, so it can
  only see the notebooks of the account that signs in.
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
| `get_page_content(page_id, max_chars=40000)` | One page as Markdown |
| `search_pages(query, limit=25)` | Pages whose **title** contains all the words (checks the 1,000 most recently modified pages) |

All tools are marked read-only.

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
retries), `convert.py` (HTML → Markdown), `server.py` (MCP tools), `cli.py` (commands).

The server uses MCP Python SDK 2.x. FastMCP was renamed `MCPServer` in that release
(`from mcp.server.mcpserver import MCPServer`), with the same decorator-based API.

## Errors and throttling

- **429 / 502 / 503 / 504:** the server retries up to 3 times. It honours `Retry-After`, with
  each wait capped at 30 s. If Graph is still throttling after that, Claude gets a message
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

- Write tools (`create_page`, `append_to_page`) with `Notes.ReadWrite`.
- A hosted remote connector (HTTPS + proper OAuth) for claude.ai web/mobile.
